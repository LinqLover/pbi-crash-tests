// Test Suite for Power BI crash tests
// Copyright © Christoph Thiede 2020.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using HWND = System.IntPtr;

using ImageMagick;


namespace PbiCrashTests
{
    /// <summary>
    /// Represents a test case for a Power BI report file.
    /// Purpose of the test will be to try to open and load the report
    /// and detect if any errors or deadlocks occur while doing so.
    /// </summary>
    public class PbiReportTestCase {
        /// <summary>
        /// The path to the report template file.
        /// </summary>
        public string Report { get; }

        public string Name => Path.GetFileName(Report);

        /// <summary>
        /// The path to the PBI Desktop executable file.
        /// </summary>
        public string Desktop { get; }

        /// <summary>
        /// The delay Power BI is granted to hang before the first window is opened. If exceeded,
        /// the test will timeout.
        /// </summary>
        public TimeSpan PreLoadDelay { get; }

        /// <summary>
        /// The delay Power BI is granted to hang after all data sources have been adapted
        /// before final error search starts.
        /// </summary>
        public TimeSpan LoadDelay { get; }

        public bool HasPassed { get; private set; }

        public bool HasFailed { get; private set; }

        public string ResultReason { get; private set; }

        protected static Dictionary<string, Bitmap> FailureIcons = Directory.GetFiles(
				Path.Combine(Directory.GetCurrentDirectory(), "data/failure_icons"),
				"*.bmp"
            ).ToDictionary(
                file => Path.GetFileNameWithoutExtension(file),
                file => new Bitmap(file)
            );

        protected Process Process { get; private set; }
        
        protected PbiProcessSnapshot Snapshot { get; private set; }
        
        public PbiReportTestCase(
                string report, string desktop,
                TimeSpan preLoadDelay, TimeSpan loadDelay) {
            Report = report;
            Desktop = desktop;
            (PreLoadDelay, LoadDelay) = (preLoadDelay, loadDelay);
        }

        /// <summary>
        /// Start this test.
        /// </summary>
        public void Start() => _logExceptions(() => {
            Process = new Process {
                StartInfo = {
                    FileName = Desktop,
                    Arguments = $"\"{Report}\""
                }
            };
            Process.Start();

            System.Threading.Thread.Sleep(LoadDelay);

            Snapshot = new PbiProcessSnapshot(Process);
        });

        /// <summary>
        /// Try to find indications for this test either having passed or failed.
        /// If an indication is found, HasPassed or HasFailed will be set accordingly.
        /// </summary>
        public void Check() => _logExceptions(() => {
            void handleFail(string reason) {
                HasFailed = true;
                ResultReason = reason;
            }
            _check(
                handlePass: () => {
                    System.Threading.Thread.Sleep(LoadDelay);
                    _check(
                        handlePass: () => HasPassed = true,
                        handleFail: handleFail
                    );
                },
                handleFail: handleFail
            );
        });

        /// <summary>
        /// Abort this test.
        /// </summary>
        public void Stop() => _logExceptions(() => {
            if (!Process.HasExited)
                Process.Kill();
        });

        /// <summary>
        /// Save the results of this test into <paramref name="path"/>.
        /// </summary>
        public void SaveResults(string path) => _logExceptions(() => {
            Snapshot.SaveArtifacts(Path.Combine(path, Path.GetFileNameWithoutExtension(Report)));
        });

        public override string ToString() {
            return $"PBI Report Test: {Name}";
        }

        private void _check(Action handlePass, Action<string> handleFail) {
            if (Process.HasExited) {
                handleFail("Power BI has unexpectedly terminated");
                return;
            }

            Snapshot.Update();

            var windows = Snapshot.Windows.ToList();
            if (windows.Count == 1 && windows[0].Title.EndsWith(" - Power BI Desktop")) {
                handlePass();
                return;
            }
            if (!windows.Any(window => string.IsNullOrWhiteSpace(window.Title))) {
                handleFail("Power BI did not open any valid window");
                return;
            }
            foreach (var (failure, icon) in FailureIcons.Select(kvp => (kvp.Key, kvp.Value)))
                foreach (var (window, index) in windows.Select((window, index) => (window, index))) {
                    Console.WriteLine($"(Checking window {index} against icon {failure})");
                    if (window.DisplaysIcon(icon, out var similarity)) {
                        handleFail($"Power BI showed an error in window {index} " +
                                   $"while loading the report: {failure} (similarity={similarity})");
                        return;
                    }
                }
        }

        /// <summary>
        /// Helper function that catches any exception and writes the stack trace to console
        /// before re-throwing the exception. This is helpful when running a C# script from
        /// PowerShell.
        /// </summary>
        /// <param name="action">The actual action to wrap.</param>
        private static void _logExceptions(Action action) {
            try {
                action();
            } catch (Exception ex) {
                Console.WriteLine(ex);
                throw;
            }
        }
    }

    /// <summary>
    /// Represents the state of an observed Power BI process at a certain point in time.
    /// </summary>
    public class PbiProcessSnapshot {
        public PbiProcessSnapshot(Process process) {
            Process = process;
        }

        public Process Process { get; }

        /// <summary>
        /// All windows that are currently opened by the process.
        /// </summary>
        public IEnumerable<PbiWindowSnapshot> Windows => _windows;
        
        private List<PbiWindowSnapshot> _windows = new List<PbiWindowSnapshot>();

        /// <summary>
        /// Update this snapshot.
        /// </summary>
        public void Update() {
            _windows = (
                from window in CollectWindows()
                where window.IsVisible
                where !window.Bounds.Equals(default(PbiWindowSnapshot.RECT))
                select window
            ).ToList();
        }

        /// <summary>
        /// Save all artifacts of this snapshot into <paramref name="path"/>.
        /// In detail, these are screenshots of all open windows.
        /// </summary>
        /// <param name="path"></param>
        public void SaveArtifacts(string path) {
            Directory.CreateDirectory(path);

            foreach (var (window, index) in Windows.Select((window, index) => (window, index))) {
                if (!window.IsVisible) continue;

                var label = new StringBuilder("window");
                label.AppendFormat("_{0}", index);
                if (!string.IsNullOrWhiteSpace(window.Title))
                    label.AppendFormat("_{0}", window.Title);
                label.Append(".png");

                var screenshot = window.Screenshot;

                screenshot?.Save(Path.Combine(path, label.ToString()));
            }
        }

        protected IEnumerable<PbiWindowSnapshot> CollectWindows() {
            foreach (var window in GetRootWindows(Process.Id))
                yield return PbiWindowSnapshot.Create(window);
        }

        private static IEnumerable<HWND> GetRootWindows(int pid)
        {
            var windows = GetChildWindows(HWND.Zero);
            foreach (var child in windows)
            {
                GetWindowThreadProcessId(child, out var lpdwProcessId);
                if (lpdwProcessId == pid)
                    yield return child;
            }
        }

        private static IEnumerable<HWND> GetChildWindows(HWND parent)
        {
            var result = new List<HWND>();
            var listHandle = GCHandle.Alloc(result);
            try
            {
                var childProc = new Win32Callback(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }

        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            var gch = GCHandle.FromIntPtr(pointer);
            var list = (List<IntPtr>)gch.Target;
            list.Add(handle);
            return true;
        }

#region DllImports
        private delegate bool Win32Callback(HWND hwnd, IntPtr lParam);

        [DllImport(DllImportNames.USER32)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(HWND parentHandle, Win32Callback callback, IntPtr lParam);

        [DllImport(DllImportNames.USER32)]
        private static extern uint GetWindowThreadProcessId(HWND hwnd, out uint lpdwProcessId);
#endregion DllImports
    }

    /// <summary>
    /// Represents the state of an observed Power BI window at a certain point in time.
    /// </summary>
    public class PbiWindowSnapshot {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public PbiWindowSnapshot(HWND hwnd) : this() {
            Hwnd = hwnd;
        }

        private PbiWindowSnapshot() {
            _screenshot = new Lazy<Bitmap>(RecordScreenshot);
        }

        /// <summary>
        /// The window handle (hWnd) of this window.
        /// </summary>
        public HWND Hwnd { get; }

        public string Title { get; private set; }

        public bool IsVisible { get; private set; }

        public RECT Bounds { get; private set; }

        /// <summary>
        /// A screenshot of this window.
        /// </summary>
        /// <remarks>
        /// Will be generated lazilly.
        /// </remarks>
        public Bitmap Screenshot => _screenshot.Value;

        protected static double IconSimilarityThreshold = 0.1;

        private readonly Lazy<Bitmap> _screenshot;

        public static PbiWindowSnapshot Create(HWND hwnd) {
            var window = new PbiWindowSnapshot(hwnd);
            window.Update();
            return window;
        }

        /// <summary>
        /// Update this snapshot.
        /// </summary>
        public void Update() {
            Title = GetWindowTitle();
            IsVisible = IsWindowVisible(Hwnd);
            Bounds = GetWindowBounds();
        }

        /// <summary>
        /// Tests whether this window displays the specified icon.
        /// </summary>
        public bool DisplaysIcon(Bitmap icon, out double similarity) {
            var screenshot = Screenshot;
            if (screenshot is null) {
                similarity = default(double);
                return false;
            }

            var magickIcon = CreateMagickImage(icon);
            var magickScreenshot = CreateMagickImage(screenshot);
            var result = magickScreenshot.SubImageSearch(magickIcon);
            return (similarity = result.SimilarityMetric) <= IconSimilarityThreshold;
        }

        private RECT GetWindowBounds() {
            if (!GetWindowRect(new HandleRef(null, Hwnd), out var rect))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            return rect;
        }

        private string GetWindowTitle()
        {
            var length = GetWindowTextLength(Hwnd);
            var title = new StringBuilder(length);
            GetWindowText(Hwnd, title, length + 1);
            return title.ToString();
        }

        private Bitmap RecordScreenshot() {
            var bmp = new Bitmap(Bounds.Right - Bounds.Left, Bounds.Bottom - Bounds.Top, PixelFormat.Format32bppArgb);
            using (var gfxBmp = Graphics.FromImage(bmp)) {
                IntPtr hdcBitmap;
                try
                {
                    hdcBitmap = gfxBmp.GetHdc();
                }
                catch
                {
                    return null;
                }
                bool succeeded = PrintWindow(Hwnd, hdcBitmap, 0);
                gfxBmp.ReleaseHdc(hdcBitmap);
                if (!succeeded)
                {
                    return null;
                }
                var hRgn = CreateRectRgn(0, 0, 0, 0);
                GetWindowRgn(Hwnd, hRgn);
                var region = Region.FromHrgn(hRgn);
                if (!region.IsEmpty(gfxBmp))
                {
                    gfxBmp.ExcludeClip(region);
                    gfxBmp.Clear(Color.Transparent);
                }
            }
            return bmp;
        }

        private static MagickImage CreateMagickImage(Bitmap bitmap) {
            var magickImage = new MagickImage();
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Bmp);

                stream.Position = 0;
                magickImage.Read(stream);
            }

            const double scaleFactor = 1 / 3d; // For performance
            magickImage.Resize(
                (int)Math.Round(magickImage.Width * scaleFactor),
                (int)Math.Round(magickImage.Height * scaleFactor)
            );
            return magickImage;
        }

#region DllImports

        [DllImport(DllImportNames.GDI32)]
        private static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect);

        [DllImport(DllImportNames.USER32)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(HandleRef hwnd, out RECT lpRect);

        [DllImport(DllImportNames.USER32)]
        private static extern int GetWindowRgn(HWND hWnd, IntPtr hRgn);

        [DllImport(DllImportNames.USER32, CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(HWND hwnd, StringBuilder lpString, int nMaxCount);

        [DllImport(DllImportNames.USER32, SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(HWND hwnd);

        [DllImport(DllImportNames.USER32)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport(DllImportNames.USER32, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PrintWindow(HWND hwnd, IntPtr hDC, uint nFlags);
#endregion DllImports
    }

    internal static class DllImportNames {
        public const string GDI32 = "gdi32.dll";
        public const string USER32 = "user32.dll";
    }
}
