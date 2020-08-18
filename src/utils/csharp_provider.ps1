#requires -Version 5
# This piece of .NET magic returns a CodeDomProvider that allows parsing modern C# scripts.
# Run scripts/setup.ps1 to install dependencies.

$DotNetCodeDomLocation = 'nuget/microsoft.codedom.providers.dotnetcompilerplatform.2.0.1'


Add-Type -Path `
	"$DotNetCodeDomLocation\lib\net45\Microsoft.CodeDom.Providers.DotNetCompilerPlatform.dll"

# This uses the public interface ICompilerSettings instead of the private class CompilerSettings
Invoke-Expression -Command @"
class RoslynCompilerSettings : Microsoft.CodeDom.Providers.DotNetCompilerPlatform.ICompilerSettings
{
    [string] get_CompilerFullPath()
    {
        return "$DotNetCodeDomLocation\tools\RoslynLatest\csc.exe"
    }
    [int] get_CompilerServerTimeToLive()
    {
        return 10
    }
}
"@
$DotNetCodeDomProvider = [Microsoft.CodeDom.Providers.DotNetCompilerPlatform.CSharpCodeProvider]::new(
	[RoslynCompilerSettings]::new()
)

return $DotNetCodeDomProvider
