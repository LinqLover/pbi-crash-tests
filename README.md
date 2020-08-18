# Power BI Crash Tests
Load &amp; crash tests for Power BI reports (`.pbix`/`.pbit` files).

## Description

These tests were originally developed in the context of a data analysis project [1, 2] which included the making of about a dozen visualization reports using the Business Intelligence software [Power BI](https://powerbi.com/) (PBI).
These reports are stored as PBIX or PBIT files in the project repository.
To assure overall stability and quickly identify regressions in any report (for instance, caused by changes to the database schema or the unavailability of external access points), tests were written that open every single report file in Power BI Desktop and make sure that it does not show any error message or crashes during loading it.

Because Power BI Desktop does not provide any programming interface for making such assertions, the tests are implemented as UI tests using mechanisms such as counting the number of open windows (> 1 means there is any loading/error window) and scanning the screenshots of these windows for known error icons.
As a side effect, screenshots of all loaded reports including possible error messages are token to provide fast insights into the source of errors.

The tests can be run automatically as a part of a CI pipeline.
Due to the system requirements of PBI Desktop, this requires a (virtual) Windows machine and a Power BI license.

**Why use this solution?**

- Theoretically, you could also check all your reports manually, too.
  However, PBI Desktop is quite slow so this does not really scale well.
  Also, automatical tests guarantee the quality of your project whereas humans are lazy and oblivious.
- Power BI Online allows scheduling automatic report updates as well.
  However, it requires an organization account, can hardly be integrated into any VCS workflow and also might involve privacy issues.

**Keywords:** Power BI, Power BI Desktop, Power BI CI testing, Power BI acceptance tests, Power BI UI tests, Power BI crash tests.

[1] Project description (German): [PDF hosted at Museum Barberini](https://www.museum-barberini.com/site/assets/files/1080779/fg_naumann_bp_barberini_2019-20.pdf)  
[2] Press release (German): [PDF hosted at Hasso Plattner Institute](https://hpi.de/fileadmin/user_upload/hpi/veranstaltungen/2020/Bachelorpodium_2020/Pressemitteilung_BP_2020_Bachelorprojekte/Pressemitteilung_BP2020_Pressemitteilung_FN1_V2.pdf)  

## Usage

### Installation

1. Install Power BI Desktop.
2. Install PowerShell dependencies:
   ```powershell
   powershell -Command "&scripts/setup.ps1"
   ```

### Running the tests

```powershell
powershell -Command "&src/test_pbi_reports.ps1" my_report.pbix
powershell -Command "&src/test_pbi_reports.ps1" my_template.pbit
powershell -Command "&src/test_pbi_reports.ps1"  # without arguments, tests all *.pbit reports
```

See [`src/test_pbi_reports.ps1`](src/test_pbi_reports.cs) for additional arguments such as timeout values and filepaths.

## Troubleshooting

- **Too many false positives/false negatives:** Probably you need to increase the different timeout values.
  PBI can really take a long time to load certain reports, depending on the machine and internet speed (we observed template files taking up to ten minutes), and if the timeout values are too low, some loading or error windows still might be (in)visible before the last check is made.
- **Specific false positives:** Probably an error icon is displayed that is not registered under `data/failure_icons`.
  Please feel free to submit a PR in this situation!

## Limitations

- Does currently not support hi-DPI screens.  
  This could be implemented by adding/fixing some scaling logic in the `PbiWindowSnapshot` class.
- Error icons are hard-coded and there is no pictogram AI or OCR to detect arbitrary error messages, so new error sources need to be manually added.
- Only the first page of every report file can be scanned.
- No user-defined acceptance criteria tests (this probably would be out of scope).

## Development

After our data analysis project (see [Description](#description)) has ended, I wanted to backup and make available these tests.
However, I think there is still a lot of extension potential.
Ideas include deploying this repository on NuGet.org, running the metatests (see [`tests/`](tests/)) in a CI pipeline on GitHub, and probably automating the deploy process as a CD job.
You can see, I am a big fan of CI/CD stuff. :D

If you use or would like to use this project, please feel free to star the repository or leave any issue/PR and I will look forward to make the project more useful.
