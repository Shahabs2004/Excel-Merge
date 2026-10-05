# Excel Merger

A small Windows desktop application for combining multiple `.xlsx` workbooks. Excel Merger is built with C#, WPF, and ClosedXML; Microsoft Excel does not need to be installed.

> **Current behavior:** The interface offers two merge modes, but the mode check in the current code compares the selected Persian label to the English text `Merge into One Workbook`. As a result, selecting either option currently runs **Merge into One Sheet**. The two-workbook mode described below is the intended behavior of the corresponding merge routine, but the UI currently cannot select it successfully.

## Features

- Select multiple `.xlsx` files at once.
- Choose where to save the output workbook.
- Intended mode 1: copy worksheets from all source workbooks into one output workbook. Duplicate worksheet names receive a numeric suffix.
- Intended mode 2: append each non-empty worksheet's used range into one output worksheet, with a source workbook and worksheet label above each range.
- Display a progress bar and status messages during merging.
- Runs locally; selected workbook contents are processed on the computer.

## Requirements

- Windows
- .NET Framework 4.8 Developer Pack to build, and .NET Framework 4.8 Runtime to run
- Visual Studio with the **.NET desktop development** workload, or compatible MSBuild tools
- NuGet package restore access for the project dependencies

Microsoft Excel is not required. The current file picker accepts `.xlsx` files only; `.xls`, `.xlsm`, and `.csv` files are not supported by the picker.

## Get started

### Run from Visual Studio

1. Clone or download this repository.
2. Open `Excel Merge.sln` in Visual Studio.
3. Allow NuGet packages to restore. If they do not restore automatically, right-click the solution and select **Restore NuGet Packages**.
4. Select the `Excel Merge` project and run it (F5 or **Start**).

### Build from a Developer Command Prompt

```powershell
msbuild "Excel Merge.sln" /restore /p:Configuration=Release
```

The executable is normally written under `Excel Merge\bin\Release\` (the precise output path may vary by MSBuild configuration). Run `Excel Merge.exe` on a Windows machine with the .NET Framework 4.8 Runtime installed.

## Create a GitHub Release

The workflow in `.github/workflows/release.yml` builds x86 and x64 versions of the app on Windows runners and attaches both ZIP packages to one GitHub Release. It runs when a tag beginning with `v` is pushed (for example, `v1.0.0`). The workflow needs the repository's default `GITHUB_TOKEN` with `contents: write`; no additional secret is required.

Create and push a version tag to start the release:

```powershell
git tag v1.0.0
git push origin v1.0.0
```

The workflow generates release notes from GitHub's release notes API. For the example tag, the assets are named `Excel-Merger-v1.0.0-windows-x86.zip` and `Excel-Merger-v1.0.0-windows-x64.zip`. Each ZIP contains the executable and its build output dependencies; `License.txt` is excluded because it contains third-party license data.

## Using the application

1. Click **فایل های اکسل را انتخاب کنید** (select Excel files) and choose one or more `.xlsx` workbooks. The selected paths appear in the status area.
2. Select a merge mode from the dropdown.
3. Click **ادغام و ذخیره** (merge and save) and choose an output path. The suggested filename is `Merged.xlsx`.
4. Wait for the completion message, then open the saved workbook in a spreadsheet application.

The current interface is in Persian while progress and some dialogs are in English. Until the mode-label comparison is corrected, the dropdown selection does not change the operation: the app uses the one-sheet merge routine.

## Merge modes

### Merge into one workbook (intended)

Each worksheet from each input workbook is copied into the output workbook. Worksheet names are retained unless a name already exists; in that case, `_1`, `_2`, and subsequent numeric suffixes are added to make the names unique.

### Merge into one sheet (currently used by the UI)

The app creates a worksheet named `MergedSheet`. For each non-empty worksheet in each input workbook, it writes a bold source marker containing the source filename and worksheet name, then copies that worksheet's used range below the marker. A blank gap separates successive ranges. Empty worksheets are skipped.

This is a vertical append, not a row-wise table join: the app does not match columns, reconcile headers, deduplicate records, or transform data.

## Project structure

| File | Purpose |
| --- | --- |
| `Excel Merge.sln` | Visual Studio solution |
| `Excel Merge.csproj` | WPF application project targeting .NET Framework 4.8 |
| `App.xaml`, `App.xaml.cs` | WPF application startup |
| `MainWindow.xaml`, `MainWindow.xaml.cs` | User interface and workbook merge implementation |
| `packages.config` | NuGet dependency list |
| `App.config` | Runtime configuration and assembly binding redirects |
| `officeexcel_oficina_13053.ico` | Application icon |

## Dependencies

The project uses [ClosedXML](https://github.com/ClosedXML/ClosedXML) to read, copy, and write Excel workbooks. Supporting packages, including DocumentFormat.OpenXml and related libraries, are declared in `packages.config` and restored by NuGet.

## Known limitations

- Only `.xlsx` files can be selected.
- The app processes worksheets in the order returned by the workbook library and files in the order selected by the dialog.
- Merged-sheet mode copies worksheet used ranges and places them below one another; it does not combine datasets by matching columns.
- The merge is performed on a background task, but progress counts worksheets rather than rows or cells. Empty worksheets are skipped in one-sheet mode, so the displayed progress may not reach its final value until the operation completes.
- The mode-selection label mismatch described above currently makes both dropdown choices run the one-sheet routine.

## Contributing

Bug reports and pull requests are welcome. For a useful issue report, include Windows and .NET Framework versions, steps to reproduce, and a small sample workbook with sensitive information removed.

## License

No repository-level open-source license has been declared. Unless a license is added, assume the code is provided without permission to redistribute or modify it. The repository also contains a `License.txt` associated with a third-party Aspose product; it does not establish a license for this application. Do not publish third-party license credentials or signatures as part of a public release.
