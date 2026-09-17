# Excel Export Helper MVP Design

## Scope

Implement the `SPEC.md` MVP as one PowerShell script, `excel-export.ps1`, with a README and Pester tests. The tool runs locally on Windows and automates an already-installed desktop Microsoft Excel through COM. It never contacts Google, Microsoft, or other external services.

## Command interface

`excel-export.ps1` accepts:

- `-Path`: a required path to one `.xlsx` or `.xlsm` file.
- `-Format`: a required `tsv`, `pdf`, or `all` value.
- `-OutputDirectory`: an optional existing-or-creatable directory. The default is the input file's directory.

Invalid input paths and extensions end with exit code 2. Excel COM startup failures end with 3; export failures end with 4; all other failures end with 1.

## Conversion flow

1. Resolve and validate the input file and output directory before starting Excel.
2. Create one invisible Excel application instance with alerts disabled and automation security set to force-disable macros when available.
3. Open the workbook read-only, without updating links or prompting for them.
4. For TSV output, create a directory named after the input file's base name under the output directory. Copy each worksheet to an isolated temporary workbook, save it using Excel's tab-delimited format, then close that temporary workbook without saving. Sheet names are made safe for Windows filenames and collisions receive `_2`, `_3`, and so on.
5. For PDF output, call `ExportAsFixedFormat` on the original workbook, producing one PDF whose name is the input file's base name.
6. In `finally`, close the source and temporary workbooks without saving, quit Excel, release COM objects, and collect COM wrappers. The source workbook is never saved or modified.

## Errors and logging

The script writes concise `[INFO]` progress messages to the host. It emits an `ERROR:` message to the error stream and terminates with the documented exit code for expected failure classes. `-Verbose` is used only for diagnostics such as resolved paths and COM cleanup details.

## Files

- `excel-export.ps1`: parameter handling, conversion orchestration, helpers, cleanup, and exit handling.
- `README.md`: prerequisites, usage examples, output layout, safety guarantees, and troubleshooting.
- `tests/ExcelExportHelper.Tests.ps1`: Pester tests for extension validation, output-path construction, safe unique sheet filenames, and argument-related helper behavior. COM conversion itself is manually verified on a Windows machine with Excel.

## Test strategy

Keep pure helper functions testable by loading the script in a no-run mode. Tests do not require Excel. A manual verification checklist in the README covers TSV/PDF conversion, macro-enabled inputs, cleanup after an intentional failure, and confirming the source file remains unchanged.
