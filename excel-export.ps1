[CmdletBinding()]
param(
    [string] $Path,
    [ValidateSet('tsv', 'pdf', 'all')]
    [string] $Format,
    [string] $OutputDirectory,
    [switch] $NoRun
)

Set-StrictMode -Version Latest

function Test-SupportedInputFile {
    param([Parameter(Mandatory)][string] $InputPath)

    return [System.IO.Path]::GetExtension($InputPath) -in '.xlsx', '.xlsm'
}

function ConvertTo-SafeFileName {
    param([Parameter(Mandatory)][string] $Name)

    return ($Name -replace '[\\/:*?"<>|]', '_')
}

function Get-UniqueSheetFileStem {
    param(
        [Parameter(Mandatory)][string] $SheetName,
        [Parameter(Mandatory)][hashtable] $UsedNames
    )

    $baseName = ConvertTo-SafeFileName $SheetName
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = 'Sheet'
    }

    $candidate = $baseName
    $number = 2
    while ($UsedNames.ContainsKey($candidate.ToLowerInvariant())) {
        $candidate = '{0}_{1}' -f $baseName, $number
        $number++
    }

    $UsedNames[$candidate.ToLowerInvariant()] = $true
    return $candidate
}

function Resolve-ExportRequest {
    param(
        [Parameter(Mandatory)][string] $InputPath,
        [Parameter(Mandatory)][ValidateSet('tsv', 'pdf', 'all')][string] $RequestedFormat,
        [string] $RequestedOutputDirectory
    )

    if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
        throw [System.ArgumentException]::new("File not found: $InputPath")
    }
    if (-not (Test-SupportedInputFile $InputPath)) {
        $extension = [System.IO.Path]::GetExtension($InputPath)
        throw [System.ArgumentException]::new("Unsupported file type: $extension")
    }

    $fullInputPath = [System.IO.Path]::GetFullPath($InputPath)
    if ([string]::IsNullOrWhiteSpace($RequestedOutputDirectory)) {
        $fullOutputDirectory = [System.IO.Path]::GetDirectoryName($fullInputPath)
    }
    else {
        $fullOutputDirectory = [System.IO.Path]::GetFullPath($RequestedOutputDirectory)
        if (-not (Test-Path -LiteralPath $fullOutputDirectory)) {
            New-Item -ItemType Directory -Path $fullOutputDirectory -Force -ErrorAction Stop | Out-Null
        }
        if (-not (Test-Path -LiteralPath $fullOutputDirectory -PathType Container)) {
            throw [System.ArgumentException]::new("Output path is not a directory: $fullOutputDirectory")
        }
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fullInputPath)
    return [pscustomobject]@{
        InputPath       = $fullInputPath
        Format          = $RequestedFormat
        OutputDirectory = $fullOutputDirectory
        TsvDirectory    = Join-Path $fullOutputDirectory $baseName
        PdfPath         = Join-Path $fullOutputDirectory ($baseName + '.pdf')
    }
}

function Write-Info {
    param([Parameter(Mandatory)][string] $Message)

    Write-Host "[INFO] $Message"
}

function Release-ComObject {
    param($ComObject)

    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Export-WorksheetsToTsv {
    param(
        [Parameter(Mandatory)] $Workbook,
        [Parameter(Mandatory)] $Excel,
        [Parameter(Mandatory)][string] $TsvDirectory
    )

    New-Item -ItemType Directory -Path $TsvDirectory -Force -ErrorAction Stop | Out-Null
    $usedNames = @{}

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        $temporaryWorkbook = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            $fileStem = Get-UniqueSheetFileStem $worksheet.Name $usedNames
            $tsvPath = Join-Path $TsvDirectory ($fileStem + '.tsv')

            $worksheet.Copy()
            $temporaryWorkbook = $Excel.ActiveWorkbook
            $temporaryWorkbook.SaveAs($tsvPath, 20)
            Write-Info "Exporting TSV: $($fileStem).tsv"
        }
        finally {
            if ($temporaryWorkbook) {
                $temporaryWorkbook.Close($false)
            }
            Release-ComObject $temporaryWorkbook
            Release-ComObject $worksheet
        }
    }
}

function Export-Workbook {
    param([Parameter(Mandatory)] $Request)

    $excel = $null
    $workbook = $null
    try {
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        }
        catch {
            throw [System.InvalidOperationException]::new(
                'Microsoft Excel is not installed or Excel COM is unavailable.',
                $_.Exception
            )
        }

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AutomationSecurity = 3 } catch { Write-Verbose 'Could not set AutomationSecurity.' }
        $workbook = $excel.Workbooks.Open($Request.InputPath, 0, $true)

        if ($Request.Format -in 'tsv', 'all') {
            Export-WorksheetsToTsv $workbook $excel $Request.TsvDirectory
        }
        if ($Request.Format -in 'pdf', 'all') {
            $workbook.ExportAsFixedFormat(0, $Request.PdfPath)
            Write-Info "Exporting PDF: $([System.IO.Path]::GetFileName($Request.PdfPath))"
        }
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if ($NoRun) {
    return
}

try {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw [System.ArgumentException]::new('Input path is required.')
    }
    if ([string]::IsNullOrWhiteSpace($Format)) {
        throw [System.ArgumentException]::new('Format is required.')
    }

    $request = Resolve-ExportRequest $Path $Format $OutputDirectory
    Write-Info "Opening: $([System.IO.Path]::GetFileName($request.InputPath))"
    Export-Workbook $request
    Write-Info 'Done.'
    exit 0
}
catch [System.ArgumentException] {
    Write-Error "ERROR: $($_.Exception.Message)"
    exit 2
}
catch [System.InvalidOperationException] {
    Write-Error "ERROR: $($_.Exception.Message)"
    exit 3
}
catch {
    Write-Error "ERROR: Failed to convert workbook. $($_.Exception.Message)"
    exit 4
}
