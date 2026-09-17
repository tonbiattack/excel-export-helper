$scriptPath = Join-Path $PSScriptRoot '..\excel-export.ps1'

Describe 'Excel Export Helper validation helpers' {
    BeforeAll {
        . $scriptPath -NoRun
    }

    It 'accepts xlsx input files case-insensitively' {
        Test-SupportedInputFile 'C:\work\report.XLSX' | Should Be $true
    }

    It 'rejects unsupported input file types' {
        Test-SupportedInputFile 'C:\work\report.csv' | Should Be $false
    }

    It 'replaces invalid Windows filename characters' {
        ConvertTo-SafeFileName 'Sales:West/2026?' | Should Be 'Sales_West_2026_'
    }

    It 'adds a numeric suffix for duplicate sheet names' {
        $usedNames = @{}
        Get-UniqueSheetFileStem 'Sheet' $usedNames | Should Be 'Sheet'
        Get-UniqueSheetFileStem 'Sheet' $usedNames | Should Be 'Sheet_2'
    }

    It 'builds default output paths from an input workbook' {
        $source = Join-Path $TestDrive 'report.xlsx'
        New-Item -ItemType File -Path $source -Force | Out-Null

        $request = Resolve-ExportRequest $source 'all' $null

        $request.OutputDirectory | Should Be ([System.IO.Path]::GetDirectoryName($source))
        $request.TsvDirectory | Should Be (Join-Path $TestDrive 'report')
        $request.PdfPath | Should Be (Join-Path $TestDrive 'report.pdf')
    }

    It 'creates a requested output directory' {
        $source = Join-Path $TestDrive 'report.xlsx'
        New-Item -ItemType File -Path $source -Force | Out-Null
        $output = Join-Path $TestDrive 'exports'

        $request = Resolve-ExportRequest $source 'pdf' $output

        Test-Path -LiteralPath $output -PathType Container | Should Be $true
        $request.OutputDirectory | Should Be $output
    }

    It 'rejects a missing input workbook' {
        $missing = Join-Path $TestDrive (([guid]::NewGuid().ToString()) + '.xlsx')
        $threw = $false
        try {
            Resolve-ExportRequest $missing 'pdf' $null
        }
        catch [System.ArgumentException] {
            $threw = $true
        }
        $threw | Should Be $true
    }

    It 'exposes the COM cleanup helper' {
        Get-Command Release-ComObject -ErrorAction Stop | Should Not BeNullOrEmpty
    }
}
