Describe "empty" {
    Context "pbit" {
        It "can load" {
            Invoke-Expression -Command @"
                & .\src\test_pbi_reports.ps1 tests\reports\empty.pbix
"@
            $? | Should -Be $true
        }
    }

    Context "pbix" {
        It "can load" {
            powershell -Command "& .\src\test_pbi_reports.ps1 tests\reports\empty.pbix"
            $? | Should -Be $true
        }
    }
}

Describe "errorneous_column" {
    Context "pbit" {
        It "cannot load" {
            powershell -Command "& .\src\test_pbi_reports.ps1 tests\reports\errorneous_column.pbix"
            $? | Should -Not -Be $true
        }
    }

    Context "pbix" {
        It "cannot load" {
            powershell -Command "& .\src\test_pbi_reports.ps1 tests\reports\errorneous_column.pbix"
            $? | Should -Not -Be $true
        }
    }
}
