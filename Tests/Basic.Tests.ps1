Describe 'Basic Tests' {
    It 'should return true for true' {
        $result = $true
        $result | Should -Be $true
    }

    It 'should return false for false' {
        $result = $false
        $result | Should -Be $false
    }
}