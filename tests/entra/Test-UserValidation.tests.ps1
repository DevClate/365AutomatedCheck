param (
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [string] $DomainName = (Get-MgDomain | Where-Object { $_.IsDefault -eq $true }).Id,
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [string] $ExcelFilePath = $env:ExcelFilePath,
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateSet('all', 'members', 'guests', 'enabled', 'disabled')]
    [string] $UserFilter = 'all'
)

#Requires -Module Pester, ImportExcel, Microsoft.Graph

$properties = @(
    'UserPrincipalName',
    'ID', 
    'MobilePhone', 
    'JobTitle', 
    'StreetAddress', 
    'City', 
    'State', 
    'PostalCode', 
    'GivenName', 
    'Surname', 
    'DisplayName', 
    'Country', 
    'AssignedLicenses', 
    'AccountEnabled', 
    'CompanyName', 
    'Department',
    'BusinessPhones', 
    'FaxNumber',
    'UserType',
    'Mail',
    'EmployeeHireDate',
    'EmployeeID',
    'ShowInAddressList',
    'SignInActivity',
    'LastPasswordChangeDateTime',
    'MailNickName'
)

BeforeAll {
    $defaultPath = "$RootPath/assets/365ValidationParameters.xlsx"
    if ([string]::IsNullOrEmpty($ExcelFilePath) -or !(Test-Path -Path $ExcelFilePath)) {
        $ExcelFilePath = $defaultPath
        Write-Host "Default Excel file being used: $ExcelFilePath"
    } else {
        Write-Host "Using provided Excel file path: $ExcelFilePath"
    }
    
    try {
        $excelData = Import-Excel -Path $ExcelFilePath | ForEach-Object {
            $_.PSObject.Properties | ForEach-Object {
                if ([string]::IsNullOrEmpty($_.Value)) {
                    $_.Value = "NA" # Replace with your desired default value
                }
            }
            $_ # Output the modified row
        }
    }
    catch {
        Write-Host "Error importing Excel file: $_. Using default Excel file path: $defaultPath"
        $ExcelFilePath = $defaultPath
        $excelData = Import-Excel -Path $ExcelFilePath
    }
}

BeforeDiscovery {
    # Connect to the Graph SDK with the correct permissions
    Connect-MgGraph -NoWelcome -Scopes AuditLog.Read.All, Directory.Read.All

    # Set up filter query based on the user filter parameter
    $filterQuery = switch ($UserFilter) {
        'members' { "userType eq 'Member'" }
        'guests' { "userType eq 'Guest'" }
        'enabled' { "accountEnabled eq true" }
        'disabled' { "accountEnabled eq false" }
        default { '' }
    }

    # Construct the URI with the appropriate filter
    $Headers = @{ConsistencyLevel="Eventual"}  
    $Uri = "https://graph.microsoft.com/beta/users?`$count=true&`$top=999&`$select=" + ($properties -join ',')
    if ($filterQuery) {
        $Uri += "&`$filter=$filterQuery"
    }
    
    [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
    [array]$Users = $Data.Value

    If (!($Users)) {
        Write-Host "Can't find any users... exiting!" ; break
    }

    # Paginate until we have all the user accounts
    While ($Null -ne $Data.'@odata.nextLink') {
        Write-Host ("Fetching more user accounts - currently at {0}" -f $Users.count)
        $Uri = $Data.'@odata.nextLink'
        [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
        $Users = $Users + $Data.Value
    }
    Write-Host ("All available user accounts fetched ({0}) - now processing" -f $Users.count)
}

#Region User Fields
Describe "Validating User Fields" -Tag "Entra", "Users", "All" {
    Context "Job Title" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Job Title" {
            $_.JobTitle | Should -BeTrue -Because "Job Title is required for all users"
        }
    }

    Context "Street Address" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Street Address" {
            $_.StreetAddress | Should -BeTrue -Because "Street Address is required for all users"
        }
    }

    Context "City" -Tag "Basic"  -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a City" {
            $_.city | Should -BeTrue -Because "City is required for all users"
        }
    }

    Context "State" -Tag "Basic"  -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a State" {
            $_.State | Should -BeTrue -Because "State is required for all users"
        }
    }

    Context "Postal Code" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Postal Code" {
            $_.PostalCode | Should -BeTrue -Because "Postal Code is required for all users"
        }
    }
    Context "Country" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Country" {
            $_.Country | Should -BeTrue -Because "Country is required for all users"
        }
    }

    Context "Department" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Department" {
            $_.Department | Should -BeTrue -Because "Department is required for all users"
        }
    }

    Context "Office Location" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have an Office Location" {
            $_.OfficeLocation | Should -BeTrue -Because "Office Location is required for all users"
        }
    }

    Context "Assigned Licenses" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have an Assigned License" {
            $_.AssignedLicenses | Should -BeTrue -Because "Assigned Licenses are required for all users"
        }
    }

    Context "Account Enabled" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have an Account Enabled" {
            $_.AccountEnabled | Should -BeTrue -Because "Account Enabled is required for all users"
        }
    }
    
    Context "Company Name" -Tag "Basic" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Company Name" {
            $_.CompanyName | Should -BeTrue -Because "Company Name is required for all users"
        }
    }

    Context "Mobile Phone" -Tag "Basic", "Communication" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Mobile Phone Number" {
            $_.MobilePhone | Should -BeTrue -Because "Mobile Phone is required for all users"
        }
    }

    Context "Business Phone" -Tag "Basic", "Communication" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Business Phone Number" {
            $_.BusinessPhones | Should -BeTrue -Because "Business Phone is required for all users"
        }
    }

    Context "Fax Number" -Tag "Basic", "Communication"  -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Fax Number" {
            $_.FaxNumber | Should -BeTrue -Because "Fax Number is required for all users"
        }
    }

    Context "Manager" -Tag "Basic", "Manager" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Manager" {
            $_.Manager | Should -BeTrue -Because "Manager is required for all users"
        }
    }

    Context "Sponsors" -Tag "Basic", "Sponsor" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Sponsor" {
            $_.Sponsor | Should -BeTrue -Because "Sponsor is required for all users"
        }
    }

    Context "UPN Formatting" -Tag "Custom" -ForEach @( $Users ) {
        It "User $($_.DisplayName) UPN should have a first initial last name all lower case" {
            $firstName = $_.DisplayName
            $lastName = $_.Surname

            $expectedUPN = $firstName.Substring(0, 1).ToLower() + $lastName.ToLower() + "@$DomainName"
            $hascorrectformat = $_.UserPrincipalName -eq $expectedUPN
            $hascorrectformat | Should -BeTrue -Because "UPN should be in the format of first initial last name all lower case"
        }
    }

    Context "User Type" -Tag "User" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a User Type" {
            $_.UserType | Should -Be "Member" -Because "User Type should be 'Member'"
        }
    }

    Context "Employee Hire Date" -Tag "HR" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a Hire Date" {
            $_.EmployeeHireDate | Should -Not -BeNullOrEmpty -Because "Employee Hire Date should not be empty"
        }
    }

    Context "Employee ID" -Tag "HR" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have an Employee ID" {
            $_.EmployeeID | Should -Not -BeNullOrEmpty -Because "Employee ID should not be empty"
        }
    }
}
#EndRegion

#Region User Sign Ins
Describe "Validating User Sign Ins" -Tag "Entra", "Users", "All" {
    Context "7 Days" -Tag "SignIns" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have logged in within the last 7 days" {
            $signInActivity = $_.SignInActivity
            ($signInActivity.LastSignInDateTime -gt (Get-Date).AddDays(-7)) -or ($signInActivity.LastSuccessfulSignInDateTime -gt (Get-Date).AddDays(-7)) | Should -BeTrue -Because "User should have logged in within the last 7 days"
        }
    }

    Context "14 Days" -Tag "SignIns" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have logged in within the last 14 days" {
            $signInActivity = $_.SignInActivity
            ($signInActivity.LastSignInDateTime -gt (Get-Date).AddDays(-14)) -or ($signInActivity.LastSuccessfulSignInDateTime -gt (Get-Date).AddDays(-14)) | Should -BeTrue -Because "User should have logged in within the last 14 days"
        }
    }

    Context "30 Days" -Tag "SignIns" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have logged in within the last 30 days" {
            $signInActivity = $_.SignInActivity
            ($signInActivity.LastSignInDateTime -gt (Get-Date).AddDays(-30)) -or ($signInActivity.LastSuccessfulSignInDateTime -gt (Get-Date).AddDays(-30)) | Should -BeTrue -Because "User should have logged in within the last 30 days"
        }
    }

    Context "60 Days" -Tag "SignIns" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have logged in within the last 60 days" {
            $signInActivity = $_.SignInActivity
            ($signInActivity.LastSignInDateTime -gt (Get-Date).AddDays(-60)) -or ($signInActivity.LastSuccessfulSignInDateTime -gt (Get-Date).AddDays(-60)) | Should -BeTrue -Because "User should have logged in within the last 60 days"
        }
    }

    Context "90 Days" -Tag "SignIns" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have logged in within the last 90 days" {
            $signInActivity = $_.SignInActivity
            ($signInActivity.LastSignInDateTime -gt (Get-Date).AddDays(-90)) -or ($signInActivity.LastSuccessfulSignInDateTime -gt (Get-Date).AddDays(-90)) | Should -BeTrue -Because "User should have logged in within the last 90 days"
        }
    }
}
#EndRegion

#Region Custom User Standards
Describe "Validating Custom User Standards" -Tag "Entra", "Users", "All", "Custom", "CompanyStandard" {
    Context "Company Name Standard Verification" -Tag "CompanyStandard" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a matching Company Name in Excel data" {
            $user = $_
            $validCompanyName = $excelData | Select-Object -ExpandProperty CompanyName -Unique
            try {
                $user.CompanyName | Should -BeIn $validCompanyName -Because "Company Name should be in the list of valid company names"
            }
            catch {
                if (![string]::IsNullOrEmpty($user.Id)) {
                    $selectedCompanyName = Read-Host "The company name of $($user.DisplayName) is not valid. Please enter a valid company name from the list: $($validCompanyName -join ', '), or just press ENTER to skip this update"
                    if ([string]::IsNullOrEmpty($selectedCompanyName) -or $selectedCompanyName -eq 'SKIP') {
                        Write-Host "Skipping update for $($user.DisplayName)"
                        Write-Host "Excel Data: $($excelData.CompanyName)"
                        throw "Update skipped for $($user.DisplayName). Test failed."
                    }
                    elseif ($selectedCompanyName -in $validCompanyName) {
                        Update-MgUser -UserId $user.Id -CompanyName $selectedCompanyName
                        Write-Host "Updated the company name of $($user.DisplayName) to $selectedCompanyName"
                    }
                    else {
                        throw "Invalid company name selected. Test failed."
                    }
                }
                else {
                    throw "Cannot update company name because user ID is null or empty. Test failed."
                }
            }
        }
    }

    Context "Street Standard Verification" -Tag "CompanyStandard" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a matching Street in Excel data" {
            $user = $_
            $validStreetName = $excelData | Select-Object -ExpandProperty StreetAddress -Unique
            try {
                $user.StreetAddress | Should -BeIn $validStreetName -Because "Street Address should be in the list of valid street names"
            }
            catch {
                if (![string]::IsNullOrEmpty($user.Id)) {
                    $selectedStreet = Read-Host "The street address of $($user.DisplayName) is not valid. Please enter a valid street address from the list: $($validStreetName -join ', '), or just press ENTER to skip this update"
                    if ([string]::IsNullOrEmpty($selectedStreet) -or $selectedStreet -eq 'SKIP') {
                        Write-Host "Skipping update for $($user.DisplayName)"
                        throw "Update skipped for $($user.DisplayName). Test failed."
                    }
                    elseif ($selectedStreet -in $validStreetName) {
                        Update-MgUser -UserId $user.Id -StreetAddress $selectedStreet
                        Write-Host "Updated the street address of $($user.DisplayName) to $selectedStreet"
                    }
                    else {
                        throw "Invalid street address selected. Test failed."
                    }
                }
                else {
                    throw "Cannot update street address because user ID is null or empty. Test failed."
                }
            }
        }
    }

    Context "City Standard Verification" -Tag "CompanyStandard" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a matching City in Excel data" {
            $user = $_
            $validCityName = $excelData | Select-Object -ExpandProperty City -Unique
            try {
                $user.City | Should -BeIn $validCityName -Because "City should be in the list of valid city names"
            }
            catch {
                if (![string]::IsNullOrEmpty($user.Id)) {
                    $selectedCity = Read-Host "The city of $($user.DisplayName) is not valid. Please enter a valid city from the list: $($validCityName -join ', '), or just press ENTER to skip this update"
                    if ([string]::IsNullOrEmpty($selectedCity) -or $selectedCity -eq 'SKIP') {
                        Write-Host "Skipping update for $($user.DisplayName)"
                        throw "Update skipped for $($user.DisplayName). Test failed."
                    }
                    elseif ($selectedCity -in $validCityName) {
                        Update-MgUser -UserId $user.Id -City $selectedCity
                        Write-Host "Updated the city of $($user.DisplayName) to $selectedCity"
                    }
                    else {
                        throw "Invalid city selected. Test failed."
                    }
                }
                else {
                    throw "Cannot update city because user ID is null or empty. Test failed."
                }
            }
        }
    }

    Context "State Standard Verification" -Tag "CompanyStandard" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a matching State in Excel data" {
            $user = $_
            $validStates = $excelData | Select-Object -ExpandProperty State -Unique
            try {
                $user.State | Should -BeIn $validStates -Because "State should be in the list of valid states"
            }
            catch {
                if (![string]::IsNullOrEmpty($user.Id)) {
                    $selectedState = Read-Host "The state of $($user.DisplayName) is not valid. Please enter a valid state from the list: $($validStates -join ', '), or just press ENTER to skip this update"
                    if ([string]::IsNullOrEmpty($selectedState) -or $selectedState -eq 'SKIP') {
                        Write-Host "Skipping update for $($user.DisplayName)"
                        throw "Update skipped for $($user.DisplayName). Test failed."
                    }
                    elseif ($selectedState -in $validStates) {
                        Update-MgUser -UserId $user.Id -State $selectedState
                        Write-Host "Updated the state of $($user.DisplayName) to $selectedState"
                    }
                    else {
                        throw "Invalid state selected. Test failed."
                    }
                }
                else {
                    throw "Cannot update state because user ID is null or empty. Test failed."
                }
            }
        }
    }

    Context "Zip Code Standard Verification" -Tag "CompanyStandard" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a matching Zip Code in Excel data" {
            $user = $_
            $validPostalCode = $excelData | Select-Object -ExpandProperty PostalCode -Unique
            try {
                $user.PostalCode | Should -BeIn $validPostalCode -Because "Postal Code should be in the list of valid postal codes"
            }
            catch {
                if (![string]::IsNullOrEmpty($user.Id)) {
                    $selectedPostalCode = Read-Host "The postal code of $($user.DisplayName) is not valid. Please enter a valid postal code from the list: $($validPostalCode -join ', '), or just press ENTER to skip this update"
                    if ([string]::IsNullOrEmpty($selectedPostalCode) -or $selectedPostalCode -eq 'SKIP') {
                        Write-Host "Skipping update for $($user.DisplayName)"
                        throw "Update skipped for $($user.DisplayName). Test failed."
                    }
                    elseif ($selectedPostalCode -in $validPostalCode) {
                        Update-MgUser -UserId $user.Id -PostalCode $selectedPostalCode
                        Write-Host "Updated the postal code of $($user.DisplayName) to $selectedPostalCode"
                    }
                    else {
                        throw "Invalid postal code selected. Test failed."
                    }
                }
                else {
                    throw "Cannot update postal code because user ID is null or empty. Test failed."
                }
            }
        }
    }

    Context "Country Standard Verification" -Tag "CompanyStandard" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a matching Country in Excel data" {
            $user = $_
            $validCountry = $excelData | Select-Object -ExpandProperty Country -Unique
            try {
                $user.Country | Should -BeIn $validCountry -Because "Country should be in the list of valid countries"
            }
            catch {
                if (![string]::IsNullOrEmpty($user.Id)) {
                    $selectedCountry = Read-Host "The country of $($user.DisplayName) is not valid. Please enter a valid country from the list: $($validCountry -join ', '), or just press ENTER to skip this update"
                    if ([string]::IsNullOrEmpty($selectedCountry) -or $selectedCountry -eq 'SKIP') {
                        Write-Host "Skipping update for $($user.DisplayName)"
                        throw "Update skipped for $($user.DisplayName). Test failed."
                    }
                    elseif ($selectedCountry -in $validCountry) {
                        Update-MgUser -UserId $user.Id -Country $selectedCountry
                        Write-Host "Updated the country of $($user.DisplayName) to $selectedCountry"
                    }
                    else {
                        throw "Invalid country selected. Test failed."
                    }
                }
                else {
                    throw "Cannot update country because user ID is null or empty. Test failed."
                }
            }
        }
    }

    Context "Department Standard Verification" -Tag "CompanyStandard" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a matching Department in Excel data" {
            $user = $_
            $validDepartments = $excelData | Select-Object -ExpandProperty Department -Unique
            try {
                $user.Department | Should -BeIn $validDepartments -Because "Department should be in the list of valid departments"
            }
            catch {
                if (![string]::IsNullOrEmpty($user.Id)) {
                    $selectedDepartment = Read-Host "The department of $($user.DisplayName) is not valid. Please enter a valid department from the list: $($validDepartments -join ', '), or just press ENTER to skip this update"
                    if ([string]::IsNullOrEmpty($selectedDepartment) -or $selectedDepartment -eq 'SKIP') {
                        Write-Host "Skipping update for $($user.DisplayName)"
                        throw "Update skipped for $($user.DisplayName). Test failed."
                    }
                    elseif ($selectedDepartment -in $validDepartments) {
                        Update-MgUser -UserId $user.Id -Department $selectedDepartment
                        Write-Host "Updated the department of $($user.DisplayName) to $selectedDepartment"
                    }
                    else {
                        throw "Invalid department selected. Test failed."
                    }
                }
                else {
                    throw "Cannot update department because user ID is null or empty. Test failed."
                }
            }
        }
    }

    Context "Job Title Standard Verification" -Tag "CompanyStandard" -ForEach @( $Users ) {
        It "User $($_.DisplayName) should have a matching Job Title in Excel data" {
            $user = $_
            $validJobTitles = $excelData | Select-Object -ExpandProperty Title -Unique
            try {
                $user.JobTitle | Should -BeIn $validJobTitles -Because "Job Title should be in the list of valid job titles"
            }
            catch {
                if (![string]::IsNullOrEmpty($user.Id)) {
                    $selectedJobTitle = Read-Host "The job title of $($user.DisplayName) is not valid. Please enter a valid job title from the list: $($validJobTitles -join ', '), or just press ENTER to skip this update"
                    if ([string]::IsNullOrEmpty($selectedJobTitle) -or $selectedJobTitle -eq 'SKIP') {
                        Write-Host "Skipping update for $($user.DisplayName)"
                        throw "Update skipped for $($user.DisplayName). Test failed."
                    }
                    elseif ($selectedJobTitle -in $validJobTitles) {
                        Update-MgUser -UserId $user.Id -JobTitle $selectedJobTitle
                        Write-Host "Updated the job title of $($user.DisplayName) to $selectedJobTitle"
                    }
                    else {
                        throw "Invalid job title selected. Test failed."
                    }
                }
                else {
                    throw "Cannot update job title because user ID is null or empty. Test failed."
                }
            }
        }
    }
}
#EndRegion
