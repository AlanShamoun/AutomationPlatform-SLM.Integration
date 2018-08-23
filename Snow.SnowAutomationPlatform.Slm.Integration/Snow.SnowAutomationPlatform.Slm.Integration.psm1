Function Get-SLMStoreApplications {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SLMServerName,

        [Parameter(Mandatory=$true)]
        [string]$SLMCustomerId,

        [Parameter(Mandatory=$true)]
        [pscredential]$Credentials,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,1000)]
        [int]$AppsPerRequest = 200,

        [Parameter(Mandatory=$false)]
        [switch]$GetAllApps
    )

    Begin 
    {
        # the status converter is used to convert SLM status values into AP statuses (in AP 0 is Active and 2 is retired)
        $statusConverter = @{
            'active'   = 0
            'inactive' = 2
        }

        #region internal functions
            Function Add-SLMApps
            {
                param(
                    $TargetArrayList,
                    $StoreAppsJSON
                )

                Foreach ($StoreAppObj in $StoreAppsJSON.Body.Body)
                {
                    $propHash = @{
                        'ApplicationOwnerApproval'       = [bool]$StoreAppObj.ApplicationOwnerApproval
                        'ApplicationOwnerUsername'       = $StoreAppObj.ApplicationOwnerUsername
                        'ApplicationTypes'               = $StoreAppObj.ApplicationTypes -join ', '
                        'ComputerGroup'                  = [bool]$StoreAppObj.ComputerGroup
                        'Currency'                       = $StoreAppObj.Currency
                        'DaysUntilUninstall'             = $StoreAppObj.DaysUntilUninstall
                        'DaysUntilUninstallNotification' = $StoreAppObj.DaysUntilUninstallNotification
                        'Description'                    = $StoreAppObj.Description
                        'GroupName'                      = $StoreAppObj.GroupName
                        'Id'                             = $StoreAppObj.Id
                        'ImageName'                      = $StoreAppObj.ImageName
                        'Manufacturer'                   = $StoreAppObj.Manufacturer
                        'Name'                           = $StoreAppObj.Name
                        'OrganizationalApproval'         = [bool]$StoreAppObj.OrganizationalApproval
                        'PublishLevel'                   = $StoreAppObj.PublishLevel
                        'PurchasePrice'                  = $StoreAppObj.PurchasePrice
                        'RentalPaymentPeriod'            = $StoreAppObj.RentalPaymentPeriod
                        'RentalPrice'                    = $StoreAppObj.RentalPrice
                        'SecondaryApproval'              = [bool]$StoreAppObj.SecondaryApproval
                        'SecondaryApprovalUsername'      = $StoreAppObj.SecondaryApprovalUsername
                        'Status'                         = $statusConverter.($StoreAppObj.Status)
                        'SubscriptionExtensionsDays'     = $StoreAppObj.SubscriptionExtensionsDays
                        'ThirdApproval'                  = [bool]$StoreAppObj.ThirdApproval
                        'ThirdApprovalUsername'          = $StoreAppObj.ThirdApprovalUsername
                        'UninstallOption'                = $StoreAppObj.UninstallOption
                        'UserGroup'                      = [bool]$StoreAppObj.UserGroup
                    }
                    $NewObject = New-Object PSObject -Property $propHash
                    $null = $TargetArrayList.Add($NewObject)
                }
            }
        #endregion
    }

    Process 
    {
        #region Declaring Variables and function
            $CleanSLMUri = "$SLMServerName/api/customers/$SLMCustomerId/appstore/applications/"
        #endregion

        #region Getting the first Applications from Rest API
            $firstRequestUri = $CleanSLMUri + '?$top=' + $AppsPerRequest + '&$inlinecount=allpages&$format=json'
            $StoreApps = Invoke-RestMethod -Uri $firstRequestUri -Credential $Credentials
        #endregion

        #region Validating collected apps and expected apps
            $AmountOfApps = [int]($StoreApps.Meta | Where-Object {$_.Name -eq "Count"}).Value
        #endregion

        #region Populating already collected SLM Apps to StoreAppsList arraylist
            [System.Collections.ArrayList]$StoreAppsList = @()
            Add-SLMApps -TargetArrayList $StoreAppsList -StoreAppsJSON $StoreApps
        #endregion

        #region Getting All SLM apps if requested with parameter GetAllApps
            if ($GetAllApps)
            {
                $AmountOfExtraCalls = [Math]::Ceiling($AmountOfApps/$AppsPerRequest)-1 #Subtraction one for the first call already done
                if ($AmountOfExtraCalls -gt 0)
                {
                    Write-Verbose "$AmountOfExtraCalls additional API calls needed to get all SLM store applications"
                    $SkipValue = $AppsPerRequest
                    for ($i = 1; $i -le $AmountOfExtraCalls; $i++) 
                    {
                        $PreSkipValueForLogging = $SkipValue
                        $extraRequestUri = $CleanSLMUri + '?$top=' + $AppsPerRequest + '&$skip=' + $SkipValue + '&$format=json'
                        $MoreSLMApps = Invoke-RestMethod -Uri $extraRequestUri -Credential $Credentials
                        Add-SLMApps -TargetArrayList $StoreAppsList -StoreAppsJSON $MoreSLMApps
                        $SkipValue = $SkipValue + $AppsPerRequest
                        Write-Verbose "Call $i/$AmountOfExtraCalls, collected SLM apps $PreSkipValueForLogging to $SkipValue"
                    }
                }
            }
        #endregion

        $StoreAppsList
    }
}

Function Get-SLMReharvestInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SLMServerName,

        [Parameter(Mandatory=$true)]
        [string]$SLMCustomerId,

        [Parameter(Mandatory=$true)]
        [pscredential]$Credentials,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,1000)]
        [int]$InitialAppLoad = 200,

        [Parameter(Mandatory=$false)]
        [boolean]$GetAllApps = $false
    )

    #region Declaring Variables and function
        $ResultsPerPage = $InitialAppLoad
        $TopResultsString = '?$top='
        $TopResults = '?$top=' + $ResultsPerPage
        $JsonData = '$format=json'
        $ShowAllPages = '$inlinecount=allpages'
        $CleanSLMUri = "$SLMServerName/api/customers/$SLMCustomerId/appstore/applications/reharvestable/"

        Function Add-SLMApps{
        param($TargetArrayList, $StoreAppsJSON)
            Foreach ($StoreAppObj in $StoreAppsJSON.Body){
                $propHash = [Ordered]@{}
                $propHash.'Name' = $StoreAppObj.Body.Name
                $propHash.'Manufacturer' = $StoreAppObj.Body.Manufacturer
                $propHash.'PurchasePrice' = $StoreAppObj.Body.PurchasePrice
                $propHash.'RentalPrice' = $StoreAppObj.Body.RentalPrice
                $propHash.'Currency' = $StoreAppObj.Body.Currency
                $propHash.'GroupName' = $StoreAppObj.Body.GroupName
                $propHash.'UserGroup' = $StoreAppObj.Body.UserGroup
                $propHash.'ComputerGroup' = [bool]$StoreAppObj.Body.ComputerGroup
                $propHash.'PublishLevel' = $StoreAppObj.Body.PublishLevel
                $propHash.'UninstallOption' = $StoreAppObj.Body.UninstallOption
                $propHash.'DaysUntilUninstall' = $StoreAppObj.Body.DaysUntilUninstall
                $propHash.'DaysUntilUninstallNotification' = $StoreAppObj.Body.DaysUntilUninstallNotification
                $propHash.'ComputerName' = $StoreAppObj.Body.ComputerName
                $propHash.'FirstUsed' = $StoreAppObj.Body.FirstUsed
                $propHash.'LastUsed' = $StoreAppObj.Body.LastUsed
                $propHash.'LastScanDate' = $StoreAppObj.Body.LastScanDate
                $propHash.'MostFrequentComputerUser' = $StoreAppObj.Body.MostFrequentComputerUser
                $propHash.'VendorItemId' = $StoreAppObj.Body.ID

                $NewObject = New-Object PSObject -Property $propHash
                $null = $TargetArrayList.Add($NewObject)
            }
        }
    #endregion

    #region Getting the first Applications from Rest API
        $StoreApps = Invoke-RestMethod -Uri $($CleanSLMUri + $TopResultsString + $ResultsPerPage + '&' + $ShowAllPages + '&' + $JsonData) -Credential $Credentials
    #endregion

    #region Validating collected apps and expected apps
        $AmountOfApps = [int]($StoreApps.Meta | Where-Object {$_.Name -eq "Count"}).Value
        if($StoreApps.Body.Count -ge $AmountOfApps){ Write-Verbose "All $AmountOfApps apps collected"}
        elseif($StoreApps.Body.Count -lt $AmountOfApps){
            if($GetAllApps){ Write-Verbose "$AmountOfApps applications available" }
            else{ Write-Verbose "$AmountOfApps applications available, only collecting the first $InitialAppLoad" }
        }
    #endregion

    #region Populating already collected SLM Apps to StoreAppsList arraylist
        [System.Collections.ArrayList]$StoreAppsList = @()
        Add-SLMApps -TargetArrayList $StoreAppsList -StoreAppsJSON $StoreApps
    #endregion

    #region Getting All SLM apps if requested with parameter GetAllApps
        if($GetAllApps){
            $AmountOfExtraCalls = [Math]::Ceiling($AmountOfApps/$ResultsPerPage)-1 #Subtraction one for the first call already done
            if($AmountOfExtraCalls -gt 0){
                Write-Verbose "$AmountOfExtraCalls additional API calls needed to get all SLM store applications"
                $SkipValue = $ResultsPerPage

                for ($i = 1; $i -le $AmountOfExtraCalls; $i++)
                {
                    $PreSkipValueForLogging = $SkipValue

                    $MoreSLMApps = Invoke-RestMethod -Uri $($CleanSLMUri + $TopResultsString + $ResultsPerPage + '&' + '$skip=' + $SkipValue + '&' + $JsonData) -Credential $Credentials
                    Add-SLMApps -TargetArrayList $StoreAppsList -StoreAppsJSON $MoreSLMApps

                    $SkipValue = $SkipValue + $ResultsPerPage
                    Write-Verbose "Call $i/$AmountOfExtraCalls, collected SLM apps $PreSkipValueForLogging to $SkipValue"
                }
            }
        }
    #endregion

    $StoreAppsList
}

Function Get-SLMCustomField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SLMServerName,

        [Parameter(Mandatory=$true)]
        [string]$SLMCustomerId,

        [Parameter(Mandatory=$true)]
        [pscredential]$Credentials,

        [Parameter(Mandatory=$true)]
        [string]$ApplicationId

    )

    Process 
    {
        $CleanSLMUri = "$SLMServerName/api/customers/$SLMCustomerId/applications/$ApplicationId/?`$format=json"
        
        $StoreApps = Invoke-RestMethod -Uri $CleanSLMUri -Credential $Credentials
        
        foreach ($CustomValue in $StoreApps.Body.CustomValues) 
        {
            $props = [ordered]@{
                Name = $CustomValue.Name
                Value = $CustomValue.Value
            }

            New-Object -TypeName psobject -Property $props
        }    
    }

}

Function Get-SLMUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SLMServerName,

        [Parameter(Mandatory=$true)]
        [string]$SLMCustomerId,

        [Parameter(Mandatory=$true)]
        [pscredential]$Credentials,

        [Parameter(Mandatory=$true)]
        [string]$UserName

    )

    Process 
    {
        $CleanSLMUri = "$SLMServerName/api/customers/$SLMCustomerId/users/?%24filter=Username+eq+%27$UserName%27&`$format=json"
        $SLMUser = Invoke-RestMethod -Uri $CleanSLMUri -Credential $Credentials
         
        if ($SLMUser.body.body) {
            if ($SLMUser.body.body.Id) {
                $props = @{
                    Id = $SLMUser.body.body.Id
                    UserName = $SLMUser.body.body.UserName
                    LastLogin = $SLMUser.body.body.LastLogin
                    FullName = $SLMUser.body.body.FullName
                }
            
                New-Object -TypeName psobject -Property $props
            }
            else {
                Write-Error "User ID from SLM is null"
            }
        } else {
            Write-Error "No user found matching $UserName"
        }
    }

}

Function Get-SLMComputer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SLMServerName,

        [Parameter(Mandatory=$true)]
        [string]$SLMCustomerId,

        [Parameter(Mandatory=$true)]
        [pscredential]$Credentials,

        [Parameter(Mandatory=$true)]
        [int]$MostFrequentUserid

    )

    Process 
    {
        # search for computers (this will get brief information about the computers)
        $CleanSLMUri = "$SLMServerName/api/customers/$SLMCustomerId/computers/?%24filter=MostFrequentUserId%20eq%20$MostFrequentUserid&`$format=json"
        $SLMComputer = Invoke-RestMethod -Uri $CleanSLMUri -Credential $Credentials
         
        if ($SLMComputer.body.body) {
            #return $SLMComputer.body.body

            # get detailed information about each computer
            foreach ($href in $SLMComputer.Body.Links.Where({$_.Rel -eq 'Self'}).Href) {
                $SLMComputerDetails = Invoke-RestMethod -Uri $href -Credential $Credentials
                if ($SLMComputerDetails.Body) {
                    $SLMComputerDetails.Body
                }
            }
        } else {
            Write-Error "No computer found with mosts recent user with id $MostFrequentUserid"
        }
    }

}
