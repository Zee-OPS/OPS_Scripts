Try{
    Set-ExecutionPolicy Bypass -Scope Process
    $ErrorActionPreference = 'stop' 

    $script:Login = ""

    #If you would like to skip confirming the User's Office 365 Login, set this to $false
    $script:confirm = $true

    #Please enter your SharePoint Online URL in the line below, eg "https://testdomain.sharepoint.com/", or we will try to retrieve it from the License List URL
    $script:SharePointUrl = ""

    #Please enter your main Office 365 email domain in the line below, eg "onmicrosoft.com", or we will prompt the user to enter it
    $script:Domain = ""

    function ConfirmLogin{
        If($script:confirm){
            $input = Read-Host "Your Office 365 Login is $script:Login, is this correct? (yes or no)"

            switch ($input) `
                {
                    'yes' {}

                    'no' {
                        PromptUserAndDomain($true)
                        }

                    default {
                        Write-Host "Exiting!"
                        Pause
                        Exit
                        }
                }
            }
        }

    function PromptUserAndDomain ([boolean]$arg1){
        $script:Username = Read-Host -Prompt 'Input your Office 365 Username (excluding domain), eg: john.smith'

        If(($script:Domain -eq "") -or $arg1){
            $script:Domain = Read-Host -Prompt 'Input your Office 365 Domain, eg: onmicrosoft.com'
            }

        $script:Login = "$script:Username@$script:Domain"

        ConfirmLogin    
        If(($script:Domain -eq "" ) -or ($script:Username -eq "")){
            Write-Host "No Username or Domain entered!" -ForegroundColor Red
            PromptUserAndDomain($true)
            }
        }

    function AddOneDriveToFavourites([string]$arg1) {
    
        $tempOneDriveUrl = $arg1
    
        Try{
            Write-Host "Adding to OnePlace Site Collections..."

            $env:Path = "$env:APPDATA" + "\OnePlace Solutions\Favorites.xml"
            [xml]$xmldoc = Get-Content $env:Path
            $newSiteCol = $xmldoc.configuration.siteCollections.AppendChild($xmldoc.CreateElement("siteCollection"))
            $newSiteCol.SetAttribute("siteColUrl",$tempOneDriveUrl)
            $newSiteCol.SetAttribute("title","OneDrive")
            $xmldoc.Save($env:Path)
            Write-Host "Successfully added OneDrive URL to OnePlace Site Collections" -ForegroundColor Green

            Write-Host "Adding to OnePlace Favorites..."

            $newFolder = $xmldoc.configuration.outlookFolders.AppendChild($xmldoc.CreateElement("outlookFolder"))
            $newFolder.SetAttribute("folderName","My OneDrive Documents")
            $newFolder.SetAttribute("folderUrl", "$tempOneDriveUrl/Documents")
            $newFolder.SetAttribute("useGlobalSettings","true")
            $newFolder.SetAttribute("listType","library")
            $newFolder.SetAttribute("listUrl","$tempOneDriveUrl/Documents")
            $newFolder.SetAttribute("folderId","")
            $newFolder.SetAttribute("isSubFolder","False")
            $xmldoc.Save($env:Path)
            Write-Host "Successfully added My OneDrive Documents to OnePlace Favourites with URL: $tempOneDriveUrl" -ForegroundColor Green
            }
        Catch{
            $_
            }
        }

    function FetchTenantFromLicenseUrl {
        Try{
            Write-Host "Trying to find License List URL in XML..." -ForegroundColor Yellow
            $env:Path = "$env:APPDATA" + "\OnePlace Solutions\CommonConfig.xml"
            [xml]$xmldoc = Get-Content $env:Path
            $licenseUrl = $xmldoc.configuration.license.licenseLocation
            If($licenseUrl -match "http"){
                $script:SharePointUrl= $licenseUrl
                Write-Host "Found License List URL in XML!" -ForegroundColor Green
                }
            Else{
                Throw
                }
            }
        Catch{
            Write-Host "Failed to find License List URL in XML, trying registry..." -ForegroundColor Yellow
            Try{
                $env:Path = "HKLM:\SOFTWARE\WOW6432Node\OnePlace Solutions"
                $tempobject = Get-ItemProperty "$env:Path\Common" -name "licenseLocation"
                If($tempobject.licenseLocation -match "http"){
                    $script:SharePointUrl = $tempobject.licenseLocation
                    Write-Host "Found License List URL in 64 bit registry!" -ForegroundColor Green
                    }
                }
            Catch{
                Write-Host "Failed to find License List URL in 64 bit registry location, trying 32 bit..." -ForegroundColor Yellow
                $env:Path = "HKLM:\SOFTWARE\OnePlace Solutions"
                $tempobject = Get-ItemProperty "$env:Path\Common" -name "licenseLocation"
                If($tempobject.licenseLocation -match "http"){
                    $script:SharePointUrl = $tempobject.licenseLocation
                    Write-Host "Found License List URL in 32 bit registry!" -ForegroundColor Green
                    }
                }
            }
        Finally{
            If($script:SharePointUrl -notmatch "http"){
                Write-Host "No License List set in OnePlace or Registry. Please add a License List URL to OnePlace and run the script again, or contact your SharePoint Administrator." -ForegroundColor Red
                Pause
                Exit
                }
            }
        }

    function GenerateOneDriveUrl{
        If($SharePointUrl -notmatch "http"){
            FetchTenantFromLicenseUrl
            }
        $script:Tenant = $script:SharePointUrl -replace ".sharepoint.*" -replace ".*//"
        $script:OneDriveUrl = "https://$script:Tenant-my.sharepoint.com/personal/$script:SanitizedUsername"
        Write-Host "Your OneDrive URL is: "
        Write-Host $script:OneDriveUrl -ForegroundColor Yellow
        }

    PromptUserAndDomain

    $script:SanitizedUsername = $script:Login -replace '[^a-z0-9 | -]','_'

    $script:Tenant = ""
    $script:OneDriveUrl = ""

    GenerateOneDriveUrl
    AddOneDriveToFavourites($script:OneDriveUrl)
    Pause
    }
Catch{Write-Host $_}