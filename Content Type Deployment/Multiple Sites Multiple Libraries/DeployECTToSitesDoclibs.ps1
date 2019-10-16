<#TODO
#>
Try{
    Set-ExecutionPolicy Bypass -Scope Process

    #Contains all our Site Collections as siteCol objects
    $script:siteColsHT = @{}

    #Contains all the data we need relating to the Site Collection we are working with, including the Document Libraries and the Site Content Type names
    class siteCol{
        [String]$name
        [String]$url
        [Hashtable]$documentLibraries=@{}
        [Array]$contentTypes
        [Boolean]$isSubSite

        siteCol([string]$name,$url){
            If($name -eq ""){
                $this.name = $url
            }
            Else{
                $this.name = $name
            }
            $tempstring = $this.name
            Write-Host "Creating siteCol object with name $tempstring"

            $this.url = $url
            $this.contentTypes = @()

            #Count forward slashes ('/'), if there are more than 4 then we need to check if this URL is for a subsite
            $countFwdSlashes = ($this.url.ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
            
            If($countFwdSlashes -gt 4){
                $indexLastFwdSlash = $this.url.LastIndexOf('/')
                $indexLastFwdSlash++
                #Check the character after the 5th '/', if there's a character we assume this is a subsite URL
                If($url[$indexLastFwdSlash].Length -eq 1){
                    $this.isSubSite = $true
                }
                Else{
                    $this.isSubSite = $false
                }
            }
            Else{
                $this.isSubSite = $false
            }
        }

        [void]addContentTypeToDocumentLibrary($contentTypeName,$docLibName){
            #Check we aren't working without a Document Library name, otherwise assume that we just want to add a Site Content Type
            If(($docLibName -ne $null) -and ($docLibName -ne "")){
                If($this.documentLibraries.ContainsKey($docLibName)){
                    Write-Host "Document Library $docLibName already listed"
                    $this.documentLibraries.$docLibName
                }
                Else{
                    $tempDocLib = [docLib]::new("$docLibName")
                    $this.documentLibraries.Add($docLibName, $tempDocLib)
                }
                $this.documentLibraries.$docLibName.addContentType($contentTypeName)
            }
            
            #If the named Content Type is not already listed in Site Content Types, add it to the Site Content Types
            If(-not $this.contentTypes.Contains($contentTypeName)){
                $this.contentTypes += $contentTypeName
            }
        }
    }

    #Contains all the data we need relating to the Document Library we are working with, including the Site Content Type names
    class docLib{
        [String]$name
        [Array]$contentTypes

        docLib([String]$name){
            Write-Host "Creating docLib object with name $name"
            $this.name = $name
            $this.contentTypes = @()
        }

        [void]addContentType([string]$contentTypeName){
            If(-not $this.contentTypes.Contains($contentTypeName)){
                $this.contentTypes += $contentTypeName
            }
        }
    }

    #Grabs the CSV file and enumerate it into siteColHT as siteCol and docLib objects to work with later
    function EnumerateSitesDocLibs([string]$csvFile){
        If($csvFile -eq ""){
             $csvFile = Read-Host -Prompt "Please enter the local path to the CSV containing the Site Collections and Document Libraries to create the Content Types in"
        }
   
        $csv = Import-Csv -Path $csvFile

        Write-Host "Enumerating Site Collections and Document Libraries from CSV file..." -ForegroundColor Yellow
        foreach ($element in $csv){
            $csv_siteName = $element.SiteName
            $csv_siteUrl = $element.SiteUrl
            $csv_docLib = $element.DocLib
            $csv_contentType = $element.CTName

            #Don't create siteCol objects that do not have a URL, this also accounts for empty lines at EOF
            If($csv_siteUrl -eq ""){Continue}
            If($csv_siteName -eq ""){$csv_siteName = $element.SiteUrl}
            If($script:siteColsHT.ContainsKey($csv_siteUrl)){
                Write-Host "Site $csv_siteName already listed by URL." -ForegroundColor Yellow
                $script:siteColsHT.$csv_siteUrl.addContentTypeToDocumentLibrary($csv_contentType, $csv_docLib)
            }
            Else{
                Write-Host "Site $csv_siteName not listed, adding..." -ForegroundColor Yellow
                $newSiteCollection = [siteCol]::new($csv_siteName, $csv_siteUrl)
                $newSiteCollection.addContentTypeToDocumentLibrary($csv_contentType, $csv_docLib)
                $script:siteColsHT.Add($csv_siteUrl, $newSiteCollection)
            }
        }
        Write-Host "Completed Enumerating Site Collections and Document Libraries from CSV file!" -ForegroundColor Green
    }

    #Facilitates connection to the SharePoint Online site collections through the SharePoint Online Management Shell
    function ConnectToSharePointOnlineAdmin([string]$tenant){
        #Prompt for SharePoint Management Site Url     
        If($tenant -eq ""){
            $tenant = Read-Host -Prompt "Please enter the name of your Office 365 organisation/tenant, eg for 'https://contoso.sharepoint.com' just enter 'contoso'."
        } 

        #Connect to site collection
        $adminSharePointUrl = "https://$tenant-admin.sharepoint.com"
        Write-Host "Enter SharePoint credentials(your email address for SharePoint Online):" -ForegroundColor Green  
        Connect-pnpOnline -url $adminSharePointUrl -SPOManagementShell
        #Sometimes you can continue before authentication has completed, this Start-Sleep adds a delay to account for this
        Start-Sleep -Seconds 3
    }

    #Facilitates connection to the on premises site collections through the root site collection
    function ConnectToSharePointOnPremises([sting]$rootsite){
        #Prompt for SharePoint Root Site Url     
        If($rootsite -eq ""){
            $rootsite = Read-Host -Prompt "Please enter the URL of your on premises SharePoint root site collection"
        }
        Write-Host "Enter SharePoint credentials(your domain login for Sharepoint):" -ForegroundColor Green
        Connect-PnPOnline -url $rootsite
    }

    function CreateEmailColumns([string]$siteCollection){
        If($siteCollection -eq ""){
            $siteCollection = Read-Host -Prompt "Please enter the Site Collection URL to add the OnePlace Solutions Email Columns to"
        }
        Connect-pnpOnline -url $siteCollection -SPOManagementShell
        #From 'https://github.com/OnePlaceSolutions/EmailColumnsPnP/blob/master/installEmailColumns.ps1'
        #Download xml provisioning template
        $WebClient = New-Object System.Net.WebClient   
        $Url = "https://raw.githubusercontent.com/OnePlaceSolutions/EmailColumnsPnP/master/email-columns.xml"    
        $Path = "$env:temp\email-columns.xml"

        Write-Host "Downloading provisioning xml template:" $Path -ForegroundColor Green 
        $WebClient.DownloadFile( $Url, $Path )   
        #Apply xml provisioning template to SharePoint
        Write-Host "Applying email columns template to SharePoint:" $SharePointUrl -ForegroundColor Green 
        Apply-PnPProvisioningTemplate -path $Path
    }
    #Start of Script
    #----------------------------------------------------------------

    #Start with getting the CSV file of Site Collections, Document Libraries and Content Types
    EnumerateSitesDocLibs

    #Connect to SharePoint Online, specifically the Admin site so we can use the SPO Shell to iterate over the site collections
    ConnectToSharePointOnlineAdmin

    $groupName = Read-Host -Prompt "Please enter the Group name containing the OnePlaceMail Email Columns in your SharePoint Site Collections (leave blank for default 'OnePlace Solutions')"
    If($groupName -eq ""){$groupName = "OnePlace Solutions"}
    
    #Loop through columns in group (name supplied above) searching for 'emSubject'
    #Write-Host "Testing for OnePlaceMail Email Columns..." -ForegroundColor Yellow
    
    #Assume columns exist for now
    
    $foundemSubject = $true
    $EmailColumns = Get-PnPField -Group $groupName -InSiteHierarchy
    <#
    ForEach($column in $emailColumns){
        $column = $column.InternalName
        If($column -eq "emSubject"){
            $foundemSubject = $true
            Break
        }
    }
    #>
    #If we don't find 'emSubject', assume our columns aren't installed properly and halt/exit
    If(-not $foundemSubject){
        Write-Host "OnePlaceMail Email column 'emSubject' not found by it's internal name. Please check you have created the Email columns correctly. Halting script. Press enter to open OnePlace Solutions Email Columns PnP Guide and exit." -ForegroundColor Red
        Pause
        Start 'https://github.com/OnePlaceSolutions/EmailColumnsPnP'
        Exit
    }
    Else{
        Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
        #Go through our siteCol objects in siteColsHT
        ForEach($site in $script:siteColsHT.Values){
            $siteName = $site.name
            
            Write-Host "Working with Site Collection: $siteName" -ForegroundColor Yellow
            If($site.isSubSite){
                Write-Host "Subsites currently not compatible with this script due to PnP CmdLet limitation. Skipping this entry"
                Pause
                Continue
            }
            
            Connect-pnpOnline -url $site.url -SPOManagementShell
            
            #Get the Content Type Object for 'Document' from SP, we will use this as the parent Content Type for our email Content Type
            $DocCT = Get-PnPContentType -Identity "Document" -InSiteHierarchy
            If($DocCT -eq $null){
                Write-Host "Couldn't get Document Content Type"
                Pause
                Continue
            }
            #For each Site Content Type listed for this siteCol/Site Collection, try and create it and add the email columns to it
            ForEach($ct in $site.contentTypes){
                Try{
                    Write-Host "Checking if Content Type $ct already exists" -ForegroundColor Yellow
                    $foundContentType =  Get-PnPContentType -Identity $ct -InSiteHierarchy
                
                    #If Content Type object returned is null, assume Content Type does not exist, create it. 
                    #If it does exist and we just failed to find it, this will throw exceptions for 'Duplicate Content Type found', and then continue.
                    If($foundContentType -eq $null){
                        Write-Host "Couldn't find Content Type $ct, might not exist" -ForegroundColor Red
                        #Creating content type
                        Try{
                            Write-Host "Creating Content Type $ct with parent of Document" -ForegroundColor Yellow
                            Add-PnPContentType -name $ct -Group "Custom Content Types" -ParentContentType $DocCT
                        }
                        Catch{
                            Write-Host "Error creating Content Type $ct with parent of Document. Details below. Halting script." -ForegroundColor Red
                            $_
                            Pause
                            Exit
                        } 
                    }
                }
                Catch{
                    Write-Host "Error checking for existence of Content Type $ct. Details below. Halting script." -ForegroundColor Red
                    $_
                    Pause
                    Exit
                }

                #Try adding columns to the Content Type
                Try{
                    Write-Host "Adding email columns to Site Content Type '$ct'"  -ForegroundColor Yellow
                    $numColumns = $emailColumns.Count
                    $i = 0
                    ForEach($column in $emailColumns){
                        $column = $column.InternalName
                        Add-PnPFieldToContentType -Field $column -ContentType $ct
                        Write-Progress -Activity "Adding column: $column" -Status "To Site Content Type: $ct in Site Collection: $siteName. Progress:" -PercentComplete ($i/$numColumns*100)
                        $i++
                    }
                    Write-Progress -Activity "Done adding Columns" -Completed
                }
                Catch{
                    Write-Host "Error adding email columns to Site Content Type $ct. Details below. Halting script." -ForegroundColor Red
                    $_
                    Pause
                    Exit
                }
            }

            #For each docLib/Document Library in our siteCol/Site Collection, get it's list of Content Types we want to add
            ForEach($library in $site.documentLibraries.Values){
                $libName = $library.name
                Write-Host "`nWorking with Document Library: $libName" -ForegroundColor Yellow
                Write-Host "Which has Content Types:" -ForegroundColor Yellow
                $library.contentTypes
                Write-Host "`n"

                Write-Host "Enabling Content Type Management in Document Library '$libName'..." -ForegroundColor Yellow
                Set-PnPList -Identity $libName -EnableContentTypes $true

                #For each Site Content Type listed for this docLib/Document Library, try to add it to said Document Library
                Try{
                    ForEach($ct in $library.contentTypes){
                        Write-Host "Adding Site Content Type '$ct' to Document Library '$libName'..." -ForegroundColor Yellow
                        Add-PnPContentTypeToList -List $libName -ContentType $ct
                    }
                }
                Catch{
                    Write-Host "Error adding Site Content Type '$ct' to Document Library '$libName'. Details below. Halting script." -ForegroundColor Red
                    $_
                    Pause
                    Exit
                }
            }
            Write-Host "`n--------------------------------------------------------------------------------`n" -ForegroundColor Red
        }
    }
}
Catch{$_}