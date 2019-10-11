 <#
        This script applies the configuration changes for the OnePlace Solutions site to existing site collections.
        A new site collection based on the Team Site template should be created manually before running this script.

        If you are looking to deploy the Content Type to multiple Document Libraries, consider using the example DocumentLibraries.csv file to create a list to point this script to when prompted.
#>

try {    
    Set-ExecutionPolicy Bypass -Scope Process
    
    $script:ContentTypeName = ""
    $script:rootSharePointUrl = ""
    $script:onsiteCredentials

    function CreateContentType([string]$arg1){
        GetContentTypeName
        If($arg1 -ne ""){$SharePointUrl = $arg1}

        Write-Host "Adding Site Content Type '$script:ContentTypeName' to Site Collection '$SharePointUrl'"  -ForegroundColor Green
        $DocCT = Get-PnPContentType -Identity "Document"
        Add-PnPContentType -name $script:ContentTypeName -Description "Email Content Type for OnePlaceMail" -Group "Custom Content Types" -ParentContentType $DocCT
        }
    
    function AddColumnsToCT{
        $EmailColumns = Get-PnPField -Group "OnePlace Solutions"
        ForEach($Column in $EmailColumns){
            $Column = $Column.InternalName
            Write-Host "Adding field '$Column' to Site Content Type '$script:ContentTypeName'"  -ForegroundColor Green
            Add-PnPFieldToContentType -Field $Column -ContentType $script:ContentTypeName
            }
        }

    function ConnectToSharePoint{
        #Prompt for SharePoint Url     
        $script:rootSharePointUrl = Read-Host -Prompt 'For onsite SharePoint, enter the url of your root site collection. For SharePoint Online, please enter the URL of your Management Site (https://<yourtenant>-admin.sharepoint.com)'

        #Connect to site collection
        If($script:rootSharePointUrl -match ".sharepoint.com"){
            Write-Host "Enter SharePoint credentials(your email address for SharePoint Online):" -ForegroundColor Green  
            Connect-pnpOnline -url $script:rootSharePointUrl -UseWebLogin -SPOManagementShell
            }
        ElseIf($script:rootSharePointUrl -match ""){
            Write-Host "No SharePoint URL entered!"
            Pause
            Break
            }
        Else{
            Write-Host "Enter SharePoint credentials(domain\username):" -ForegroundColor Green  
            $script:onsiteCredentials = Get-Credential
            Connect-pnpOnline -url $script:rootSharePointUrl -Credentials $script:onsiteCredentials
            }
        }

    <#
    function GetSiteCollectionUrlsCSV{     
        $SiteCollectionListFile = Read-Host -Prompt "Please enter the local path to the CSV containing the Site Collections to create the Content Type in"
        $SiteCollectionList = Import-Csv -Path $SiteCollectionListFile

        Write-Host "Importing Site Collections from CSV..." -ForegroundColor Green
        $script:SiteCollectionUrls.Clear()

        foreach ($SiteCollection in $SiteCollectionList){
            $SiteColName = $SiteCollection.Name
            $SiteColUrl = $SiteCollection.Url
            $script:SiteCollectionUrls.Add($siteColName, $SiteColUrl)
            }
        Write-host "Site Collections imported: `n"
        $script:SiteCollectionUrls
        Pause
        }
    #>

    function CreateContentTypeInSiteCollectionCSV{
        $SiteCollectionListFile = Read-Host -Prompt "Please enter the local path to the CSV containing the Site Collections to create the Content Type in"
        $SiteCollectionList = Import-Csv -Path $SiteCollectionListFile

        Write-Host "Adding Content Type to Site Collections from CSV..." -ForegroundColor Green

        foreach ($SiteCollection in $SiteCollectionList){
            $SiteColName = $SiteCollection.Name
            $SiteColUrl = $SiteCollection.Url
            Connect-pnpOnline -url $SiteColUrl

            }
        }

    function GetContentTypeName{
        $script:ContentTypeName = ""
        $script:ContentTypeName = Read-Host -Prompt "Please enter the name of your Email Content Type. Leave blank for default 'OnePlaceMail Email'"
        If($script:ContentTypeName -eq ""){$script:ContentTypeName = "OnePlaceMail Email"}
        }

    #start of script
    #begin with connecting to SharePoint
    ConnectToSharePoint
    do{
        ShowMenu
        $input = Read-Host "Please make a selection"
        switch ($input){
            '1'{
                cls
                CreateContentType
                AddColumnsToCT
                Pause
                }
            '2'{
                cls
                AddContentTypeToLibrary
                }
            '3'{
                cls
                AddContentTypeToDocumentLibrariesCSV
                }
            '4'{
                cls
                'Opening link to SharePoint PnP CmdLets Github...'
                start 'https://github.com/SharePoint/PnP-PowerShell'
                Pause
                }
            }
        }
    until($input -eq 'q')
}

catch {
  write-host "Caught an exception:" -ForegroundColor Red
  write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
  write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}