 <#
        This script applies the configuration changes for the OnePlace Solutions site to existing site collections.
        A new site collection based on the Team Site template should be created manually before running this script.

        If you are looking to deploy the Content Type to multiple Document Libraries, consider using the example DocumentLibraries.csv file to create a list to point this script to when prompted.
#>

try {    
    Set-ExecutionPolicy Bypass -Scope Process
    
    $script:ContentTypeName = ""
    $script:adminSharePointUrl = ""

    function CreateContentType([string]$arg1){
        Write-Host "Adding Site Content Type '$script:ContentTypeName' to Site Collection '$arg1'"  -ForegroundColor Yellow
        $DocCT = Get-PnPContentType -Identity "Document"
        Add-PnPContentType -name $script:ContentTypeName -Description "Email Content Type for OnePlaceMail" -Group "Custom Content Types" -ParentContentType $DocCT
        }
    
    function AddEmailColumnsToCT([string]$arg1){
        $EmailColumns = Get-PnPField -Group "OnePlace Solutions"
        ForEach($Column in $EmailColumns){
            $Column = $Column.InternalName
            Write-Host "Adding field '$Column' to Site Content Type '$arg1'"  -ForegroundColor Yellow
            Add-PnPFieldToContentType -Field $Column -ContentType $arg1
            }
        }

    function ConnectToSharePoint{
        #Prompt for SharePoint Management Site Url     
        $script:adminSharePointUrl = Read-Host -Prompt 'Please enter the URL of your SharePoint Online Management Site (https://<yourtenant>-admin.sharepoint.com)'

        #Connect to site collection
        If($script:adminSharePointUrl -match "-admin.sharepoint."){
            Write-Host "Enter SharePoint credentials(your email address for SharePoint Online):" -ForegroundColor Green  
            Connect-pnpOnline -url $script:adminSharePointUrl -SPOManagementShell
            }
        Else{
            Write-Host "No valid SharePoint Online Management Site URL entered! Exiting script"
            Pause
            Exit
            }
        }

    function CreateContentTypeInSiteCollectionCSV([string]$arg1){
        $SiteCollectionListFile = Read-Host -Prompt "Please enter the local path to the CSV containing the Site Collections to create the Content Type in"
        $SiteCollectionList = Import-Csv -Path $SiteCollectionListFile
        GetContentTypeName
        Write-Host "Adding Content Type $script:ContentTypeName to Site Collections from CSV..." -ForegroundColor Green

        foreach ($SiteCollection in $SiteCollectionList){
            $SiteColName = $SiteCollection.Name
            $SiteColUrl = $SiteCollection.Url
            Connect-pnpOnline -url $SiteColUrl -SPOManagementShell
            CreateContentType $SiteColName
            AddEmailColumnsToCT $script:ContentTypeName
            Write-Host "`n"
            }
        Write-Host "Complete!" -ForegroundColor Green
        }

    function GetContentTypeName{
        $script:ContentTypeName = ""
        $script:ContentTypeName = Read-Host -Prompt "Please enter the name of your Email Content Type. Leave blank for default 'OnePlaceMail Email'"
        If($script:ContentTypeName -eq ""){$script:ContentTypeName = "OnePlaceMail Email"}
        }

    #start of script
    #begin with connecting to SharePoint
    ConnectToSharePoint

    #Start getting the 
    CreateContentTypeInSiteCollectionCSV
}

catch {
  write-host "Caught an exception:" -ForegroundColor Red
  write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
  write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}