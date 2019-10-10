 <#
        This script applies the configuration changes for the OnePlace Solutions site to existing site collections.
        A new site collection based on the Team Site template should be created manually before running this script.
#>

try {    
    Set-ExecutionPolicy Bypass -Scope Process
    
    $script:ContentTypeName = ""
    $script:SharePointUrl = ""

    function ShowMenu{
        cls
        Write-Host "`n--------------------------------------------------------------------------------`n"  -ForegroundColor Red
        Write-Host 'Welcome to the OnePlace Solutions Content Type Deployment Script.'  -ForegroundColor Yellow
        Write-Host 'Please ensure you have the correct PnP CmdLets installed for your SharePoint Environment before continuing.'  -ForegroundColor Yellow
        Write-Host "`n--------------------------------------------------------------------------------`n"  -ForegroundColor Red

        Write-Host "1: Add Content Type to a single Site Collection"
        Write-Host "2: Add Content Type to a CSV list of Site Collections"
        Write-Host "3: Visit SharePoint PnP CmdLets Github"
        Write-Host "Q: Press 'Q' to quit."
        }
    
    function CreateContentType{
        #Prompt for Content Type name, or default to 'OnePlaceMail Email'
        $script:ContentTypeName = ""
        $script:ContentTypeName = Read-Host -Prompt "Enter the name you want for the Email Content Type. Leave blank for default 'OnePlaceMail Email'"
        If($script:ContentTypeName -eq ""){$script:ContentTypeName = "OnePlaceMail Email"}

        Write-Host "Adding Site Content Type '$script:ContentTypeName' to Site Collection '$SharePointUrl'"  -ForegroundColor Green
        $DocCT = Get-PnPContentType -Identity "Document"
        Add-PnPContentType -name $script:ContentTypeName -Description "Email Content Type for OnePlaceMail" -Group "Custom Content Types" -ParentContentType $DocCT
        }
    
    function AddColumnsToCT {
        $EmailColumns = Get-PnPField -Group "OnePlace Solutions"
        ForEach($Column in $EmailColumns){
            $Column = $Column.InternalName
            Write-Host "Adding field '$Column' to Site Content Type '$ContentTypeName'"  -ForegroundColor Green
            Add-PnPFieldToContentType -Field $Column -ContentType $ContentTypeName
            }
        }

    function GetSiteCollectionsCSV{
         #Prompt for Tenant url
        $TenantUrl = Read-Host -Prompt 'Enter your SharePoint online tenant url'
        Connect-pnpOnline -url $TenantUrl

        $siteCollectionList = Import-Csv -Path "C:\temp\SiteCollections.csv"
        
        #Loop through csv and provision site collection from each csv entry
        foreach ($siteCollection in $siteCollectionList){
            $SharePointUrl = $siteCollection.Url
            $SiteOwner = $siteCollection.Owner
            $Title = $siteCollection.Title
            $Template = $siteCollection.SiteTemplate
            $TimeZone = $siteCollection.TimeZone       

            #Create site collection based on values above        
            New-PnPTenantSite -Owner $SiteOwner -Title $Title -Url $SharePointUrl -Template $Template -TimeZone $TimeZone  
            }
        
        }

    function ConnectToSharePoint{
        #Prompt for SharePoint Url     
        $script:SharePointUrl = Read-Host -Prompt 'Enter the url of your site collection to add the Email Content Type to'

        #Connect to site collection
        If($script:SharePointUrl -match ".sharepoint.com/"){
            Write-Host "Enter SharePoint credentials(your email address for SharePoint Online):" -ForegroundColor Green  
            Connect-pnpOnline -url $script:SharePointUrl -UseWebLogin
            }
        ElseIf($script:SharePointUrl -match ""){
            Write-Host "No SharePoint URL entered!"
            Pause
            Break
            }
        Else{
            Write-Host "Enter SharePoint credentials(domain\username):" -ForegroundColor Green  
            Connect-pnpOnline -url $script:SharePointUrl
            }
        }

    do{
        ShowMenu
        $input = Read-Host "Please make a selection"
        switch ($input){
            '1'{
                cls
                ConnectToSharePoint
                CreateContentType
                AddColumnsToCT
                Pause
                }
            '2'{
                cls
                Write-Host "Not Implemented yet!"
                <#
                ConnectToSharePoint
                CreateContentType
                AddColumnsToCT
                #>
                Pause
                }
            '3'{
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