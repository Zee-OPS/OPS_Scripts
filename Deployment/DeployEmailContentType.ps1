 <#
        This script applies the configuration changes for the OnePlace Solutions site to existing site collections.
        A new site collection based on the Team Site template should be created manually before running this script.

        If you are looking to deploy the Content Type to multiple Document Libraries, consider using the example DocumentLibraries.csv file to create a list to point this script to when prompted.
#>

try {    
    Set-ExecutionPolicy Bypass -Scope Process
    
    $script:ContentTypeName = ""
    $script:SharePointUrl = ""

    function ShowMenu{
        cls
        Write-Host "`n--------------------------------------------------------------------------------`n"  -ForegroundColor Red
        Write-Host 'Welcome to the OnePlace Solutions Content Type Deployment Script.'  -ForegroundColor Yellow
        Write-Host 'Please ensure you have the correct PnP CmdLets installed for your SharePoint Environment, and you have installed the Email Site Columns before continuing.'  -ForegroundColor Yellow
        Write-Host "`n--------------------------------------------------------------------------------`n"  -ForegroundColor Red

        Write-Host "1: Create Content Type in a single Site Collection"
        Write-Host "2: Add Content Type to a single Document Library"
        Write-Host "3: Add Content Type to a CSV list of Document Libraries"
        Write-Host "4: Visit SharePoint PnP CmdLets Github"
        Write-Host "Q: Press 'Q' to quit."
        }
    
    function CreateContentType{
        GetContentTypeName

        Write-Host "Adding Site Content Type '$script:ContentTypeName' to Site Collection '$SharePointUrl'"  -ForegroundColor Green
        $DocCT = Get-PnPContentType -Identity "Document"
        Add-PnPContentType -name $script:ContentTypeName -Description "Email Content Type for OnePlaceMail" -Group "Custom Content Types" -ParentContentType $DocCT
        }
    
    function AddColumnsToCT {
        $EmailColumns = Get-PnPField -Group "OnePlace Solutions"
        ForEach($Column in $EmailColumns){
            $Column = $Column.InternalName
            Write-Host "Adding field '$Column' to Site Content Type '$script:ContentTypeName'"  -ForegroundColor Green
            Add-PnPFieldToContentType -Field $Column -ContentType $script:ContentTypeName
            }
        }

    function ConnectToSharePoint{
        #Prompt for SharePoint Url     
        $script:SharePointUrl = Read-Host -Prompt 'Enter the url of your site collection to work with'

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
    
    function EnableContentTypeManagement([string]$arg1){
        Set-PnPList -Identity $arg1 -EnableContentTypes $true
        }

    function GetContentTypeName{
        $script:ContentTypeName = ""
        $script:ContentTypeName = Read-Host -Prompt "Please enter the name of your Email Content Type. Leave blank for default 'OnePlaceMail Email'"
        If($script:ContentTypeName -eq ""){$script:ContentTypeName = "OnePlaceMail Email"}
        }

    function AddContentTypeToLibrary([string]$arg1){
        If($script:ContentTypeName -eq ""){GetContentTypeName}

        If($arg1 -eq ""){
            $DocLib = Read-Host -Prompt "Enter the name of the Document Library to enable Content Type Management for and add the Email Content Type"
            }
        Else{
            $DocLib = $arg1
            }

        Write-Host "Enabling Content Type Management in Document Library '$DocLib'..." -ForegroundColor Green
        Set-PnPList -Identity $DocLib -EnableContentTypes $true
        Write-Host "Adding Email Content Type '$script:ContentTypeName' to Document Library '$DocLib'..." -ForegroundColor Green
        Add-PnPContentTypeToList -List $DocLib -ContentType $script:ContentTypeName
        }

     function AddContentTypeToDocumentLibrariesCSV{
        $DocumentLibraryListFile = Read-Host -Prompt "Please enter the local path to the CSV containing the Document Libraries to add the Content Type to"
        $DocumentLibraryList = Import-Csv -Path $DocumentLibraryListFile

        foreach ($DocumentLibrary in $DocumentLibraryList){
            $DocLibName = $DocumentLibrary.Name
            AddContentTypeToLibrary $DocLibName
            }
        Write-Host "Finished list!" -ForegroundColor Green
        Pause
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