 <#
        This script applies the configuration changes for the OnePlace Solutions site to an existing site collection.
        A new site collection based on the Team Site template should be created manually before running this script.
#>

try {    
    Set-ExecutionPolicy Bypass -Scope Process

    #Prompt for Content Type name, or default to 'OnePlaceMail Email'
    $ContentTypeName = Read-Host -Prompt "Enter the name you want for the Email Content Type. Leave blank for default 'OnePlaceMail Email'"
    If($ContentTypeName -eq ""){$ContentTypeName = "OnePlaceMail Email"}

    #Prompt for SharePoint Url     
    #$SharePointUrl = Read-Host -Prompt 'Enter the url of your site collection to add the Email Content Type to'
    $SharePointUrl = "https://opsdevtest.sharepoint.com/sites/agtestoneplacesolutionsadmin"
       
    #Connect to site collection
    If($SharePointUrl -match ".sharepoint.com/"){
        Write-Host "Enter SharePoint credentials(your email address for SharePoint Online):" -ForegroundColor Green  
        Connect-pnpOnline -url $SharePointUrl -UseWebLogin
        }
    Else{
        Write-Host "Enter SharePoint credentials(domain\username):" -ForegroundColor Green  
        Connect-pnpOnline -url $SharePointUrl
        }
 
 

    Write-Host "Applying configuration changes..." -ForegroundColor Green
    Write-Host "Adding Site Content Type '$ContentTypeName' to Site Collection '$SharePointUrl'"  -ForegroundColor Green

    Add-PnPContentType -name $ContentTypeName -Description "Email Content Type for OnePlaceMail" -Group "Custom Content Types" -ParentContentType "Document"

    $EmailColumns = Get-PnPField -Group "OnePlace Solutions"
    ForEach($Column in $EmailColumns){
        $Column = $Column.InternalName
        Write-Host "Adding field '$Column' to Site Content Type '$ContentTypeName'"  -ForegroundColor Green
        Add-PnPFieldToContentType -Field $Column -ContentType $ContentTypeName
        }
    
    #Apply-PnPProvisioningTemplate -path $Path    

}

catch {
  write-host "Caught an exception:" -ForegroundColor Red
  write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
  write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}