<#
OnePlace Solutions Troubleshooting Script
Version: 1.5 Read Only
Author: agregory
LastEdit: agregory

This script is intended for use by OnePlace Solutions staff, their partners, and support contract holders.
No registry or XML values are written by this script, it is purely a reporting tool that can offer troubleshooting direction and diagnostic data for further support.
If launched with the Launcher script it will save it's output to a text file in the same directory.

#>
Try{
    Set-ExecutionPolicy Bypass -Scope Process
    $script:VerboseMode=$false
    $ErrorActionPreference = 'stop' 

    $OfficeInfo = New-Object -TypeName psobject -Property ($OfficeInfoHT = [ordered]@{})

    $LoadBehavior = New-Object -TypeName psobject -Property ($LoadBehaviorHT = [ordered]@{})

    $Resiliency = New-Object -TypeName psobject -Property ($ResiliencyHT = [ordered]@{})

    $LicenseAndSolutions = New-Object -TypeName psobject -Property ($LicenseAndSolutionsHT = [ordered]@{})

    $DOTNET = New-Object -TypeName psobject -Property ($DOTNETHT = [ordered]@{})

    $VSTO = New-Object -TypeName psobject -Property ($VSTOHT = [Ordered]@{})

    $XML = New-Object -TypeName psobject -Property ($XMLHT = [Ordered]@{})

    $OPM = New-Object -TypeName psobject -Property ($OPMHT = [Ordered]@{})

    $LocalMachineDetails = New-Object -TypeName psobject -Property ($LocalMachineDetailsHT = [Ordered]@{
        OSInstalled = "Not Found"
        OSBitness = "86"
        OS64Bitness = $false
    })

    $OPMBuilds = New-Object -TypeName psobject -Property ($OPMBuildsHT = @{
        R792 = "30.29.19226.3"
        R791 = "30.29.19121.5"
        R78 = "30.29.18269.3"
        R77 = "30.29.18130.9"
        R761 = "30.29.18064.7"
        R752 = "30.23.17289.0"
    })

                                                                                                function ReInitialiseVariables{
    $OfficeInfoHT.Clear()
    $OfficeInfo = New-Object -TypeName psobject -Property ($OfficeInfoHT = [ordered]@{})
    $LoadBehaviorHT.Clear()
    $LoadBehavior = New-Object -TypeName psobject -Property ($LoadBehaviorHT = [ordered]@{})
    $ResiliencyHT.Clear()
    $Resiliency = New-Object -TypeName psobject -Property ($ResiliencyHT = [ordered]@{})
    $LicenseAndSolutionsHT.Clear()
    $LicenseAndSolutions = New-Object -TypeName psobject -Property ($LicenseAndSolutionsHT = [ordered]@{})
    $DOTNETHT.Clear()
    $DOTNET = New-Object -TypeName psobject -Property ($DOTNETHT = [ordered]@{})
    $VSTOHT.Clear()
    $VSTO = New-Object -TypeName psobject -Property ($VSTOHT = [Ordered]@{})
    $XMLHT.Clear()
    $XML = New-Object -TypeName psobject -Property ($XMLHT = [Ordered]@{})
    $OPMHT.Clear()
    $OPM = New-Object -TypeName psobject -Property ($OPMHT = [Ordered]@{})
    $LocalMachineDetailsHT.Clear()
    $LocalMachineDetails = New-Object -TypeName psobject -Property ($LocalMachineDetailsHT = [Ordered]@{
        OSInstalled = "Not Found"
        OSBitness = "86"
        OS64Bitness = $false
        })
    }

    #functions for output text colour, yellow for informative, green for good, red for bad
    function Receive-Output-G{
        process { Write-Host $_ -ForegroundColor Green }
    }
    function Receive-Output-R{
        process { Write-Host $_ -ForegroundColor Red }
    }
    function Receive-Output-Y{
        process { Write-Host $_ -ForegroundColor Yellow }
    }
    function Receive-Output-W{
        process { Write-Host $_ -ForegroundColor White}
    }

                                                                                function CheckOSandBitness{
   #Display what the OS of this computer is
   Write-Output "Checking OS and Bitness..." | Receive-Output-Y
    Try{
        $env:Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $tempstring = Get-ItemPropertyValue -path $env:Path -name ProductName
        If($VerboseMode){Write-Output "This workstation is running $tempstring."}
        $LocalMachineDetails.OSInstalled = $tempstring.ToSTring()
    }
    Catch{
        Write-Verbose "Catching unexpected Error: $_" -verbose
    }

    #set string for use later
    If([System.Environment]::Is64BitOperatingSystem){
        $LocalMachineDetails.OSBitness = "64"
        $LocalMachineDetails.OS64Bitness = $true
    }
    $tempstring = $LocalMachineDetails.OSBitness
    If($VerboseMode){Write-Output "It's bitness is $tempstring" | Receive-Output-W}

    Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R  
    }

    function CheckOfficeVersions{
        #Check for the currently installed version of Office using the CurVer key, and check Bitness of Outlook for the Bitness of the install
        Write-Output "Checking Office version Installed...`n" | Receive-Output-Y

        $OfficeVer = (Get-ItemProperty HKLM:\SOFTWARE\Classes\Outlook.Application\CurVer)."(default)".Replace("Outlook.Application.", "")

        $OfficeInfoHT.Add("Version", $OfficeVer)
        Try{
            $OfficeBitness = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\$OfficeVer.0\Outlook")."Bitness"
            Write-Output "$OfficeBitness" | Receive-Output-W
            $OfficeInfoHT.Add("Bitness", $OfficeBitness.Replace("x", ""))
            If($VerboseMode){Write-Output "Office $OfficeVer.0 $OfficeBitness installed." | Receive-Output-G}
            }
        Catch [System.Management.Automation.ItemNotFoundException]{
            If($LocalMachineDetails.OSBitness -eq "64"){
                Try{
                    #if the following key exists, assume Office_32 is installed and record it
                    $OfficeBitness = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$OfficeVer.0\Outlook")."Bitness"
                    $OfficeInfoHT.Add("Bitness", $OfficeBitness.Replace("x", ""))
                    If($VerboseMode){Write-Output "Office $OfficeVer.0 $OfficeBitness installed." | Receive-Output-G}
                    }
                Catch [System.Management.Automation.ItemNotFoundException]{
                    If($VerboseMode){Write-Output "$_" | Receive-Output-R}
                    }
                Catch{
                    Write-Verbose "Catching unexpected Error: $_" -verbose
                    }
                }
            }
        Catch [System.Management.Automation.ErrorRecord]{
            If($LocalMachineDetails.OSBitness -eq "64"){
                Try{
                    #if the following key exists, assume Office_32 is installed and record it
                    $OfficeBitness = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$OfficeVer.0\Outlook")."Bitness"
                    $OfficeInfoHT.Add("Bitness", $OfficeBitness.Replace("x", ""))
                    If($VerboseMode){Write-Output "Office $OfficeVer.0 $OfficeBitness installed." | Receive-Output-G}
                    }
                Catch [System.Management.Automation.ItemNotFoundException]{
                    If($VerboseMode){Write-Output "$_" | Receive-Output-R}
                    }
                Catch{
                    Write-Verbose "Catching unexpected Error: $_" -verbose
                    }
                }
            }
        Catch{
            Write-Verbose "Catching unexpected Error: $_" -verbose
            }
        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
    }

    function CheckLoadBehaviors{
        Write-Output "Checking HKCU and HKLM for OnePlaceMail Addin configurations...`n" | Receive-Output-Y
        $temproot = "HKCU"
        $tempkeys32 = @{}
        $tempkeys64 = @{}

        for($i = 0; $i -le 1; $i++){

            $env:Path = $temproot + ":\SOFTWARE\Microsoft\Office\Outlook\Addins\OnePlaceMail" 
            Try{
                $tempobject = Get-ItemProperty $env:Path -name "LoadBehavior"
                $lb = $tempobject.LoadBehavior
                If($VerboseMode -and ($lb -eq 3)){Write-Output "LoadBehavior is at $env:Path and it's value is $lb" | Receive-Output-G}
                ElseIf($VerboseMode){
                    Write-Output "LoadBehavior is at $env:Path and it's value is $lb" | Receive-Output-R
                    }
                If($LocalMachineDetails.OSBitness -eq "64"){
                    $tempkeys64.Add($temproot, $lb)
                    }
                Else{
                    $tempkeys32.Add($temproot, $lb)
                    }
            
                }
            Catch [System.Management.Automation.ItemNotFoundException]{
                If($VerboseMode){Write-Output "$_" | Receive-Output-Y}
                }
            Catch{
                Write-Verbose "Catching unexpected Error: $_" -verbose
                }
            If($LocalMachineDetails.OSBitness -eq "64"){
                $env:Path = $temproot + ":\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\OnePlaceMail"
                Try{
                    $tempobject = Get-ItemProperty $env:Path -name "LoadBehavior"
                    $lb = $tempobject.LoadBehavior
                    If($VerboseMode -and ($lb -eq 3)){Write-Output "LoadBehavior is at $env:Path and it's value is $lb" | Receive-Output-G}
                    ElseIf($VerboseMode){
                        Write-Output "LoadBehavior is at $env:Path and it's value is $lb" | Receive-Output-R
                        }
                    $tempkeys32.Add($temproot, $lb)
                    }
                Catch [System.Management.Automation.ItemNotFoundException]{
                    If($VerboseMode){Write-Output "$_" | Receive-Output-Y}
                    }
                Catch{
                    Write-Verbose "Catching unexpected Error: $_" -verbose
                    }
                }
            $temproot = "HKLM"
            }
        If($tempkeys32.Count -ne 0){$LoadBehaviorHT.Add("keys32", $tempkeys32)}
        If($tempkeys64.Count -ne 0){$LoadBehaviorHT.Add("keys64", $tempkeys64)}
    
        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
    }

    function CheckResiliencyKeys{
        Write-Output "Checking HKCU for Resiliency keys...`n" | Receive-Output-Y


        Try{
            $env:Path = "HKCU:\Software\Microsoft\Office\" + $OfficeInfoHT.Version +".0\Outlook\Resiliency\DisabledItems"
            $Regkey = Get-Item $env:Path

            $Regkey |
                Select-Object -ExpandProperty Property |
                ForEach-Object{
                    $keyrawvalue = (Get-ItemProperty -Path $env:Path -Name $_).$_
                    $keystringvalue = [System.Text.Encoding]::Unicode.GetString($keyrawvalue)
                    If($VerboseMode){Get-Item $env:Path}
                    If($keystringvalue -match "oneplacemail2013.dll"){
                        Write-Output "`noneplacemail2013.dll is under DisabledItems. Property name: $_`n" | Receive-Output-R
                        $ResiliencyHT.Add("Disabled", $true)
                        }
                    }
            }
        Catch [System.Management.Automation.ItemNotFoundException]{
            If($VerboseMode){Write-Output "$_" | Receive-Output-G}
            }
        Catch{
            Write-Verbose "Catching unexpected Error: $_" -verbose
            }

        Try{
            $env:Path = "HKCU:\Software\Microsoft\Office\" + $OfficeInfoHT.Version +".0\Outlook\Resiliency\DoNotDisableAddinList"
            If($VerboseMode){Get-Item $env:Path}
            $Regkey = Get-ItemProperty -path $env:Path -name "OnePlaceMail"
            If($Regkey.OnePlaceMail -eq 1){
                $ResiliencyHT.Add("DoNotDisable", $true)
                If($VerboseMode){Write-Output "`nOnePlaceMail is on the DoNotDisableAddinList." | Receive-Output-G}
                }
            }
        Catch [System.Management.Automation.ItemNotFoundException]{
            If($VerboseMode){Write-Output "$_" | Receive-Output-R}
            }
        Catch [System.Management.Automation.PSArgumentException]{
            If($VerboseMode){Write-Output "$_" | Receive-Output-R}
            }
        Catch{
            Write-Verbose "Catching unexpected Error: $_" -verbose
            }

        Try{
            $env:Path = "HKCU:\Software\Microsoft\Office\" + $OfficeInfoHT.Version +".0\Outlook\Options"
            If($VerboseMode){Get-Item $env:Path}
            $Regkey = Get-ItemProperty -path $env:Path -name "RenderForMonitorDpi"
            If($Regkey.RenderForMonitorDpi -eq 1){
                $ResiliencyHT.Add("OptimizeForCompatibility", $false)
                If($VerboseMode){Write-Output "`nOutlook UI not set to Optimize for compatibility" | Receive-Output-R}
                }
            Else{
                $ResiliencyHT.Add("OptimizeForCompatibility", $true)
                If($VerboseMode){Write-Output "`nOutlook UI set to Optimize for compatibility" | Receive-Output-G}
                }
            }
        Catch [System.Management.Automation.ItemNotFoundException]{
            If($VerboseMode){Write-Output "$_" | Receive-Output-R}
            }
        Catch [System.Management.Automation.PSArgumentException]{
            If($VerboseMode){Write-Output "$_" | Receive-Output-R}
            }
        Catch{
            Write-Verbose "Catching unexpected Error: $_" -verbose
            }

        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
    }

    function CheckOPMSoftwareKey{
        #Checking our registry entry containing items such as licenceLocation, settingsUrl, what product version we are running and what version of Office we are /expecting/ to run on. 
        #Note that outlookVersion will return 2016 for 2016/2019/365.
        Write-Output "Checking if 'OnePlace Solutions' key exists under 'HKLM\Software\' or 'HKLM\Software\WOW6432Node\'..." | Receive-Output-Y

        Try
        {
            $env:Path = "HKLM:\SOFTWARE\OnePlace Solutions"
            $tempobject = get-childitem -path $env:Path
            If($VerboseMode){get-childitem -path $env:Path}
                
            $tempstring = "OnePlace Solutions key exists under 'HKLM\Software\'."
            If($VerboseMode){Write-Output "`n$tempstring" | Receive-Output-G}
            $OPMHT.Add("regOPM","$tempstring")
            $LicenseAndSolutionsHT.Add("Registry", @{})

            #grab licenselocation and settingsUrl while we are here
            $tempobject = Get-ItemProperty "$env:Path\Common" -name "licenseLocation"
            If($tempobject.licenseLocation -match "http"){$LicenseAndSolutionsHT.Registry.Add("License", $tempobject.licenseLocation)}
            $tempobject = Get-ItemProperty "$env:Path\Common" -name "settingsUrl"
            If($tempobject.settingsUrl -match "http"){$LicenseAndSolutionsHT.Registry.Add("Settings", $tempobject.settingsUrl)}
            If($LicenseAndSolutionsHT.Registry.Count -eq 0){$LicenseAndSolutionsHT.Remove("Registry")}

            CheckOPSVersions $env:Path
        }
        Catch [System.Management.Automation.ItemNotFoundException]
        {
            If($LocalMachineDetails.OS64Bitness)
            {
                Try
                {
                    $env:Path = "HKLM:\SOFTWARE\WOW6432Node\OnePlace Solutions"
                    If($VerboseMode){get-childitem -path $env:Path}
                
                    $tempstring = "OnePlace Solutions key exists under 'HKLM\Software\WOW6432Node\'."
                    If($VerboseMode){Write-Output "`n$tempstring" | Receive-Output-G}
                    $OPMHT.Add("regOPM","$tempstring")
                    $LicenseAndSolutionsHT.Add("Registry", @{})
                    $tempobject = Get-ItemProperty "$env:Path\Common" -name "licenseLocation"
                    If($tempobject.licenseLocation -match "http"){$LicenseAndSolutionsHT.Registry.Add("License", $tempobject.licenseLocation)}
                    $tempobject = Get-ItemProperty "$env:Path\Common" -name "settingsUrl"
                    If($tempobject.settingsUrl -match "http"){$LicenseAndSolutionsHT.Registry.Add("Settings", $tempobject.settingsUrl)}
                    If($LicenseAndSolutionsHT.Registry.Count -eq 0){$LicenseAndSolutionsHT.Remove("Registry")}

                    CheckOPSVersions $env:Path
                }
                Catch [System.Management.Automation.ItemNotFoundException]
                {
                   If($VerboseMode){ write-Output "OnePlace Solutions key does not exist under 'HKLM\Software\' or 'HKLM\Software\WOW6432Node\'. Is it installed?" | Receive-Output-R}
                }
                Catch
                {
                    Write-Verbose "Catching unexpected Error: $_" -verbose
                }
            }
            Else
            {
                If($VerboseMode){ write-Output "OnePlace Solutions key does not exist under 'HKLM\Software\'. Is it installed?" | Receive-Output-R}
            }
        }
        Catch
        {
            Write-Verbose "Catching unexpected Error: $_" -verbose
        }
        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
    }

    function CheckVSTOInstallKeys{
        #Check for VSTO install keys, if 32 bit not there then check for 64 bit
        Write-Output "Checking VSTO Runtime Setup Keys...`n" | Receive-Output-Y

        Try
        {
            $env:Path = "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup"
            If($VerboseMode){get-childitem -path $env:Path}
            $tempstring = "32 bit VSTO install keys exist under 'HKLM\Software\'."
            If($VerboseMode){Write-Output "`n$tempstring" | Receive-Output-G}
            $VSTOHT.Add("Installed",$true)
        }
        Catch [System.Management.Automation.ItemNotFoundException]
        {
        
            If($LocalMachineDetails.OS64Bitness)
            {
                Try
                {
                    $env:Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup"
                    If($VerboseMode){get-childitem -path $env:Path}
                    $tempstring = "64 bit VSTO install keys exist under 'HKLM\Software\WOW6432Node\'."
                    If($VerboseMode){Write-Output "`n$tempstring" | Receive-Output-G}
                    $VSTOHT.Add("Installed",$true)
                }
                Catch [System.Management.Automation.ItemNotFoundException]
                {
                    If($VerboseMode){write-Output "`nVSTO install keys do not exist under 'HKLM\Software\' or 'HKLM\Software\WOW6432Node\'. Is it installed?" | Receive-Output-R}
                }
                Catch
                {
                    Write-Verbose "Catching unexpected Error: $_" -verbose
                }
            }
            Else
            {
                If($VerboseMode){write-Output "`nVSTO install keys do not exist under 'HKLM\Software\'. Is it installed?" | Receive-Output-R}
            }
        }
        Catch
        {
            Write-Verbose "Catching unexpected Error: $_" -verbose
        }
        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
    }

    function CheckDotNetSolutionMetaData{
        Write-Output "Checking what the path is for the .NET version OnePlaceMail is trying to use...`n" | Receive-Output-Y


        #netpathactual is out here so we can use it to check bitness for .NET later
        $netpathactual = ""
        Try{
            $env:Path = "HKCU:\SOFTWARE\Microsoft\VSTO\SolutionMetaData"
            $temppathpart = ""
            If($LocalMachineDetails.OS64Bitness){$temppathpart = " (x86)"}
            If($VerboseMode){get-itemproperty -path $env:Path -name "file:///C:\Program Files$temppathpart\OnePlace Solutions\OnePlaceMail2013.vsto"}
            $env:Path = $env:Path + "\" + (get-itempropertyvalue -path $env:Path -name "file:///C:\Program Files$temppathpart\OnePlace Solutions\OnePlaceMail2013.vsto".ToString())
    
            Try{
                If($VerboseMode){get-itemproperty -path $env:Path}

                $netpathpointerversion = (get-itempropertyvalue -path $env:Path -name PreferredClr).ToString()
                $netpathpointer32 = "C:\Windows\Microsoft.NET\Framework\" + $netpathpointerversion + "\"
                $netpathpointer64 = "C:\Windows\Microsoft.NET\Framework64\" + $netpathpointerversion + "\"
            
                $tempstring = "OnePlaceMail is trying to use the .NET version corresponding to path ending " + $netpathpointerversion
        
                If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}

                Try{
                    $env:Path = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Client"
                    $netpathactual = get-itempropertyvalue -path $env:Path -name InstallPath

                    if($netpathactual.ToString() -like $netpathpointer32.ToString()){
                        $tempstring = "32 bit .NET should be installed at $netpathactual and OnePlaceMail is pointing to it."
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                        $VSTOHT.Add("HKCUSolutionMetaData" , $tempstring)
                        }
                    else{
                        if(($netpathactual.ToString() -like $netpathpointer64.ToString()) -and $LocalMachineDetails.OS64Bitness){
                            $tempstring = "64 bit .NET should be installed at $netpathactual and OnePlaceMail is pointing to it."
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                            $VSTOHT.Add("HKCUSolutionMetaData" , $tempstring) 
                            }
                        ElseIf($LocalMachineDetails.OS64Bitness){
                            $tempstring = "64 bit .NET should be installed at $netpathactual and OnePlaceMail is expecting $netpathpointer."
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                            $VSTOHT.Add("HKCUSolutionMetaData" , $tempstring)
                            }
                        Else{
                            $tempstring = "32 bit .NET should be installed at $netpathactual and OnePlaceMail is expecting $netpathpointer."
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                            $VSTOHT.Add("HKCUSolutionMetaData" , $tempstring)
                            }
                        }

                    }
                Catch{
                    If($VerboseMode){Write-Output "$_" | Receive-Output-R}
                    }
                }
            Catch{
                If($VerboseMode){
                    Write-Output "$_" | Receive-Output-R
                    Write-Output "Corresponding key pointing to .NET path does not exist. OnePlaceMail should create it on startup of Outlook." | Receive-Output-Y
                    $tempstring = "OnePlaceMail does not have a .NET path set"
                    Write-Output "$tempstring" | Receive-Output-R
                    }
                $OPM.HKCUSolutionMetaData = "$tempstring"
                }
        }
        Catch [System.Management.Automation.ItemNotFoundException]{
            If($VerboseMode){Write-Output "$_" | Receive-Output-R}
            If($VerboseMode){write-Output "OnePlaceMail VSTO addin does not exist under 'HKCU:\SOFTWARE\Microsoft\VSTO\SolutionMetaData'. Is it installed?" | Receive-Output-Y}
            }
        Catch{
            Write-Verbose "Catching unexpected Error: $_" -verbose
        }

        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R

        #Check for DOTNET version  
        CheckDOTNET

        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
    }

    function CheckDOTNET{
        #Check for .NET installed versions and output specific version number v4 is. We don't care about version numbers for any other versions.
        Write-Output "`nChecking what .NET versions are installed...`n" | Receive-Output-Y

        #Checking for DOTNET that matches the OS bitness
        Try{
            $env:Path = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\"
            If($VerboseMode){get-childitem -path $env:Path}

            $osbitnesstostring = "32 bit"
            If($LocalMachineDetails.OS64Bitness){$osbitnesstostring = "64 bit"}
            $versionstring = get-itempropertyvalue -path $env:Path\v4\Full\ -name Version
            $tempstring = $osbitnesstostring + " .NET " + $versionstring + " is installed."
            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}

            If($LocalMachineDetails.OS64Bitness){
                $DOTNETHT.Add("64_bit", $versionstring)
                }
            Else{
                $DOTNETHT.Add("32_bit", $versionstring)
                }
            }
        Catch [System.Management.Automation.ItemNotFoundException]{
            If($VerboseMode){Write-Output ".NET $tempstring is installed." | Receive-Output-R}
            }
        Catch{
            Write-Verbose "Catching unexpected Error: $_" -verbose
            }
        #If OS bitness is 64, check to see 32 bit DOTNET also exists or not
        If($LocalMachineDetails.OS64Bitness){
            Try{
                $env:Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\"
                If($VerboseMode){get-childitem -path $env:Path}

                $osbitnesstostring = "32 bit"
                $versionstring = get-itempropertyvalue -path $env:Path\v4\Full\ -name Version
                $tempstring = $osbitnesstostring + " .NET " + $versionstring + " is installed."
                If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                $DOTNETHT.Add("32_bit", $versionstring)
        
                }
            Catch [System.Management.Automation.ItemNotFoundException]{
                If($VerboseMode){Write-Output "64 bit .NET v4 or above is not installed." | Receive-Output-R}
                }
            Catch{
                Write-Verbose "Catching unexpected Error: $_" -verbose
                }
            }
        }

    function CheckConfigFiles{
        Write-Output "`nChecking license and solutions site URLs in CommonConfig.xml, XML corruption and PST size...`n" | Receive-Output-Y

        Try{
            $env:Path = "$env:APPDATA" + "\OnePlace Solutions\CommonConfig.xml"
            [xml]$xmldoc = Get-Content $env:Path
            $templicenseurl = $xmldoc.configuration.license.licenseLocation
            $tempsolutionsurl = $xmldoc.configuration.settingsUrl.value
            $tempstring1 = "License list located at $templicenseurl."
            $tempstring2 = "Solutions site located at $tempsolutionsurl."
        
            If($VerboseMode){
                Write-output "$tempstring1" | Receive-Output-G
                Write-output "$tempstring2" | Receive-Output-G
                }
            $LicenseAndSolutionsHT.Add("XML" , @{})
            $LicenseAndSolutionsHT.XML.Add("License" , "$templicenseurl")
            $LicenseAndSolutionsHT.XML.Add("Solutions", "$tempsolutionsurl")
            }
        Catch [System.Management.Automation.ItemNotFoundException]{
            if($VerboseMode){Write-Output "Cannot find Common Config xml file to get license and config URLs. Is it corrupt?" | Receive-Output-R}
            }
        Catch{
            'hmm'
            Write-Verbose "Catching unexpected Error: $_" -verbose
            }

        #check the other xml files while we are here, and the size of the PST
        $xmlfiles = @("CommonConfig.xml","DocsConfig.xml","Favorites.xml","last_update.xml","LiveConfig.xml","MailConfig.xml","Recents.xml","settingsupdatesconfig.xml","solutionprofiles.xml","solutionprofilesconfig.xml")
        $xmlerrors = $false
        Foreach($filename in $xmlfiles){
            Try{
                $env:Path = "$env:APPDATA" + "\OnePlace Solutions\$filename"
                [xml]$xmldoc = Get-Content $env:Path
                switch ($filename){
                    MailConfig.xml {
                        $usePST = $xmldoc.configuration.outlook.usePST.value
                        $profilePath = $xmldoc.configuration.outlook.usePST.profilePath
                        #$xmldoc.configuration.outlook.usePST | Format-Table
                        if($VerboseMode){Write-Output "Using PST? $usePST; Where is it if we are? $profilePath" | Receive-Output-G}
                        $XMLHT.Add("usePST", $usePST)
                        $XMLHT.Add("profilePath", $profilePath)

                        if($usePST -and ($profilePath -eq "roaming")){
                            $env:Path = "$env:APPDATA" + "\OnePlace Solutions\oneplacemail.pst"
                            if((Get-Item $env:Path).length -gt 300kb){
                                $tempstring = "Abnormally large 'oneplacemail.pst'. Recovered Items likely present"
                                $XMLHT.Add("sizePST", $tempstring)
                                if($VerboseMode){Write-Output "$tempstring" | Receive-Output-Y}
                                }
                            }
                        elseif($usePST -and ($profilePath -eq "local")){
                            $env:Path = "$env:LOCALAPPDATA" + "\OnePlace Solutions\oneplacemail.pst"
                            if((Get-Item $env:Path).length -gt 300kb){
                                $tempstring = "Abnormally large 'oneplacemail.pst'. Recovered Items likely present"
                                $XMLHT.Add("sizePST", $tempstring)
                                if($VerboseMode){Write-Output "$tempstring" | Receive-Output-Y}
                                }
                            }
                        }
                    }
                }
            Catch [System.Management.Automation.ItemNotFoundException]{
                if($VerboseMode){Write-Output "Cannot find $filename xml file. Is it corrupt?" | Receive-Output-R}
                $XMLHT.Add("$filename", "Missing")
                $xmlerrors = $true
                }
            Catch [System.Management.Automation.RuntimeException]{
                if($VerboseMode){Write-Output "$filename xml file is corrupt or missing" | Receive-Output-R}
                $XMLHT.Add("$filename", "Corrupt/Missing")
                $xmlerrors = $true
                }
            Catch{
                Write-Verbose "Catching unexpected Error: $_" -verbose
                }
            }
        if($xmlerrors){
            if($verboseMode){Write-Output "Some XML corruption may be present. Recommend deleting offending files"| Receive-Output-R}
            }
        else{$XMLHT.Add("No XML Issues Found",$True)}


    

        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
    }

    function CheckOPSVersions([string]$arg1){
        $env:path = $arg1
        $OPMVersion = ""
        $OPDVersion = ""
        $OPLVersion = ""
        $OPSCheckBuild = ""

        Try{
            $OPMVersion = Get-ItemProperty "$env:Path\OnePlaceMail" -name "productVersion"
            $OPMVersion = $OPMVersion.productVersion
            $OPMHT.Add("OnePlaceMailBuild", $OPMVersion)
            $OPSCheckBuild = $OPMVersion
            Try{
                $tempobject = Get-ItemProperty "$env:Path\OnePlaceMail" -name "outlookVersion"
                $officeBuild = ""
                switch ($tempobject.outlookversion){
                    '2010'{
                        $officeBuild = "14.0"
                        }
                    '2013'{
                        $officeBuild = "15.0"
                        }
                    '2016'{
                        $officeBuild = "16.0"
                        }
                }

                If(($OfficeInfoHT.Version + ".0")-eq $officeBuild){
                    $tempstring = $officeBuild
                    $OPMHT.Add("OnePlaceMailForOutlook", $tempstring)
                    }
                Else{
                    $tempstring = ("OnePlaceMail for Outlook " + $officeBuild + ".0 installed, expecting " + $OfficeInfoHT.Version + ".0")
                    If($VerboseMode){Write-Output $tempstring | Receive-Output-R}
                    $OPMHT.Add("OnePlaceMailForOutlook", $tempstring)
                    }
                }
            Catch{
                }
            }
        Catch{
            $OPMHT.Add("OnePlaceMailBuild", "Not Installed")
            }
        Try{    
            $OPDVersion = Get-ItemProperty "$env:Path\OnePlaceDocs" -name "productVersion"
            $OPDVersion = $OPDVersion.productVersion
            $OPMHT.Add("OnePlaceDocsBuild", $OPDVersion)
            $OPSCheckBuild = $OPDVersion
            }
        Catch{
            $OPMHT.Add("OnePlaceDocsBuild", "Not Installed")
            }
        Try{
            $OPLVersion = Get-ItemProperty "$env:Path\OnePlaceLive" -name "productVersion"
            $OPLVersion = $OPLVersion.productVersion
            $OPMHT.Add("OnePlaceLiveBuild", $OPLVersion)
            }
        Catch{
            $OPMHT.Add("OnePlaceLiveBuild", "Not Installed")
            }


        Try{
            If($OPSCheckBuild -ne ""){
                ForEach($key in $OPMBuildsHT.Keys){
                    If($OPMBuildsHT[$key] -eq $OPMVersion){
                        $OPMHT.Add("OnePlaceVersion", $key)
                        If($VerboseMode){Write-Output "OPS $key installed" | Receive-Output-Y}
                        break
                        }
                    }
                }
            }
        Catch{}
        }
    function CheckTypeLib{
    
        Write-Output "Checking TypeLib entries for Office Object libraries...`n" | Receive-Output-Y
        Try{
            $env:Path = "HKLM:\Software\Classes\Interface\{000C03A7-0000-0000-C000-000000000046}\TypeLib"
            $typeLibVerExpected = Get-ItemProperty $env:Path -name "Version"
            $typeLibVerExpected = $typeLibVerExpected.Version

            If($VerboseMode){Write-Output "TypeLib we are expecting to use is $typeLibVerExpected." | Receive-Output-Y}

            switch ($typeLibVerExpected){
                '2.4'{
                    $tempstring = ""
                    If($OfficeInfoHT.Version -eq 12){
                        $tempstring = "TypeLib", $typeLibVerExpected + " for Office 2007"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                        }
                    Else{
                        $tempstring = "TypeLib mismatch! " + $_ + " for Office 12.0 referenced but we are running Office " + $OfficeInfoHT.Version + ".0"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                        }
                    $DOTNETHT.Add("typelib_office", $tempstring)
                    }
                '2.5'{
                    $tempstring = ""
                    If($OfficeInfoHT.Version -eq 14){
                        $tempstring = "TypeLib", $typeLibVerExpected + " for Office 2010"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                        }
                    Else{
                        $tempstring = "TypeLib mismatch! " + $_ + " for Office 14.0 referenced but we are running Office " + $OfficeInfoHT.Version + ".0"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                        }
                    $DOTNETHT.Add("typelib_office", $tempstring)
                    }
                '2.6'{
                    $tempstring = ""
                    If(($OfficeInfoHT.Version -eq 15) -or ($OfficeInfoHT.Version -eq 16)){
                        $tempstring = "TypeLib " + $typeLibVerExpected + " for Office 2013/2016/2019"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                        }
                    Else{
                       $tempstring = "TypeLib mismatch! " + $_ + " for Office 15.0/16.0 referenced but we are running Office " + $OfficeInfoHT.Version + ".0"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                        }
                    $DOTNETHT.Add("typelib_office", $tempstring)
                    }
                '2.7'{
                    $tempstring = ""
                    If(($OfficeInfoHT.Version -eq 15) -or ($OfficeInfoHT.Version -eq 16)){
                        $tempstring = "TypeLib " + $typeLibVerExpected + " for Office 2013/2016/2019"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                        }
                    Else{
                        $tempstring = "TypeLib mismatch! " + $_ + " for Office 15.0/16.0 referenced but we are running Office " + $OfficeInfoHT.Version + ".0"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                        }
                    $DOTNETHT.Add("typelib_office", $tempstring)
                    }
                '2.8'{
                    $tempstring = ""
                    If($OfficeInfoHT.Version -eq 16){
                        $tempstring = "TypeLib " + $typeLibVerExpected + " for Office 2016/2019"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                        }
                    Else{
                        $tempstring = "TypeLib mismatch! " + $_ + " for Office 16.0 referenced but we are running Office " + $OfficeInfoHT.Version + ".0"
                        If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                        }
                    $DOTNETHT.Add("typelib_office", $tempstring)
                    }
                }
        
            $env:Path = "HKLM:\Software\Classes\TypeLib\{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"
            $typeLibVerAvailable = Get-ChildItem $env:Path
            If($VerboseMode){$typeLibVerAvailable
                Write-Output "Empty entries here should be backed up and removed!" | Receive-Output-Y
                }
            }
        Catch{
            $DOTNETHT.Add("typelib_office", "Not found!")
            }

        Write-Output "Checking TypeLib entries for Outlook Object libraries...`n" | Receive-Output-Y
            Try{
                If($LocalMachineDetails.OS64Bitness){$env:Path = "HKLM:\Software\Classes\WOW6432Node\Interface\{00063001-0000-0000-C000-000000000046}\TypeLib"}
                Else{$env:Path = "HKLM:\Software\Classes\Interface\{00063001-0000-0000-C000-000000000046}\TypeLib"}
                $typeLibVerExpected = Get-ItemProperty $env:Path -name "Version"
                $typeLibVerExpected = $typeLibVerExpected.Version
                If($VerboseMode){Write-Output "TypeLib we are expecting to use is $typeLibVerExpected." | Receive-Output-Y}

                switch ($typeLibVerExpected){
                    '9.4'{
                        $tempstring = ""
                        If($OfficeInfoHT.Version -eq 14){
                            $tempstring = "TypeLib " + $typeLibVerExpected + " for Outlook 2010"
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                            }
                        Else{
                           $tempstring = "TypeLib mismatch! " + $_ + " for Outlook 14.0 referenced but we are running Office " + $OfficeInfoHT.Version + ".0"
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                            }
                        $DOTNETHT.Add("typelib_outlook", $tempstring)
                        }
                    '9.5'{
                        $tempstring = ""
                        If($OfficeInfoHT.Version -eq 15){
                            $tempstring = "TypeLib " + $typeLibVerExpected + " for Outlook 2013"
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                            }
                        Else{
                            $tempstring = "TypeLib mismatch! " + $_ + " for Outlook 15.0 referenced but we are running Office " + $OfficeInfoHT.Version + ".0"
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                            }
                        $DOTNETHT.Add("typelib_outlook", $tempstring)
                        }
                    '9.6'{
                        $tempstring = ""
                        If($OfficeInfoHT.Version -eq 16){
                            $tempstring = "TypeLib " + $typeLibVerExpected + " for Outlook 2016/2019"
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-G}
                            }
                        Else{
                            $tempstring = "TypeLib mismatch! " + $_ + " for Office 16.0 referenced but we are running Office " + $OfficeInfoHT.Version + ".0"
                            If($VerboseMode){Write-Output "$tempstring" | Receive-Output-R}
                            }
                        $DOTNETHT.Add("typelib_outlook", $tempstring)
                        }
                    }
        
                $env:Path = "HKLM:\Software\Classes\TypeLib\{00062FFF-0000-0000-C000-000000000046}"
                $typeLibVerAvailable = Get-ChildItem $env:Path
                If($VerboseMode){$typeLibVerAvailable
                    Write-Output "Empty entries here should be backed up and removed!" | Receive-Output-Y
                    }
                }
            Catch{
                $DOTNETHT.Add("typelib_outlook", "Not found!")
                }
        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
        }

    function Checkstdole{
    
        Write-Output "Checking Classes for valid stdole registry...`n" | Receive-Output-Y
        Try{
            $env:Path = "HKLM:\Software\Classes\CLSID\{0BE35203-8F91-11CE-9DE3-00AA004BB851}\InprocServer32"
            $stdoleAssemblyBoth = Get-ItemProperty $env:Path -name "Assembly"
            $stdoleAssemblyBoth= $stdoleAssemblyBoth.Assembly
            #If($VerboseMode){Get-ItemProperty $env:Path -name "Assembly"}

            $env:Path = "HKLM:\Software\Classes\CLSID\{0BE35204-8F91-11CE-9DE3-00AA004BB851}\InprocServer32"
            $stdoleAssemblyApartment = Get-ItemProperty $env:Path -name "Assembly"
            $stdoleAssemblyApartment =  $stdoleAssemblyApartment.Assembly
            #If($VerboseMode){Get-ItemProperty $env:Path -name "Assembly"}

            If($stdoleAssemblyBoth -eq $stdoleAssemblyApartment){
                If($VerboseMode){Write-Output "stdole found individually" | Receive-Output-G}
                }
            $env:Path = "HKLM:\Software\Classes\Installer\Assemblies\Global"
        
            If($VerboseMode){
                #Get-ItemProperty $env:Path -name "stdole*"
                Write-Output "stdole found in global register" | Receive-Output-G
                }
            $DOTNETHT.Add("stdole", $stdoleAssemblyBoth) 
            }

        Catch{
            $DOTNETHT.Add("stdole", "Not found!")
            If($VerboseMode){Write-Output "stdole not found!" | Receive-Output-R}
            }
        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
        }
    #calls all check functions
    function CheckAll{
        CheckOSandBitness

        CheckOfficeVersions

        CheckLoadBehaviors

        CheckResiliencyKeys

        CheckOPMSoftwareKey

        CheckVSTOInstallKeys

        CheckDotNetSolutionMetaData

        CheckConfigFiles

        CheckTypeLib

        Checkstdole
        }

    function OutputReportSummary{

        Write-Output "Local Machine Details:`n" | Receive-Output-Y
        Write-Output $LocalMachineDetails | Format-List | Out-String | Receive-Output-W
        Write-Output "--------`n" | Receive-Output-R

        Write-Output "Office Details:`n" | Receive-Output-Y
        $OfficeInfoHT.GetEnumerator() | Foreach-Object{
            Write-Output $_.key | Receive-Output-W
            Write-Output $_.Value | Format-List | Out-String | Receive-Output-W
            }
        Write-Output "`n--------`n" | Receive-Output-R

        Write-Output "LoadBehavior Keys:`n" | Receive-Output-Y
        $LoadBehaviorHT.GetEnumerator() | Foreach-Object{
            $_.key | Receive-Output-W
            $_.Value | Format-List | Out-String | Receive-Output-W
            }
        Write-Output "`n--------`n" | Receive-Output-R

        Write-Output "Resiliency Keys:`n" | Receive-Output-Y
        $ResiliencyHT.GetEnumerator() | Foreach-Object{
            $_.key | Receive-Output-W
            $_.Value | Format-List | Out-String | Receive-Output-W
            }
        Write-Output "`n--------`n" | Receive-Output-R

        Write-Output "OPM Keys:`n" | Receive-Output-Y
        $OPMHT.GetEnumerator() | Foreach-Object{
            $_.key | Receive-Output-W
            $_.Value | Format-List | Out-String | Receive-Output-W
            }
        Write-Output "`n--------`n" | Receive-Output-R

        Write-Output "VSTO:`n" | Receive-Output-Y
        $VSTOHT.GetEnumerator() | Foreach-Object{
            $_.key | Receive-Output-W
            $_.Value | Format-List | Out-String | Receive-Output-W
            }
        Write-Output "`n--------`n" | Receive-Output-R

        Write-Output "DOTNET:`n" | Receive-Output-Y
        $DOTNETHT.GetEnumerator() | Foreach-Object{
            $_.key | Receive-Output-W
            $_.Value | Format-List | Out-String | Receive-Output-W
            }
        Write-Output "`n--------`n" | Receive-Output-R

        Write-Output "License and Solutions:`n" | Receive-Output-Y
        $LicenseandSolutionsHT.GetEnumerator() | Foreach-Object{
            $_.key | Receive-Output-W
            $_.Value | Format-List | Out-String | Receive-Output-W
            }
        Write-Output "`n--------`n" | Receive-Output-R

        Write-Output "Config Files:`n" | Receive-Output-Y
        $XMLHT.GetEnumerator() | Foreach-Object{
            $_.key | Receive-Output-W
            $_.Value | Format-List | Out-String | Receive-Output-W
            }
    }

    function ShowMenu 
    { 
        $VerboseMode = $False 
        cls 
        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
        write-output 'Welcome to the OnePlace Solutions Troubleshooting Script.' | Receive-Output-G
        write-output "Please note if you do not run this as the user that operates this computer, HKCU key retrieval will not be accurate. `nRun this as the local user to retrieve that those properties correctly." | Receive-Output-Y

        Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
        Write-Host "1: Basic Report" 
        Write-Host "2: Verbose Report"
        Write-Host "3: Link Latest .NET"
        Write-Host "4: Link VSTO 2010"
        Write-Host "5: Link OnePlace Help"
        Write-Host "6: Link Create Support Ticket"
        Write-Host "Q: Press 'Q' to quit." 
    } 

    do{ 
         ReInitialiseVariables
         ShowMenu 
         $input = Read-Host "Please make a selection" 
         switch ($input) { 
            '1'{ 
                cls
                CheckAll
                OutputReportSummary
                } 
            '2'{ 
                cls 
                $VerboseMode = $True
                Write-output "The following is a key for how to read the verbose report in PowerShell:" | Receive-Output-Y
                Write-output "Green items are good" | Receive-Output-G
                Write-output "Yellow/White items are informational" | Receive-Output-Y
                Write-output "Red items require attention" | Receive-Output-R
                Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
                CheckAll
                OutputReportSummary
                $VerboseMode = $False
                Write-Output "`n--------------------------------------------------------------------------------`n" | Receive-Output-R
                Write-output "Please scroll to top for report detail. If you have ran this through the launcher, the output will be in the text file in the same directory as this script." | Receive-Output-Y
                } 
            '3'{ 
                cls 
                'Opening link to latest .NET...'
                start 'https://dotnet.microsoft.com/download/dotnet-framework'
                }
            '4'{
                cls
                CheckDOTNET
                $baseDOTNETVer = $DOTNETHT[0].Substring(0,1)
                If($baseDOTNETVer -eq "4"){
                    Write-Output 'Opening link to VSTO installers...' | Receive-Output-Y
                    start 'https://www.microsoft.com/en-us/download/details.aspx?id=56961'
                    }
                Else{
                    Write-Output "DOTNET $baseDOTNETVer is installed, version 4 is minimum requirement to cover both VSTO and OnePlace Solutions suite. Please install this before trying to install VSTO" | Receive-Output-R
                    }
                }
            '5'{
                cls
                'Opening link to OnePlace Solutions Help for Desktop...'
                start 'https://www.oneplacesolutions.com/help/'
                }
            '6'{
                cls
                'Opening link to Submit Support Request...'
                start 'https://www.oneplacesolutions.com/premium-support.html'
                }
            'q'{
                return 
                }
            } 
        pause 
        } 
    until($input -eq 'q') 
    }
Catch{}