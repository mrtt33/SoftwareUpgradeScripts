##########################################################
#                                                        #
#    Area to declare global variables and functions      #
#                                                        #
##########################################################


#### Variable README ####
<#

This is the area to define static software display name strings to be searched for in the registry.

This must match the exact name found in Add/Remove Programs, or the display name parameter in the
Windows\Uninstall registry key.

Example:  $office1 = "Microsoft Office 365 ProPlus - en-us"

#>

############################################################
##              Uninstall Script variables                ##
############################################################
                                                           #
# Office display names                                     #
$office1 = "Microsoft Office 365 ProPlus - en-us"          #
$office2 = "Microsoft Office Professional Plus 2013"       #
$office3 = $null                                           #
                                                           #
                                                           #
# Adobe display names                                      #
$adobeAcrobat = "Adobe Acrobat"                            #
$adobe1 = "FRS Adobe CS6"                                  # 
$adobe2 = "FRSCS6"                                         #
                                                           #
############################################################ 


# Turn loop on with $true, and loop off with $false
$loop = $false

#### Function README ####
<#

Functions are defined here, and run in the order of:

1 --- Build-SelectionMenu
      > This function constructs the menu at the beginning
        of the script. No real needs to modify this until
        the next upgrade.

        Calls Run-SelectionMenu

2 --- Run-SelectionMenu
      > This function process the selection made in the
        menu. If more options are added the script run
        will need to be added here. If more programs are
        added to the static variable list declared at the
        beginning of the script, they will need to be put
        in this section as well.

        Calls Search-FullReg 

3 --- Search-FullReg
      > Short function to run the next to functions in
        succession. Created for ease of writing the 
        script.

        Calls Search-Reg32
        Calls Search-Reg64

4 --- Search-Reg32
      > Function that searches 32-bit installation path
        in the registry. If it finds the applications it
        will extract the uninstall string property and 
        then attempt to uninstall using Remove-MSI.

        Calls Remove-MSI

5 --- Search-Reg64
      > Function that searches 64-bit installation path
        in the registry. If it finds the applications it
        will extract the uninstall string property and 
        then attempt to uninstall using Remove-MSI.

        Calls Remove-MSI

6 --- Remove-MSI
      > Attempts to run MSIExec.exe to uninstall with
        uninstall string found in registry. Verifies that
        the string found is MSI before extracting string
        and running the uninstall command. If the string
        is not an MSI command, passes to Remove-EXE.

        Calls Remove-EXE

7 --- Remove-EXE
      > If application uninstall string is not an MSI
        this is run to uninstall the application via it's
        custom uninstall string. 
#>

# function to search 32 bit registry values
function Search-Reg32($software){

$uninstall32 = gci "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | `
                    foreach { gp $_.PSPath } | ? { $_ -match "$software" } | `
                    select UninstallString
 
                If ($uninstall32 -ne $null)
                    {
                     Write-Host ""
                     Write-Host "Found " -ForegroundColor Green -NoNewline; `
                     Write-Host "$software" -ForegroundColor White -NoNewline; `
                     Write-Host " 32-bit uninstall string" -ForegroundColor Green

                     Remove-MSI $uninstall32[0]
                     
                    }
 
                Else
                    {
                     Write-Host ""
                     Write-Host "32-bit uninstall string for " -ForegroundColor Yellow -NoNewline; `
                     Write-Host "$software" -ForegroundColor White -NoNewline; `
                     Write-Host " not found." -ForegroundColor Yellow
                    }

}

# function to search 64 bit registry values
function Search-Reg64($software){

$uninstall64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | `
                   foreach { gp $_.PSPath } | ? { $_ -match "$software" } | `
                   select UninstallString
 
                If ($uninstall64 -ne $null)
                    {
                     Write-Host ""
                     Write-Host "Found 64-bit " -ForegroundColor Green -NoNewline; `
                     Write-Host "$software" -ForegroundColor White -NoNewline; `
                     Write-Host " uninstall string" -ForegroundColor Green

                     Remove-MSI $uninstall64[0]
                     
                    }
 
                Else
                    {
                     Write-Host ""
                     Write-Host "64-bit uninstall string for " -ForegroundColor Yellow -NoNewline; `
                     Write-Host "$software" -ForegroundColor White -NoNewline; `
                     Write-Host " not found." -ForegroundColor Yellow
                    }

}

# function running the previous two in sequence
function Search-FullReg($software){

Search-Reg32 $software
Search-Reg64 $software

}

# function to remove files installed via MSI
function Remove-MSI($uninstall){

    If($uninstall -imatch "MSI")
    {
         
         Try{
             ""
             Write-Host "Extracting string..." -ForegroundColor Gray
             $uninstall = $uninstall.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
             $uninstall = $uninstall.Trim()

             Try{
                 Write-Host ""
                 Write-Host "Uninstalling..."
                 Start-Process "msiexec.exe" -ArgumentList "/X $uninstall /qb" -Wait -WindowStyle Minimized
             }

             Catch{
                   Write-Host ""
                   Write-Host "Error occurred during the uninstall process"
                   Write-Host "String used: $uninstall"
             }


         }
         Catch{
               ""
               Write-Host "Error extracting string from registry." -ForegroundColor Red
         }
    }

    Else{
         Remove-EXE
    }
}

# function to remove files installed via EXE
function Remove-EXE($software){

    $uninstall = $uninstall.UninstallString
    $uninstall = $uninstall.Trim()

    Try{
         ""
         Write "Uninstalling..."
         Start-Process "cmd.exe" -ArgumentList "/c `"$uninstall`"" -Wait -WindowStyle Minimized
    }

    Catch{
          ""
          Write-Host "Error occured during the uninstall process" -ForegroundColor Red
          Write-Host "String used: ", "$uninstall" -ForegroundColor Red,White
    }
}

# function to run start menu for script options
function Build-SelectionMenu{

$global:title = "Uninstall Software"

$message = "Select software removal option for removing Office and Adobe - "
 
$all = New-Object System.Management.Automation.Host.ChoiceDescription "&All", `
    "Removes Office and Adobe."
 
$onlyAdobe = New-Object System.Management.Automation.Host.ChoiceDescription "Ado&be", `
    "Removes Adobe only."
 
$onlyOffice = New-Object System.Management.Automation.Host.ChoiceDescription "Offi&ce", `
    "Removes Office only."

$other = New-Object System.Management.Automation.Host.ChoiceDescription "User &Defined", `
    "Removes user specified software."

$search = New-Object System.Management.Automation.Host.ChoiceDescription "S&earch Only", `
    "Removes user specified software."

$exit = New-Object System.Management.Automation.Host.ChoiceDescription "E&xit", `
    "Exits software uninstaller."
 
$options = [System.Management.Automation.Host.ChoiceDescription[]]($all, $onlyAdobe, $onlyOffice, $other, $search, $exit)
 
$global:result = $host.ui.PromptForChoice($title, $message, $options, 0)

Run-SelectionMenu $global:result
}

# function with switch to process choice
function Run-SelectionMenu($selection){

    switch ($global:result){
            ### Full uninstallation ###
             0 {
                ""
                Write-Host "You have selected All." -ForegroundColor Gray
                ""

                Search-FullReg $office1
                Search-FullReg $office2
                Search-FullReg $adobeAcrobat
                Search-FullReg $adobe1
                Search-FullReg $adobe2 
               }
       
 
            ### Adobe only uninstallation ###
             1 {
                ""
                Write-Host "You have selected Adobe." -ForegroundColor Gray
                ""

                Search-FullReg $adobeAcrobat
                Search-FullReg $adobe1
                Search-FullReg $adobe2    
               }
       
 
            ### Office only uninstallation ###
             2 {
                ""
                Write-Host "You have selected Office." -ForegroundColor Gray
                ""
                Search-FullReg $office1
               }
       

            ### User defined software removal ###
             3 {
                ""
                Write-Host "You have selected Other." -ForegroundColor Gray
                ""
                [string]$software = Read-Host "Type software name"
                Search-FullReg $software
               }
                
            ### Only query the registry ###
             4 {
                ""
                Write-Host "You have selected Search." -ForegroundColor Gray
                ""
                [string]$software = Read-Host "Type software name"
                Search-NoUninstallAll $software
               }

            ### User requested uninstallation ###
             5 {
                ""
                Write-Host "Cancel..." -ForegroundColor Gray
                ""
                Break
               }
    }
}

# function to search 32 bit registry with no uninstall
function Search-NoUninstall32($software){

$uninstall32 = gci "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | `
                    foreach { gp $_.PSPath } | ? { $_ -match "$software" } | `
                    select DisplayName,UninstallString
 
                If ($uninstall32 -ne $null)
                    {
                     Write-Host ""
                     Write-Host "Found " -ForegroundColor Green -NoNewline; `
                     Write-Host "$software" -ForegroundColor White -NoNewline; `
                     Write-Host " 32-bit uninstall string" -ForegroundColor Green

                     foreach($child in $uninstall32){
                         ""
                         Write-Host $child
                     }
                     
                    }
 
                Else
                    {
                     Write-Host ""
                     Write-Host "32-bit uninstall string for " -ForegroundColor Yellow -NoNewline; `
                     Write-Host "$software" -ForegroundColor White -NoNewline; `
                     Write-Host " not found." -ForegroundColor Yellow
                    }

}

# function to search 32 bit registry with no uninstall
function Search-NoUninstall64($software){

$uninstall64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | `
                   foreach { gp $_.PSPath } | ? { $_ -match "$software" } | `
                   select DisplayName,UninstallString
 
                If ($uninstall64 -ne $null)
                    {
                     Write-Host ""
                     Write-Host "Found 64-bit " -ForegroundColor Green -NoNewline; `
                     Write-Host "$software" -ForegroundColor White -NoNewline; `
                     Write-Host " uninstall string" -ForegroundColor Green

                     foreach($child in $uninstall64){
                         ""
                         Write-Host $child
                     }
                     
                    }
 
                Else
                    {
                     Write-Host ""
                     Write-Host "64-bit uninstall string for " -ForegroundColor Yellow -NoNewline; `
                     Write-Host "$software" -ForegroundColor White -NoNewline; `
                     Write-Host " not found." -ForegroundColor Yellow
                    }

}

# function to call both searches above
function Search-NoUninstallAll($software){

Search-NoUninstall32 $software
Search-NoUninstall64 $software

}


##########################################################
#                                                        #
#                     Run the script                     #
#                                                        #
##########################################################


cls

Do{

#Build-SelectionMenu

Run-SelectionMenu ($result = 0)

}Until($loop -eq $false)   