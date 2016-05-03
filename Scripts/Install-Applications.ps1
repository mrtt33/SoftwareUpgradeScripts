##########################################################
#                                                        #
#    Area to declare global variables and functions      #
#                                                        #
##########################################################

# Office display names
$officeRoot = "C:\Office2016"
$officeSetup = "$officeRoot\setup.exe"
$officeConfig = "$officeRoot\configuration.xml"

# Adobe display names
$adobeRoot = "C:\FRSCC2016"
$adobeSetup = "C:\FRSCC2016\FRSCC\Build\setup.exe"

# Turn loop on with $true, and loop off with $false
$loop = $false


# function to run start menu for script options
function Build-SelectionMenu{

$global:title = "Install Software"

$global:message = "Select software install option - "
 
$global:all = New-Object System.Management.Automation.Host.ChoiceDescription "&All", `
    "Installs Office and Adobe."
 
$global:onlyAdobe = New-Object System.Management.Automation.Host.ChoiceDescription "Ado&be", `
    "Installs Adobe only."
 
$global:onlyOffice = New-Object System.Management.Automation.Host.ChoiceDescription "Offi&ce", `
    "Installs Office only."

$global:exit = New-Object System.Management.Automation.Host.ChoiceDescription "E&xit", `
    "Exits the script."
 
$global:options = [System.Management.Automation.Host.ChoiceDescription[]]($all, $onlyAdobe, $onlyOffice, $exit)
 
$global:result = $host.ui.PromptForChoice($title, $message, $options, 0)

Run-SelectionMenu $global:result
}

# function with switch to process choice
function Run-SelectionMenu($result){

    switch ($global:result){
            ### Full uninstallation ###
             0 {
                ""
                Write-Host "You have selected All." -ForegroundColor Gray
                ""

                Install-Office2016
                Install-AdobeCC
               }
       
 
            ### Adobe only uninstallation ###
             1 {
                ""
                Write-Host "You have selected Adobe." -ForegroundColor Gray
                ""

                Install-AdobeCC 
               }
       
 
            ### Office only uninstallation ###
             2 {
                ""
                Write-Host "You have selected Office." -ForegroundColor Gray
                ""
                Install-Office2016
               }


            ### User requested uninstallation ###
             3 {
                ""
                Write-Host "Exiting..." -ForegroundColor Gray
                ""
                Exit
               }
    }
}

# function to install Office 2016
function Install-Office2016(){

    If((Test-Path -Path $officeRoot) -eq $true){
        
        Write-Host "Installing Office 365 - version 2016..."
        ""

        # Run setup for Office 2016
        Start-Process "$officeSetup" -ArgumentList "/configure $officeConfig" -Wait -WindowStyle Minimized

        Write-Host "Install has finished."
        ""

        # Add removal on reboot to registry
        Remove-Item -Path $officeRoot -Recurse -Force
        }

    Else{
        
         Write-Host "Error: $officeRoot does not exist." -ForegroundColor Yellow

        }
}

# function to install Adobe CC 2015
function Install-AdobeCC(){

    If((Test-Path -Path $adobeRoot) -eq $true){
        Write-Host "Installing Adobe CC 2015..."
        ""

        # Run setup for Adobe CC
        Start-Process "$adobeSetup" -Wait -WindowStyle Minimized

        Write-Host "Install has finished."
        ""

        # Add removal on reboot to registry
        Remove-Item -Path "$adobeRoot" -Recurse -Force
        }

    Else{

         Write-Host "Error: $adobeRoot does not exist." -ForegroundColor Yellow
        }

}





##########################################################
#                                                        #
#                     Run the script                     #
#                                                        #
##########################################################

cls

Do{

#Build-SelectionMenu

Run-SelectionMenu ($result=0)

}Until($loop -eq $false)   