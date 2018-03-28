<###################################################################################
 # 
 # NOTICE: All information contained herein is, and remains
 # the property of David Wiggs. The intellectual and 
 # technical concepts contained herein are proprietary to  David Wiggs 
 # and may be covered by U.S. and Foreign Patents, patents in process,
 # and are protected by trade secret or copyright law. Dissemination of
 # this information or reproduction of this material is strictly
 # forbidden unless prior written permission is obtained from David Wiggs
 #
 # THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
 # ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 # WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 # DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR
 # ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 # (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 # LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 # ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 # (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 # SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 #
####################################################################################>

#####################################################################################
################                    Begin Get-Help                   ################
<####################################################################################
.SYNOPSIS
ToDo

.DESCRIPTION
ToDo

.PARAMETER
ToDo

.EXAMPLE
>Set-NetworkSecurityRules.ps1

.NOTES
Author     : David Wiggs - dwiggs@westmonroepartners.com 
Requires   : Script to be run from the same directory as the ExcelFunctionalLibrary.ps1 
             and AzureRmFunctionalLibrary.ps1 files

#####################################################################################
################                     End Get-Help                    ################
#####################################################################################>

#####################################################################################
#################                    Begin script                   #################
#####################################################################################

# Remove variables that may be residual in the session 
Remove-Variable * -ErrorAction Ignore

# Script requires AzureRmNetworkSecurityGroups.xlsx
if(-not (Test-Path .\AzureRmNetworkSecurityGroups.xlsx))
{
    Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + ": Cannot find helper file AzureRmNetworkSecurityGroups.xlsx. Please ensure the working directory includes the Excel Workbook. Ending script.")
    exit
}

# Include Azure Rm Function Library
if(Test-Path .\AzureRmFunctionalLibrary.ps1)
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Including AzureRmFunctionlLibrary.ps1.")
    . ".\AzureRmFunctionalLibrary.ps1"
}

else 
{
    Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Cannot find helper file AzureRmFunctionalLibrary.ps1. Please ensure the working directory includes the Excel Workbook. Ending script.")
    exit
}

# Include Excel Function Library
if(Test-Path .\ExcelFunctionalLibrary.ps1)
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Including ExcelFunctionalLibrary.ps1.")
    . ".\ExcelFunctionalLibrary.ps1"
}

else 
{
    Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Cannot find helper file ExcelFunctionalLibrary.ps1. Please ensure the working directory includes the Excel Workbook. Ending script.")
    exit
}

# Login to Azure with a resource manager account
Login-AzureRmAccount | Out-Null
$context = Get-AzureRmContext
Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Using account, $($context.Account).")

# Get current subscriptions 
[array]$subscriptions = Get-AzureRmSubscription

if ($subscriptions.Count -gt 1)
{
    Write-Output ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + ' : Please select a subscription for the network update.')

    # Select appropriate Azure subscription
    # Function SelectAzureRmSubscription included in AzureRmFunctionalLibrary
    $subscription = SelectAzureRmSubscription
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Using subscription, $($subscription.Name).")
}

else
{
    $subscription = Select-AzureRmSubscription -SubscriptionId $($subscriptions[0].Id)
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Using subscription, $($subscription.Subscription.Name).")
}

# Get resource groups to be used
# Import-Excel function included in AzureRmFunctionalLibrary
try
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Importing resource groups from AzureRmNetworkSecurityGroups.xlsx.")
    [array]$resourceGroups = Import-Excel -inputObject .\AzureRmNetworkSecurityGroups.xlsx `
                                          -SheetName "Resource Groups" `
                                          -closeExcel `
                                          -ErrorAction Stop
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Imported resource groups from AzureRmNetworkSecurityGroups.xlsx.")
}

catch
{
    Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to import resource groups from AzureRmNetworkSecurityGroups.xlsx.")
}

# Create listed resource groups if the object is not null
# resourceGroups properties pulled from Import-Excel function
if($resourceGroups -ne $null)
{  
    # Get current resource groups if any
    [array]$currentResourceGroups = Get-AzureRmResourceGroup
    $currentResourceGroups = $currentResourceGroups | Select-Object ResourceGroupName, Location
    
    if ($currentResourceGroups -eq $null)
    {
        # No recouse groups deployed to Azure subscription - it is likely a brand new subscription
        $resourceGroupsToCreate = $resourceGroups

        Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Will now create the resource group(s).")

        foreach ($resourceGroup in $resourceGroupsToCreate)
        {
            Write-Progress -Activity "Creating resource groups.." `
                           -Status "Working on $($resourceGroup.ResourceGroupName)" `
                           -PercentComplete ((($resourceGroupsToCreate.IndexOf($resourceGroup)) / $resourceGroupsToCreate.Count) * 100)
            try
            {
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Creating resource group, $($resourceGroup.ResourceGroupName), in location, $($resourceGroup.Location).")
                New-AzureRmResourceGroup -Name $resourceGroup.ResourceGroupName `
                                         -Location $resourceGroup.Location `
                                         -Force `
                                         -ErrorAction Stop | Out-Null
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Created resource group, $($resourceGroup.ResourceGroupName), in location, $($resourceGroup.Location).")
            }

            catch
            {
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to create resource group, $($resourceGroup.ResourceGroupName), in location, $($resourceGroup.Location).")
            }
        }

        Write-Progress -Activity "Creating resource groups.." `
                       -Status "Done" `
                       -PercentComplete 100 `
                       -Completed
    }

    else
    {    
        # Compare curent resource groups to ones that exist in the resource group inventory
        $compareResourceGroups = Compare-Object $currentResourceGroups $resourceGroups -Property ResourceGroupName, Location -IncludeEqual
        
        # Create variables for existing, extra, and to be deployed resource groups
        foreach ($resourceGroup in $compareResourceGroups)
        {    
            if ($($resourceGroup.SideIndicator) -like '==')
            {
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : The resource group, $($resourceGroup.ResourceGroupName), in location, $($resourceGroup.Location), already exists.")
                [array]$existingResourceGroups += ($resourceGroup | Select-Object ResourceGroupName, Location)
            }

            elseif ($($resourceGroup.SideIndicator) -like '<=')
            {
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : The resource group, $($resourceGroup.ResourceGroupName), in location, $($resourceGroup.Location), exists but is not in the inventory of resource groups in AzureRmNetwork.xlsx.")
                [array]$extraResourceGroups += ($resourceGroup | Select-Object ResourceGroupName, Location)
            }

            elseif ($($resourceGroup.SideIndicator) -like '=>')
            {
                [array]$resourceGroupsToCreate += ($resourceGroup | Select-Object ResourceGroupName, Location)
            }
        }

        foreach ($resourceGroup in $resourceGroupsToCreate)
        {
            Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Will now create the resource group(s).")
            
            Write-Progress -Activity "Creating resource groups.." `
                           -Status "Working on $($resourceGroup.ResourceGroupName)" `
                           -PercentComplete ((($resourceGroupsToCreate.IndexOf($resourceGroup)) / $resourceGroupsToCreate.Count) * 100)
            try
            {
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Creating resource group, $($resourceGroup.ResourceGroupName), in location, $($resourceGroup.Location).")
                New-AzureRmResourceGroup -Name $resourceGroup.ResourceGroupName `
                                         -Location $resourceGroup.Location `
                                         -Force `
                                         -ErrorAction Stop | Out-Null
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Created resource group, $($resourceGroup.ResourceGroupName), in location, $($resourceGroup.Location).")
            }

            catch
            {
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to create resource group, $($resourceGroup.ResourceGroupName), in location, $($resourceGroup.Location).")
            }
        }
    }

    Write-Progress -Activity "Creating resource groups.." `
                   -Status "Done" `
                   -PercentComplete 100 `
                   -Completed
    }

else
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Did not find any resource groups to create.")
}

try
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Importing network security rules from AzureRmNetworkSecurityGroups.xlsx.")
    [array]$networkSecurityRules = Import-Excel -inputObject .\AzureRmNetworkSecurityGroups.xlsx `
                                                -SheetName "Network Security Rules" `
                                                -closeExcel `
                                                -ErrorAction Stop
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Imported network security rules from AzureRmNetworkSecurityGroups.xlsx.")
}

catch
{
    Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to import network security rules from AzureRmNetworkSecurityGroups.xlsx.")
}

if($networkSecurityRules -ne $null)
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + ' : Will now create the network security rule(s).')

    # Initialize empty hash table
    $networkSecurityGroups = @{}

    # Build current hash table of current network security groups and rules
    # Initialize empty array
    $currentNetworkSecurityRules = @()
    $currentNetworkSecurityGroups = Get-AzureRmNetworkSecurityGroup
    foreach ($group in $currentNetworkSecurityGroups)
    {
        $currentNetworkSecurityGroupName = $group.Name
        $objNetworkSecurityRules = Get-AzureRmNetworkSecurityRuleConfig -NetworkSecurityGroup $group | Select-Object Name, `
                                                                                                                     Access, `
                                                                                                                     Protocol, `
                                                                                                                     Direction, `
                                                                                                                     Priority, `
                                                                                                                     SourceAddressPrefix, `
                                                                                                                     SourcePortRange, `
                                                                                                                     DestinationAddressPrefix, `
                                                                                                                     DestinationPortRange
        
        foreach ($rule in $objNetworkSecurityRules)
        {
            $rule | Add-Member -MemberType NoteProperty -Name 'networkSecurityGroupName' -Value "$currentNetworkSecurityGroupName"
            $currentNetworkSecurityRules += $rule
        }
    }

    if ($currentNetworkSecurityRules -eq $null)
    {
        $networkSecurityRulesToCreate = $networkSecurityRules

        Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Will now create the network security rule(s).")

        foreach ($rule in $networkSecurityRulesToCreate)
        {
            Write-Progress -Activity "Creating network security rules.." `
                           -Status "Working on $($rule.InputObject.Name)" `
                           -PercentComplete ((($networkSecurityRulesToCreate.IndexOf($rule)) / $networkSecurityRulesToCreate.Count) * 100)
            try
            {
                $ruleConfig = New-AzureRmNetworkSecurityRuleConfig -Name $rule.Name `
                                                                   -Access $rule.Access `
                                                                   -Protocol $rule.Protocol `
                                                                   -Direction $rule.Direction `
                                                                   -Priority $rule.Priority `
                                                                   -SourceAddressPrefix $rule.SourceAddressPrefix `
                                                                   -SourcePortRange $rule.SourcePortRange `
                                                                   -DestinationAddressPrefix $rule.DestinationAddressPrefix `
                                                                   -DestinationPortRange $rule.DestinationPortRange    

                # Build an initial list of network security groups
                if ($networkSecurityGroups.Keys -notcontains $rule.InputObject.networkSecurityGroupName)
                {
                    $ruleList = New-Object System.Collections.Generic.List[Microsoft.Azure.Commands.Network.Models.PSSecurityRule]
                    $ruleList.Add($ruleConfig)
                    $networkSecurityGroups.Add(($rule.networkSecurityGroupName), ($ruleList))
                }

                else
                {
                    $networkSecurityGroups[$rule.networkSecurityGroupName].Add($ruleConfig)
                }      
                
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Created network security rule, $($rule.Name), in network security group, $($rule.networkSecurityGroupName).")
            }

            catch
            {
                $ErrorMessage = $_.Exception.Message
                write-output $ErrorMessage
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to create network security rule, $($rule.Name), in network security group, $($rule.networkSecurityGroupName).")
            }
        }

        Write-Progress -Activity "Creating network security rules.." `
                    -Status "Done" `
                    -PercentComplete 100 `
                    -Completed
    }

    else
    {   
        $compareNetworkSecurityRules = Compare-Object $currentNetworkSecurityRules $networkSecurityRules -IncludeEqual 
        
        # Create variables for existing, extra, and to be deployed resource groups
        foreach ($rule in $compareNetworkSecurityRules)
        {    
            if ($($rule.SideIndicator) -like '==')
            {
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : The network security rule, $($rule.InputObject.name), in network security group, $($rule.InputObject.networkSecurityGroupName), already exists.")
                [array]$existingNetworkSecurityRules += $rule
            }

            elseif ($($rule.SideIndicator) -like '<=')
            {
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : The network security rule, $($rule.InputObject.name), in network security group, $($rule.InputObject.networkSecurityGroupName), exists but is not in the inventory of network security rules in AzureRmNetworkSecurityRules.xlsx.")
                [array]$extraNetworkSecurityRules += $rule
            }

            elseif ($($rule.SideIndicator) -like '=>')
            {
                [array]$networkSecurityRulesToCreate += $rule
            }
        }

        foreach ($rule in $networkSecurityRulesToCreate)
        {
            Write-Progress -Activity "Creating network security rules.." `
                           -Status "Working on $($rule.InputObject.Name)" `
                           -PercentComplete ((($networkSecurityRulesToCreate.IndexOf($rule)) / $networkSecurityRulesToCreate.Count) * 100)
            try
            {
                $ruleConfig = New-AzureRmNetworkSecurityRuleConfig -Name $rule.InputObject.Name `
                                                                   -Access $rule.InputObject.Access `
                                                                   -Protocol $rule.InputObject.Protocol `
                                                                   -Direction $rule.InputObject.Direction `
                                                                   -Priority $rule.InputObject.Priority `
                                                                   -SourceAddressPrefix $rule.InputObject.SourceAddressPrefix `
                                                                   -SourcePortRange $rule.InputObject.SourcePortRange `
                                                                   -DestinationAddressPrefix $rule.InputObject.DestinationAddressPrefix `
                                                                   -DestinationPortRange $rule.InputObject.DestinationPortRange    

                # Build an initial list of network security groups
                if ($networkSecurityGroups.Keys -notcontains $rule.InputObject.networkSecurityGroupName)
                {
                    $ruleList = New-Object System.Collections.Generic.List[Microsoft.Azure.Commands.Network.Models.PSSecurityRule]
                    $ruleList.Add($ruleConfig)
                    $networkSecurityGroups.Add(($rule.InputObject.networkSecurityGroupName), ($ruleList))
                }

                else
                {
                    $networkSecurityGroups[($rule.InputObject.networkSecurityGroupName)].Add($ruleConfig)
                }      
                
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Created network security rule, $($rule.InputObject.Name), in network security group, $($rule.InputObject.networkSecurityGroupName).")
            }

            catch
            {
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to create network security rule, $($rule.InputObject.Name), in network security group, $($rule.InputObject.networkSecurityGroupName).")
            }
        }
    }

    Write-Progress -Activity "Creating network security rules.." `
                -Status "Done" `
                -PercentComplete 100 `
                -Completed
    }

else
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + ' : Did not find any network security rules to create.')
}

try
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Importing network security groups AzureRmNetworkSecurityGroups.xlsx.")
    $nsgs =  Import-Excel -inputObject .\AzureRmNetworkSecurityGroups.xlsx `
                          -SheetName "Network Security Groups" `
                          -closeExcel `
                          -ErrorAction Stop
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Imported network security groups from AzureRmNetworkSecurityGroups.xlsx.")
}

catch
{
    Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to import network security groups from AzureRmNetworkSecurityGroups.xlsx.")
}

# Make sure that the Microsoft.Network namespace is available
$resourceProviders = Get-AzureRmResourceProvider -ListAvailable | Select-Object ProviderNamespace, RegistrationState
$networkNamespace = $resourceProviders | Where-Object {$_.ProviderNamespace -like 'Microsoft.Network'}

if (($networkNamespace.RegistrationState) -like 'NotRegistered')
{
    Register-AzureRmResourceProvider -ProviderNamespace Microsoft.Network

    do 
    {
        $resourceProviders = Get-AzureRmResourceProvider -ListAvailable | Select-Object ProviderNamespace, RegistrationState
        $networkNamespace = $resourceProviders | Where-Object {$_.ProviderNamespace -like 'Microsoft.Network'}
    } 
    until (($networkNamespace).RegistrationState -like 'Registered')
}

if($nsgs -ne $null)
{
    # Initialize array
    $currentNetworkSecurityGroups = @()

    # Get current network security groups
    $currentNetworkSecurityGroups = Get-AzureRmNetworkSecurityGroup    

    if ($currentNetworkSecurityGroups -eq $null)
    {
        $networkSecurityGroupsToCreate = $nsgs

        Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Will now create the network security group(s).")

        foreach ($networkSecurityGroup in $networkSecurityGroupsToCreate)
        {
            Write-Progress -Activity "Creating network security groups.." `
                           -Status "Working on $($networkSecurityGroup.name)" `
                           -PercentComplete ((($networkSecurityGroupsToCreate.IndexOf($networkSecurityGroup)) / $networkSecurityGroupsToCreate.Count) * 100)
            try
            {
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Creating network security group, $($networkSecurityGroup.name), in resource group, $($networkSecurityGroup.resourcegroupName).")
                New-AzureRmNetworkSecurityGroup -ResourceGroupName $networkSecurityGroup.resourceGroupName `
                                                -Location $networkSecurityGroup.location `
                                                -Name $networkSecurityGroup.name `
                                                -SecurityRules ($networkSecurityGroups[($networkSecurityGroup.name)]) `
                                                -WarningAction Ignore | Out-Null
                
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Created network security group, $($networkSecurityGroup.name), in resource group, $($networkSecurityGroup.resourcegroupName).")
            }

            catch
            {
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to create network security group, $($networkSecurityGroup.name), in resource group, $($networkSecurityGroup.resourcegroupName).")
            }
        }

        Write-Progress -Activity "Creating network security groups.." `
                       -Status "Done" `
                       -PercentComplete 100 `
                       -Completed
    }

    elseif ($currentNetworkSecurityGroups.Count -gt 0)
    {    
        $compareNetworkSecurityGroups = Compare-Object $currentNetworkSecurityGroups $nsgs -Property name, resourceGroupName, location -IncludeEqual
        
        # Create variables for existing, extra, and to be deployed resource groups
        foreach ($networkSecurityGroup in $compareNetworkSecurityGroups)
        {    
            if ($($networkSecurityGroup.SideIndicator) -like '==')
            {
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : The network security group, $($networkSecurityGroup.name), in resource group, $($networkSecurityGroup.resourcegroupName), already exists.")
                [array]$existingNetworkSecurityGroups += ($networkSecurityGroup | Select-Object name, resourceGroupName, location)
            }

            elseif ($($networkSecurityGroup.SideIndicator) -like '<=')
            {
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : The network security group, $($networkSecurityGroup.name), in resource group, $($networkSecurityGroup.resourcegroupName), exists but is not in the inventory of network security groups in AzureRmNetworkSecurityGroups.xlsx.")
                [array]$extraNetworkSecurityGroups += ($networkSecurityGroup | Select-Object name, resourceGroupName, location)
            }

            elseif ($($networkSecurityGroup.SideIndicator) -like '=>')
            {
                [array]$networkSecurityGroupsToCreate += ($networkSecurityGroup | Select-Object name, resourceGroupName, location)
            }
        }

        Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Will now update existing network security group(s).")

        foreach ($rule in $networkSecurityRulesToCreate)
        {
            Write-Progress -Activity "Updating network security groups.." `
                        -Status "Working on $($rule.InputObject.networkSecurityGroupname)" `
                        -PercentComplete ((($networkSecurityRulesToCreate.IndexOf($rule)) / $networkSecurityRulesToCreate.Count) * 100)
            
            if ($existingNetworkSecurityGroups.name.Contains(($rule.InputObject.networkSecurityGroupname)))
            {            
                $workingNetworkSecurityGroup = Get-AzureRmNetworkSecurityGroup -Name ($rule.InputObject.networkSecurityGroupname) `
                                                                               -ResourceGroupName ($nsgs | Where-Object {$_.name -like ($rule.InputObject.networkSecurityGroupname)}).resourceGroupName
                
                try 
                {
                    $workingNetworkSecurityGroup | Add-AzureRmNetworkSecurityRuleConfig -Name $rule.InputObject.name `
                                                                                        -Protocol $rule.InputObject.Protocol `
                                                                                        -Access $rule.InputObject.access `
                                                                                        -Direction $rule.InputObject.Direction `
                                                                                        -Priority $rule.InputObject.Priority `
                                                                                        -SourcePortRange $rule.InputObject.SourcePortRange `
                                                                                        -SourceAddressPrefix $rule.InputObject.SourceAddressPrefix `
                                                                                        -DestinationPortRange $rule.InputObject.DestinationPortRange `
                                                                                        -DestinationAddressPrefix $rule.InputObject.DestinationAddressPrefix | Out-Null
                }
                
                catch 
                {
                    Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to create network security rule, $($rule.InputObject.name), in network security group, $($networkSecurityGroup.name).")
                }
            }
            
            $workingNetworkSecurityGroup | Set-AzureRmNetworkSecurityGroup | Out-Null
        }

        Write-Progress -Activity "Updating network security groups.." `
                        -Status "Done" `
                        -PercentComplete 100 `
                        -Completed
    }

    if ($networkSecurityGroupsToCreate.Count -gt 0 -and $currentNetworkSecurityGroups -ne $null)
    {
        Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Will now create new network security group(s).")

        foreach ($networkSecurityGroup in $networkSecurityGroupsToCreate)
        {            
            Write-Progress -Activity "Creating network security groups.." `
                            -Status "Working on $($networkSecurityGroup.name)" `
                            -PercentComplete ((($networkSecurityGroupsToCreate.IndexOf($networkSecurityGroup)) / $networkSecurityGroupsToCreate.Count) * 100)
            try
            {
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " INFO: Creating network security group, $($networkSecurityGroup.name), in resource group, $($networkSecurityGroup.resourcegroupName).")
                New-AzureRmNetworkSecurityGroup -ResourceGroupName $networkSecurityGroup.resourceGroupName`
                                                -Location $networkSecurityGroup.location `
                                                -Name $networkSecurityGroup.name `
                                                -SecurityRules ($networkSecurityGroups[$networkSecurityGroup.name]) `
                                                -WarningAction Ignore | Out-Null
                
                Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Created network security group, $($networkSecurityGroup.name), in resource group, $($networkSecurityGroup.resourcegroupName).")
            }

            catch
            {
                Write-Warning ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + " : Unable to create network security group, $($networkSecurityGroup.name), in resource group, $($networkSecurityGroup.resourcegroupName).")
            }
        }

        Write-Progress -Activity "Creating network security groups.." `
                    -Status "Done" `
                    -PercentComplete 100 `
                    -Completed
    }
}

else
{
    Write-Verbose ((Get-Date -Format MM/dd/yyy_hh:mm:ss) + ' : Did not find any network security groups to create.')
}