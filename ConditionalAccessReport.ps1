#region Description
<#     
       .NOTES
       ==============================================================================
       Created on:         2025/04/14 
       Created by:         Drago Petrovic
       Organization:       MSB365.blog
       Filename:           ConditionalAccessReport.ps1
       Current version:    V1.0     

       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ==============================================================================

.DESCRIPTION
	Automate Microsoft 365 User Offboarding with PowerShell           
       

.SYNOPSIS
    Azure Conditional Access Policy Management Tool
.DESCRIPTION
    This script provides options to:
    1. Create an HTML report of Conditional Access policies
    2. Create a CSV report of Conditional Access policies
    3. Export Conditional Access policies for backup/migration
    4. Import Conditional Access policies from a file
.NOTES
    Requires Microsoft Graph PowerShell modules

.EXAMPLE
    .\ConditionalAccessReport.ps1
             

.COPYRIGHT
    Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
    to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
    and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
    WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    ===========================================================================
    
.CHANGE LOG
    V1.00, 2025/04/14 - DrPe - Initial version

             
			 




--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "Conditional Access Report"
$RKEY = "MSB365_ConditionalAccessReport"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2025 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################
write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
###############################################################################


# Check if required modules are installed
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Identity.SignIns")
$modulesToInstall = @()

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $modulesToInstall += $module
    }
}

if ($modulesToInstall.Count -gt 0) {
    Write-Host "Installing required modules: $($modulesToInstall -join ', ')" -ForegroundColor Yellow
    foreach ($module in $modulesToInstall) {
        Install-Module -Name $module -Scope CurrentUser -Force
    }
}

# Import required modules
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.SignIns

# Function to connect to Microsoft Graph
function Connect-ToMsGraph {
    try {
        # Check if already connected
        $graphContext = Get-MgContext
        if ($null -eq $graphContext) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes "Policy.Read.All", "Policy.ReadWrite.ConditionalAccess", "Directory.Read.All", "Directory.ReadWrite.All"
        }
        else {
            Write-Host "Already connected to Microsoft Graph as $($graphContext.Account)" -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
        return $false
    }
}

# Function to get all Conditional Access policies
function Get-AllConditionalAccessPolicies {
    try {
        $policies = Get-MgIdentityConditionalAccessPolicy
        return $policies
    }
    catch {
        Write-Host "Error retrieving Conditional Access policies: $_" -ForegroundColor Red
        return $null
    }
}

# Function to create HTML report
function Create-HtmlReport {
    $outputPath = Join-Path -Path $PWD -ChildPath "ConditionalAccessReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    if (Connect-ToMsGraph) {
        $policies = Get-AllConditionalAccessPolicies
        
        if ($null -ne $policies) {
            # Create HTML header with styling
            $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Azure Conditional Access Policies Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0078D4; }
        h2 { color: #0078D4; margin-top: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th { background-color: #0078D4; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .policy-state-enabled { color: green; font-weight: bold; }
        .policy-state-disabled { color: #FF8C00; }
        .policy-details { margin-left: 20px; }
        .timestamp { color: #666; font-size: 0.8em; }
    </style>
</head>
<body>
    <h1>Azure Conditional Access Policies Report</h1>
    <p class="timestamp">Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <h2>Summary</h2>
    <p>Total Policies: $($policies.Count)</p>
    <p>Enabled Policies: $($policies | Where-Object { $_.State -eq "enabled" } | Measure-Object | Select-Object -ExpandProperty Count)</p>
    <p>Disabled Policies: $($policies | Where-Object { $_.State -eq "disabled" } | Measure-Object | Select-Object -ExpandProperty Count)</p>
    
    <h2>Policies</h2>
"@

            # Add each policy to the HTML
            foreach ($policy in $policies) {
                $stateClass = if ($policy.State -eq "enabled") { "policy-state-enabled" } else { "policy-state-disabled" }
                
                $html += @"
    <div class="policy">
        <h3>$($policy.DisplayName) <span class="$stateClass">[$($policy.State)]</span></h3>
        <div class="policy-details">
            <p><strong>ID:</strong> $($policy.Id)</p>
            <p><strong>Created:</strong> $($policy.CreatedDateTime)</p>
            <p><strong>Modified:</strong> $($policy.ModifiedDateTime)</p>
            
            <h4>Conditions</h4>
            <table>
                <tr>
                    <th>Condition</th>
                    <th>Details</th>
                </tr>
"@

                # Users
                $html += @"
                <tr>
                    <td>Users</td>
                    <td>
                        <strong>Include:</strong> $(Format-ConditionValue $policy.Conditions.Users.IncludeUsers)<br>
                        <strong>Exclude:</strong> $(Format-ConditionValue $policy.Conditions.Users.ExcludeUsers)<br>
                        <strong>Include Groups:</strong> $(Format-ConditionValue $policy.Conditions.Users.IncludeGroups)<br>
                        <strong>Exclude Groups:</strong> $(Format-ConditionValue $policy.Conditions.Users.ExcludeGroups)<br>
                        <strong>Include Roles:</strong> $(Format-ConditionValue $policy.Conditions.Users.IncludeRoles)<br>
                        <strong>Exclude Roles:</strong> $(Format-ConditionValue $policy.Conditions.Users.ExcludeRoles)
                    </td>
                </tr>
"@

                # Applications
                $html += @"
                <tr>
                    <td>Applications</td>
                    <td>
                        <strong>Include:</strong> $(Format-ConditionValue $policy.Conditions.Applications.IncludeApplications)<br>
                        <strong>Exclude:</strong> $(Format-ConditionValue $policy.Conditions.Applications.ExcludeApplications)
                    </td>
                </tr>
"@

                # Locations
                $html += @"
                <tr>
                    <td>Locations</td>
                    <td>
                        <strong>Include:</strong> $(Format-ConditionValue $policy.Conditions.Locations.IncludeLocations)<br>
                        <strong>Exclude:</strong> $(Format-ConditionValue $policy.Conditions.Locations.ExcludeLocations)
                    </td>
                </tr>
"@

                # Platforms
                $html += @"
                <tr>
                    <td>Platforms</td>
                    <td>
                        <strong>Include:</strong> $(Format-ConditionValue $policy.Conditions.Platforms.IncludePlatforms)<br>
                        <strong>Exclude:</strong> $(Format-ConditionValue $policy.Conditions.Platforms.ExcludePlatforms)
                    </td>
                </tr>
"@

                # Client Apps
                $html += @"
                <tr>
                    <td>Client Apps</td>
                    <td>
                        <strong>Include:</strong> $(Format-ConditionValue $policy.Conditions.ClientAppTypes)
                    </td>
                </tr>
"@

                # Sign-in Risk
                $html += @"
                <tr>
                    <td>Sign-in Risk</td>
                    <td>
                        <strong>Include:</strong> $(Format-ConditionValue $policy.Conditions.SignInRiskLevels)
                    </td>
                </tr>
"@

                $html += @"
            </table>
            
            <h4>Grant Controls</h4>
"@

                if ($null -ne $policy.GrantControls) {
                    $html += @"
            <p><strong>Operator:</strong> $($policy.GrantControls.Operator)</p>
            <p><strong>Built-in Controls:</strong> $(Format-ConditionValue $policy.GrantControls.BuiltInControls)</p>
            <p><strong>Custom Controls:</strong> $(Format-ConditionValue $policy.GrantControls.CustomAuthenticationFactors)</p>
"@
                }
                else {
                    $html += "<p>No grant controls specified</p>"
                }

                $html += @"
            <h4>Session Controls</h4>
"@

                if ($null -ne $policy.SessionControls) {
                    $html += @"
            <p><strong>Application Enforced Restrictions:</strong> $($policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled)</p>
            <p><strong>Cloud App Security:</strong> $($policy.SessionControls.CloudAppSecurity.IsEnabled)</p>
            <p><strong>Sign-in Frequency:</strong> $($policy.SessionControls.SignInFrequency.IsEnabled)</p>
            <p><strong>Persistent Browser:</strong> $($policy.SessionControls.PersistentBrowser.IsEnabled)</p>
"@
                }
                else {
                    $html += "<p>No session controls specified</p>"
                }

                $html += @"
        </div>
    </div>
"@
            }

            # Close HTML
            $html += @"
</body>
</html>
"@

            # Save HTML to file
            $html | Out-File -FilePath $outputPath -Encoding UTF8
            
            Write-Host "HTML report created successfully at: $outputPath" -ForegroundColor Green
            # Open the report in the default browser
            Start-Process $outputPath
        }
    }
}

# Function to create CSV report
function Create-CsvReport {
    $outputPath = Join-Path -Path $PWD -ChildPath "ConditionalAccessReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if (Connect-ToMsGraph) {
        $policies = Get-AllConditionalAccessPolicies
        
        if ($null -ne $policies) {
            $csvData = @()
            
            foreach ($policy in $policies) {
                $policyData = [PSCustomObject]@{
                    DisplayName = $policy.DisplayName
                    Id = $policy.Id
                    State = $policy.State
                    CreatedDateTime = $policy.CreatedDateTime
                    ModifiedDateTime = $policy.ModifiedDateTime
                    IncludeUsers = ($policy.Conditions.Users.IncludeUsers -join ", ")
                    ExcludeUsers = ($policy.Conditions.Users.ExcludeUsers -join ", ")
                    IncludeGroups = ($policy.Conditions.Users.IncludeGroups -join ", ")
                    ExcludeGroups = ($policy.Conditions.Users.ExcludeGroups -join ", ")
                    IncludeRoles = ($policy.Conditions.Users.IncludeRoles -join ", ")
                    ExcludeRoles = ($policy.Conditions.Users.ExcludeRoles -join ", ")
                    IncludeApplications = ($policy.Conditions.Applications.IncludeApplications -join ", ")
                    ExcludeApplications = ($policy.Conditions.Applications.ExcludeApplications -join ", ")
                    IncludeLocations = ($policy.Conditions.Locations.IncludeLocations -join ", ")
                    ExcludeLocations = ($policy.Conditions.Locations.ExcludeLocations -join ", ")
                    IncludePlatforms = ($policy.Conditions.Platforms.IncludePlatforms -join ", ")
                    ExcludePlatforms = ($policy.Conditions.Platforms.ExcludePlatforms -join ", ")
                    ClientAppTypes = ($policy.Conditions.ClientAppTypes -join ", ")
                    SignInRiskLevels = ($policy.Conditions.SignInRiskLevels -join ", ")
                    GrantControlsOperator = $policy.GrantControls.Operator
                    GrantControlsBuiltIn = ($policy.GrantControls.BuiltInControls -join ", ")
                }
                
                $csvData += $policyData
            }
            
            $csvData | Export-Csv -Path $outputPath -NoTypeInformation
            
            Write-Host "CSV report created successfully at: $outputPath" -ForegroundColor Green
            # Open the folder containing the CSV
            Invoke-Item (Split-Path -Parent $outputPath)
        }
    }
}

# Function to export policies
function Export-ConditionalAccessPolicies {
    $outputPath = Join-Path -Path $PWD -ChildPath "ConditionalAccessPolicies_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    
    if (Connect-ToMsGraph) {
        $policies = Get-AllConditionalAccessPolicies
        
        if ($null -ne $policies) {
            # Convert to JSON and save to file
            $policies | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputPath -Encoding UTF8
            
            Write-Host "Conditional Access policies exported successfully to: $outputPath" -ForegroundColor Green
            # Open the folder containing the export
            Invoke-Item (Split-Path -Parent $outputPath)
        }
    }
}

# Function to import policies
function Import-ConditionalAccessPolicies {
    # Prompt for file selection
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "JSON files (.json)|.json|All files (.)|."
    $openFileDialog.Title = "Select Conditional Access Policies JSON File"
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $filePath = $openFileDialog.FileName
        
        try {
            # Read and parse the JSON file
            $policiesJson = Get-Content -Path $filePath -Raw
            $policies = ConvertFrom-Json -InputObject $policiesJson
            
            if (Connect-ToMsGraph) {
                $importedCount = 0
                $skippedCount = 0
                
                foreach ($policy in $policies) {
                    # Check if policy already exists
                    $existingPolicy = $null
                    try {
                        $existingPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -ErrorAction SilentlyContinue
                    }
                    catch {
                        # Policy doesn't exist, which is fine
                    }
                    
                    # Prepare policy object for import
                    $policyObj = @{
                        DisplayName = $policy.DisplayName
                        State = $policy.State
                        Conditions = $policy.Conditions
                    }
                    
                    if ($null -ne $policy.GrantControls) {
                        $policyObj.GrantControls = $policy.GrantControls
                    }
                    
                    if ($null -ne $policy.SessionControls) {
                        $policyObj.SessionControls = $policy.SessionControls
                    }
                    
                    if ($null -eq $existingPolicy) {
                        # Create new policy
                        try {
                            New-MgIdentityConditionalAccessPolicy -BodyParameter $policyObj
                            Write-Host "Created policy: $($policy.DisplayName)" -ForegroundColor Green
                            $importedCount++
                        }
                        catch {
                            Write-Host "Error creating policy $($policy.DisplayName): $_" -ForegroundColor Red
                        }
                    }
                    else {
                        # Ask if user wants to update existing policy
                        $updateChoice = Read-Host "Policy '$($policy.DisplayName)' already exists. Update it? (Y/N)"
                        
                        if ($updateChoice -eq "Y" -or $updateChoice -eq "y") {
                            try {
                                Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter $policyObj
                                Write-Host "Updated policy: $($policy.DisplayName)" -ForegroundColor Green
                                $importedCount++
                            }
                            catch {
                                Write-Host "Error updating policy $($policy.DisplayName): $_" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "Skipped policy: $($policy.DisplayName)" -ForegroundColor Yellow
                            $skippedCount++
                        }
                    }
                }
                
                Write-Host "Import completed. Imported: $importedCount, Skipped: $skippedCount" -ForegroundColor Cyan
            }
        }
        catch {
            Write-Host "Error importing policies: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Import cancelled" -ForegroundColor Yellow
    }
}

# Helper function to format condition values for HTML display
function Format-ConditionValue {
    param (
        [Parameter(Mandatory = $false)]
        [object]$Value
    )
    
    if ($null -eq $Value -or $Value.Count -eq 0) {
        return "None"
    }
    elseif ($Value -is [array]) {
        return ($Value -join ", ")
    }
    else {
        return $Value
    }
}

# Main menu function
function Show-Menu {
    Clear-Host
    Write-Host "===== Azure Conditional Access Policy Management =====" -ForegroundColor Cyan
    Write-Host "1: Create HTML Report of Conditional Access Policies" -ForegroundColor White
    Write-Host "2: Create CSV Report of Conditional Access Policies" -ForegroundColor White
    Write-Host "3: Export Conditional Access Policies" -ForegroundColor White
    Write-Host "4: Import Conditional Access Policies" -ForegroundColor White
    Write-Host "Q: Quit" -ForegroundColor White
    Write-Host "===================================================" -ForegroundColor Cyan
}

# Main script execution
do {
    Show-Menu
    $choice = Read-Host "Please enter your choice"
    
    switch ($choice) {
        "1" {
            Create-HtmlReport
            Read-Host "Press Enter to continue"
        }
        "2" {
            Create-CsvReport
            Read-Host "Press Enter to continue"
        }
        "3" {
            Export-ConditionalAccessPolicies
            Read-Host "Press Enter to continue"
        }
        "4" {
            Import-ConditionalAccessPolicies
            Read-Host "Press Enter to continue"
        }
        "Q" {
            Write-Host "Exiting..." -ForegroundColor Yellow
        }
        "q" {
            Write-Host "Exiting..." -ForegroundColor Yellow
        }
        default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            Read-Host "Press Enter to continue"
        }
    }
} while ($choice -ne "Q" -and $choice -ne "q")

# Disconnect from Microsoft Graph if connected
$graphContext = Get-MgContext
if ($null -ne $graphContext) {
    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Cyan
}
