
#    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
#    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,        
#    INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
#    We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute
#    the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks
#    to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on
#    Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us
#    and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or resultfrom the 
#    use or distribution of the Sample Code.
#    Please note: None of the conditions outlined in the disclaimer above will supersede the terms and conditions contained 
#    within the Premier Customer Services Description.

#Put your app registration info here. You must use application permission with permission Files.Read.All
$clientID = "81195f1c-28f7-4eaa-83d7-baa708a66e40"
$tenantID = "cdcae3ff-a663-4732-9cf5-1e33db81acf1"
$clientSecret = "G9ck~jy-bA.6afAH9w_jziAfnnV5J.46n2"

Function Get-OauthToken{

    Param(
    [Parameter(Mandatory=$true)]$clientID,
    [Parameter(Mandatory=$true)]$tenantID,
    [Parameter(Mandatory=$true)]$clientSecret
)

    $stringUrl = "https://login.microsoftonline.com/" + $tenantId + "/oauth2/v2.0/token/"
    $postData = "client_id=" + $clientId + "&scope=https://graph.microsoft.com/.default&client_secret=" + $clientSecret + "&grant_type=client_credentials"
    try{
        $accessToken = Invoke-RestMethod -Method post -Uri $stringUrl -ContentType "application/x-www-form-urlencoded" -Body $postData -ErrorAction Stop
        return $accessToken
    }
    catch{
        Write-Warning -Message $_.Exception.Message
    }
}

Function Send-GraphRequest{

    Param(
        [Parameter(Mandatory=$true)]$Method,
        [Parameter(Mandatory=$false)]$BearerToken,
        [Parameter(Mandatory=$false)]$Path,
        [Parameter(Mandatory=$false)]$Json,
        [Parameter(Mandatory=$false)]$TenantID
    )

    $Query = "https://graph.microsoft.com/v1.0" + $Path
    $headers = @{
        "Authorization" = "Bearer $($BearerToken)"
    }

    try{
        $graphRequest = Invoke-RestMethod -Method $Method -Headers $headers -Uri $Query -ContentType 'application/json' -Body $json -ErrorAction Stop
        return $graphRequest
    }
    catch{
        return $_.Exception.Message
    }
}

Function Get-OneDriveUserURL{

    param(
        $UserPrincipalName,
        $GroupSiteDisplayName
    )

    $accessToken = Get-OauthToken -clientID $clientID -tenantID $tenantID -clientSecret $clientSecret
    $oneDriveURLRequest = Send-GraphRequest -Method Get -BearerToken $accessToken.access_token -Path "/users/$($userPrincipalName)/drive"
    $groupRequestFilter = '?$filter=displayName eq '
    $groupRequestFilter += "'$($GroupSiteDisplayName)'"
    $groupRequest = Send-GraphRequest -Method Get -BearerToken $accessToken.access_token -Path "/groups$($groupRequestFilter)"
    $groupSiteRequest = Send-GraphRequest -Method Get -BearerToken $accessToken.access_token -Path "/groups/$($groupRequest.value.Id)/sites/root"

    $lastURLDir = $oneDriveURLRequest.webUrl.split("/").split("/")[-1]
    $oneDriveURL = $oneDriveURLRequest.webUrl.Replace($lastURLDir,"")

    $obj = New-Object psobject 
    $obj = @{
        OneDriveOnwer = $oneDriveURLRequest.owner.user.displayName
        OneDriveURL = $oneDriveURL
        GroupSiteDisplayName = $groupSiteRequest.displayName
        GroupSiteURL = $groupSiteRequest.webUrl
    }

    return $obj
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'OneDrive URL finder'
$form.Size = New-Object System.Drawing.Size(800,440)
$form.MinimumSize = New-Object System.Drawing.Size(800,440)
$form.MaximumSize = New-Object System.Drawing.Size(800,440)
$form.StartPosition = 'CenterScreen'
$form.AcceptButton = $okButton

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(280,73)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton_OnClick = {    
    $form.Text = 'Contacting Microsoft Graph...'
    $graphRequest = Get-OneDriveUserURL -UserPrincipalName $textBoxUPN.Text -GroupSiteDisplayName $textBoxGroupSiteDisplayName.Text
    if(!$graphRequest.OneDriveURL){
        $labelResultOwner.Text = $graphRequest.OneDriveOnwer
    }
    else{
        $labelResultOwner.Text = $graphRequest.OneDriveOnwer
        $labelResultOneDriveURL.Text = $graphRequest.OneDriveURL
        $labelGroupSiteDisplayName.Text = $graphRequest.GroupSiteDisplayName
        $labelGroupSiteURL.Text = $graphRequest.GroupSiteURL
    }
    $form.Text = 'OneDrive URL finder'
    $form.Update()
}
$okButton.add_Click($okButton_OnClick)
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(360,73)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$labelUpn = New-Object System.Windows.Forms.Label
$labelUpn.Location = New-Object System.Drawing.Point(10,10)
$labelUpn.Size = New-Object System.Drawing.Size(280,20)
$labelUpn.Text = "Insert user's UPN/e-mail:"
$form.Controls.Add($labelUpn)

$textBoxUPN = New-Object System.Windows.Forms.TextBox
$textBoxUPN.Location = New-Object System.Drawing.Point(10,30)
$textBoxUPN.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBoxUPN)

$labelGroupSiteDisplayName = New-Object System.Windows.Forms.Label
$labelGroupSiteDisplayName.Location = New-Object System.Drawing.Point(10,55)
$labelGroupSiteDisplayName.Size = New-Object System.Drawing.Size(280,20)
$labelGroupSiteDisplayName.Text = "Insert group/Teams site name:"
$form.Controls.Add($labelGroupSiteDisplayName)

$textBoxGroupSiteDisplayName = New-Object System.Windows.Forms.TextBox
$textBoxGroupSiteDisplayName.Location = New-Object System.Drawing.Point(10,75)
$textBoxGroupSiteDisplayName.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBoxGroupSiteDisplayName)


$labelOnwerlabel = New-Object System.Windows.Forms.Label
$labelOnwerlabel.Location = New-Object System.Drawing.Point(10,100)
$labelOnwerlabel.Size = New-Object System.Drawing.Size(85,20)
$labelOnwerlabel.Text = 'Owner:'
$form.Controls.Add($labelOnwerlabel)

$labelResultOwner = New-Object System.Windows.Forms.TextBox
$labelResultOwner.Location = New-Object System.Drawing.Point(10,120)
$labelResultOwner.Size = New-Object System.Drawing.Size(675,20)
$labelResultOwner.ReadOnly = $true
$form.Controls.Add($labelResultOwner)

$labelOneDriveURLlabel = New-Object System.Windows.Forms.Label
$labelOneDriveURLlabel.Location = New-Object System.Drawing.Point(10,150)
$labelOneDriveURLlabel.Size = New-Object System.Drawing.Size(85,20)
$labelOneDriveURLlabel.Text = 'OneDrive URL:'
$form.Controls.Add($labelOneDriveURLlabel)

$labelResultOneDriveURL = New-Object System.Windows.Forms.TextBox
$labelResultOneDriveURL.Location = New-Object System.Drawing.Point(10,170)
$labelResultOneDriveURL.Size = New-Object System.Drawing.Size(675,20)
$labelResultOneDriveURL.ReadOnly = $true
$form.Controls.Add($labelResultOneDriveURL)

#

$labelGroupSiteDisplayNamelabel = New-Object System.Windows.Forms.Label
$labelGroupSiteDisplayNamelabel.Location = New-Object System.Drawing.Point(10,200)
$labelGroupSiteDisplayNamelabel.Size = New-Object System.Drawing.Size(200,20)
$labelGroupSiteDisplayNamelabel.Text = 'Group Site Display Name:'
$form.Controls.Add($labelGroupSiteDisplayNamelabel)

$labelGroupSiteDisplayName = New-Object System.Windows.Forms.TextBox
$labelGroupSiteDisplayName.Location = New-Object System.Drawing.Point(10,220)
$labelGroupSiteDisplayName.Size = New-Object System.Drawing.Size(675,20)
$labelGroupSiteDisplayName.ReadOnly = $true
$form.Controls.Add($labelGroupSiteDisplayName)

$labelGroupSiteURLlabel = New-Object System.Windows.Forms.Label
$labelGroupSiteURLlabel.Location = New-Object System.Drawing.Point(10,250)
$labelGroupSiteURLlabel.Size = New-Object System.Drawing.Size(200,20)
$labelGroupSiteURLlabel.Text = 'Group Site URL:'
$form.Controls.Add($labelGroupSiteURLlabel)

$labelGroupSiteURL = New-Object System.Windows.Forms.TextBox
$labelGroupSiteURL.Location = New-Object System.Drawing.Point(10,270)
$labelGroupSiteURL.Size = New-Object System.Drawing.Size(675,20)
$labelGroupSiteURL.ReadOnly = $true
$form.Controls.Add($labelGroupSiteURL)

#

$form.Topmost = $true

$form.Add_Shown({$textBoxUPN.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $textBoxUPN.Text
    $x
}