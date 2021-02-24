﻿# GUI to Manage Exchange Online PowerShell Based Move Requests .
#
##Requires ExchangeOnline PowerShell module. 
#	Install-Module -Name ExchangeOnlineManagement
#
#
#You will need Mailbox Admin rights.
#Created by Cedric Abrahams - cedric@inobits.com
#
#Version 0.6 2021-02-19
#Bulk options not fully functional 
function Add-Clock {
 $code = { 
    $pattern = '\d{2}:\d{2}:\d{2}'
    do {
      $clock = Get-Date -format 'HH:mm:ss'

      $oldtitle = [system.console]::Title
      if ($oldtitle -match $pattern) {
        $newtitle = $oldtitle -replace $pattern, $clock
      } else {
        $newtitle = "$clock $oldtitle"
      }
      [System.Console]::Title = $newtitle
      Start-Sleep -Seconds 1
    } while ($true)
  }

 $ps = [PowerShell]::Create()
 $null = $ps.AddScript($code)
 $ps.BeginInvoke()
}
Add-clock

#Check for administrative rights
<# If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script in a PowerShell Session as Administrator!"
    Break
} #>

$Dom = "medikredit.co.za"

Write-host "Checking if Exchange Online Management PS Module is installed" -Fore Cyan
$PSM = Get-InstalledModule -Name ExchangeOnlineManagement
If ($PSM) {Write-host "Exchange Online Management PS Module is installed" -Fore Green}
If (!$PSM) {Write-host "Exchange Online Management PS Module is not installed" -Fore Yellow
Sleep -s 2
Install-Module -Name ExchangeOnlineManagement}

Write-host "Import Exchange Online Management PS Module"
Import-Module ExchangeOnlineManagement

$CheckCon = Get-EXOCasMailbox
if (!$CheckCon) {Write-host "Connecting to Exchange Online - Please enter your credentials"  -BackgroundColor Blue
Sleep -s 2
Connect-ExchangeOnline}
if ($CheckCon) {Write-Host "Already connected to Exchange Online PowerShell"  -fore Green}
$CheckCon = Get-EXOCasMailbox
Write-host "Connecting to Exchange Online PowerShell" -fore Green
#Variables
$check = $null

If (!$credential) {
	Write-host "- Please enter your credentials for the on premise Exchange Server"  -BackgroundColor Blue
	$credential = get-credential}


#Menu Options
$NewMoveForm = 'Create New Move Requests'
$CompleteMoveForm = 'Complete Move Request'
$ViewMoveForm = 'View Existing Move Requests'
$RemoveMoveForm = 'Remove Move Requests'

$NewBulkMoveForm = "New Bulk Move Request"
$CompleteBulkMoveForm = "Finalise Bulk Move Requests"
$ViewBulkMoveForm = "View Bulk Moves Status"
$RemoveBulkMoveForm = "Remove Bulk Move Requests"


####Data Import####
$Time= Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "
Incomplete mailbox migration data import started $Time." -fore green
Write-host "This may take some time." -fore yellow
$allUsers = Get-MailUser
$Time= Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Incomplete mailbox migration data import finished $Time." -fore green

$AM = $allusers | measure
Write-host "Incomplete Mailbox count:" $AM.count 
Write-host ""
Write-host "Mailbox Move data import started $Time" -fore Green
Write-host "This may take some time." -fore yellow
$moves = Get-moverequest
$Time= Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Import complete at $time" -fore green
$MM = $moves | measure
Write-host "Current move request count:" $MM.count

############Form Functions Start #######################################

Function ChooseForm {
	Add-Type -AssemblyName System.Windows.Forms
	Write-host "ChooseForm" -Fore Yellow
# Create a new form
$ActionForm                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$ActionForm.ClientSize         = '500,200'
$ActionForm.text               = "Mailbox Move Management"
$ActionForm.BackColor          = "#ffffff"

# Create a Title for our form. We will use a label for it.
$TitleOperationChoice                           = New-Object system.Windows.Forms.Label
$TitleOperationChoice.text                      = "Mailbox Move Management"
$TitleOperationChoice.AutoSize                  = $true
$TitleOperationChoice.width                     = 25
$TitleOperationChoice.height                    = 10
$TitleOperationChoice.location                  = New-Object System.Drawing.Point(20,0)
$TitleOperationChoice.Font                      = 'Microsoft Sans Serif,13'

<# #ClockCode
$ClockText                           = New-Object system.Windows.Forms.Label
$ClockText.text                      = Clockcode
$ClockText.AutoSize                  = $true
$ClockText.width                     = 25
$ClockText.height                    = 10
$ClockText.ForeColor					= "#32a852"
$ClockText.location                  = New-Object System.Drawing.Point(150,0)
$ClockText.Font                      = 'Microsoft Sans Serif,13'
$ActionForm.controls.AddRange(@($ClockText)) #>

#Buttons
$SingleUserBtn                   = New-Object system.Windows.Forms.Button
$SingleUserBtn.BackColor         = "#32a852"
$SingleUserBtn.text              = "Individual User Operations"
$SingleUserBtn.width             = 200
$SingleUserBtn.height            = 30
$SingleUserBtn.location          = New-Object System.Drawing.Point(20,30)
$SingleUserBtn.Font              = 'Microsoft Sans Serif,10'
$SingleUserBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($SingleUserBtn)
$SingleUserBtn.Add_Click({FindUserForm})

#Buttons
$BulkUserBtn                   = New-Object system.Windows.Forms.Button
$BulkUserBtn.BackColor         = "#32a852"
$BulkUserBtn.text              = "Bulk Operations*"
$BulkUserBtn.width             = 200
$BulkUserBtn.height            = 30
$BulkUserBtn.location          = New-Object System.Drawing.Point(240,30)
$BulkUserBtn.Font              = 'Microsoft Sans Serif,10'
$BulkUserBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($BulkUserBtn)
$BulkUserBtn.Add_Click({BulkActionForm})


$NoteText                           = New-Object system.Windows.Forms.Label
$NoteText.text                      = '* Bulk functions under development'
$NoteText.AutoSize                  = $true
$NoteText.width                     = 25
$NoteText.height                    = 10
$NoteText.ForeColor					= "#32a852"
$NoteText.location                  = New-Object System.Drawing.Point(20,80)
$NoteText.Font                      = 'Microsoft Sans Serif,13'
$ActionForm.controls.AddRange(@($NoteText))

#Buttons
$CheckMovesBtn                  = New-Object system.Windows.Forms.Button
$CheckMovesBtn.BackColor         = "#32a852"
$CheckMovesBtn.text              = "Check Existing Moves"
$CheckMovesBtn.width             = 200
$CheckMovesBtn.height            = 30
$CheckMovesBtn.location          = New-Object System.Drawing.Point(20,110)
$CheckMovesBtn.Font              = 'Microsoft Sans Serif,10'
$CheckMovesBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($CheckMovesBtn)
$CheckMovesBtn.Add_Click({$Moves|Select displayname,Status|Out-GridView -PassThru})


#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Close"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(20,150)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000fff"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($cancelBtn)

$ActionForm.controls.AddRange(@($TitleOperationChoice,$Description,$Status))


# Display the form
$result = $ActionForm.ShowDialog()

}

function NewMoveForm  {
	Write-host "NewMoveForm" -fore Yellow
Add-Type -AssemblyName System.Windows.Forms
# Result form
$NewMoveForm                    = New-Object system.Windows.Forms.Form
$NewMoveForm.ClientSize         = '500,200'
$NewMoveForm.text               = "New Move Request"
$NewMoveForm.BackColor          = "#bababa"


if ($Valid -eq 1)  { [void]$ResultForm2.Close() }
########### Result Form cont.
#Account Name Heading

$NewPrimaryText                           = New-Object system.Windows.Forms.Label
$NewPrimaryText.text                      = 'You have chosen to move mailbox ' + $ID.name
$NewPrimaryText.AutoSize                  = $true
$NewPrimaryText.width                     = 25
$NewPrimaryText.height                    = 10
$NewPrimaryText.location                  = New-Object System.Drawing.Point(20,10)
$NewPrimaryText.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$NewMoveForm.controls.AddRange(@($NewPrimaryText))

# 

Write-host "NewMoveForm" $NewMove.text

#Execute Button
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#026075"
$ExecBtn.text              = "Start Move - Auto complete"
$ExecBtn.width             = 220
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,60)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$NewMoveForm.CancelButton   = $cancelBtn3
$NewMoveForm.Controls.Add($ExecBtn)

$ExecBtn.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain $Dom -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 2020-01-01 -erroraction silentlycontinue -warningaction silentlycontinue
MoveConfirmForm})

#Execute Button
$ExecBtn1                   = New-Object system.Windows.Forms.Button
$ExecBtn1.BackColor         = "#026075"
$ExecBtn1.text              = "Start Move - Manual complete"
$ExecBtn1.width             = 220
$ExecBtn1.height            = 30
$ExecBtn1.location          = New-Object System.Drawing.Point(20,100)
$ExecBtn1.Font              = 'Microsoft Sans Serif,10'
$ExecBtn1.ForeColor         = "#ffffff"
$NewMoveForm.CancelButton   = $cancelBtn3
$NewMoveForm.Controls.Add($ExecBtn1)

$ExecBtn1.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 9999-01-01 -erroraction silentlycontinue -warningaction silentlycontinue
MoveConfirmForm})

#Cancel Button
$cancelBtn3                       = New-Object system.Windows.Forms.Button
$cancelBtn3.BackColor             = "#ffffff"
$cancelBtn3.text                  = "Close"
$cancelBtn3.width                 = 90
$cancelBtn3.height                = 30
$cancelBtn3.location              = New-Object System.Drawing.Point(20,170)
$cancelBtn3.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn3.ForeColor             = "#000fff"
$cancelBtn3.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$NewMoveForm.CancelButton   = $cancelBtn3
$NewMoveForm.Controls.Add($cancelBtn3)

$cancelBtn3.Add_Click({ $NewMoveForm.Close() })

Write-host "NewMoveForm open Resultsform"

# Display the form
$result = $NewMoveForm.ShowDialog()
}

function NewBulkMoveForm  {
	Write-host "NewMoveForm" -fore Yellow
Add-Type -AssemblyName System.Windows.Forms
# Result form
$NewBulkMoveForm                    = New-Object system.Windows.Forms.Form
$NewBulkMoveForm.ClientSize         = '500,300'
$NewBulkMoveForm.text               = "New Move Request"
$NewBulkMoveForm.BackColor          = "#bababa"


if ($Valid -eq 1)  { [void]$ResultForm2.Close() }
########### Result Form cont.
#Account Name Heading
$NewPrimaryText                           = New-Object system.Windows.Forms.Label
$NewPrimaryText.text                      = 'You have chosen to do a bulk mailbox move.' + $ID.name
$NewPrimaryText.AutoSize                  = $true
$NewPrimaryText.width                     = 25
$NewPrimaryText.height                    = 10
$NewPrimaryText.location                  = New-Object System.Drawing.Point(20,10)
#$NewPrimaryText.ForeColor         = "#bababa"
$NewPrimaryText.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$NewBulkMoveForm.controls.AddRange(@($NewPrimaryText))

$NewSecondaryText                           = New-Object system.Windows.Forms.Label
$NewSecondaryText.text                      = 'File must be in CSV format with a DisplayName Column.' + $ID.name
$NewSecondaryText.AutoSize                  = $true
$NewSecondaryText.width                     = 25
$NewSecondaryText.height                    = 10
$NewSecondaryText.location                  = New-Object System.Drawing.Point(20,40)
$NewSecondaryText.ForeColor		            = "#dd00ff"
$NewSecondaryText.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$NewBulkMoveForm.controls.AddRange(@($NewSecondaryText))

# 

Write-host "NewBulkMoveForm" $NewBulkMove.text

#Import Button
$ImportBtn                   = New-Object system.Windows.Forms.Button
$ImportBtn.BackColor         = "#0000ff"
$ImportBtn.text              = "Import Mailbox List"
$ImportBtn.width             = 220
$ImportBtn.height            = 30
$ImportBtn.location          = New-Object System.Drawing.Point(20,80)
$ImportBtn.Font              = 'Microsoft Sans Serif,10'
$ImportBtn.ForeColor         = "#ffffff"
$NewBulkMoveForm.CancelButton   = $cancelBtn3
$NewBulkMoveForm.Controls.Add($ImportBtn)


$ImportBtn.Add_Click({BulkUserImport})

#Execute Button
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#026075"
$ExecBtn.text              = "Start Move - Auto complete"
$ExecBtn.width             = 220
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,140)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$NewBulkMoveForm.CancelButton   = $cancelBtn3
$NewBulkMoveForm.Controls.Add($ExecBtn)

$ExecBtn.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 2020-01-01})

#Execute Button
$ExecBtn1                   = New-Object system.Windows.Forms.Button
$ExecBtn1.BackColor         = "#026075"
$ExecBtn1.text              = "Start Move - Manual complete"
$ExecBtn1.width             = 220
$ExecBtn1.height            = 30
$ExecBtn1.location          = New-Object System.Drawing.Point(20,180)
$ExecBtn1.Font              = 'Microsoft Sans Serif,10'
$ExecBtn1.ForeColor         = "#ffffff"
$NewBulkMoveForm.CancelButton   = $cancelBtn3
$NewBulkMoveForm.Controls.Add($ExecBtn1)

$ExecBtn1.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 9999-01-01})

#Cancel Button
$cancelBtn3                       = New-Object system.Windows.Forms.Button
$cancelBtn3.BackColor             = "#ffffff"
$cancelBtn3.text                  = "Close"
$cancelBtn3.width                 = 90
$cancelBtn3.height                = 30
$cancelBtn3.location              = New-Object System.Drawing.Point(20,250)
$cancelBtn3.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn3.ForeColor             = "#000fff"
$cancelBtn3.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$NewBulkMoveForm.CancelButton   = $cancelBtn3
$NewBulkMoveForm.Controls.Add($cancelBtn3)

$cancelBtn3.Add_Click({ $NewBulkMoveForm.Close() })

Write-host "NewBulkMoveForm open Resultsform"

# Display the form
$result = $NewBulkMoveForm.ShowDialog()
}

function FinaliseMoveForm {
Write-Host "FinaliseMoveForm" -fore yellow
Add-Type -AssemblyName System.Windows.Forms
# Add Alias form
$CompleteMoveForm                    = New-Object system.Windows.Forms.Form
$CompleteMoveForm.ClientSize         = '600,200'
$CompleteMoveForm.text               = "Complete a Mailbox Move"
$CompleteMoveForm.BackColor          = "#bababa"


<# if ($Valid -eq 1)  { [void]$ResultForm2.Close() } #>

########### Result Form cont.
#Account Name Heading
$FinaliseMoveText                           = New-Object system.Windows.Forms.Label
$FinaliseMoveText.text                      = 'You are completing the mailbox move for ' + $ID.name
$FinaliseMoveText.AutoSize                  = $true
$FinaliseMoveText.width                     = 30
$FinaliseMoveText.height                    = 10
$FinaliseMoveText.location                  = New-Object System.Drawing.Point(20,10)
$FinaliseMoveText.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$CompleteMoveForm.controls.AddRange(@($FinaliseMoveText))

$FinaliseMoveText2                           = New-Object system.Windows.Forms.Label
$FinaliseMoveText2.text                      = "(This may take a few seconds)"
$FinaliseMoveText2.AutoSize                  = $true
$FinaliseMoveText2.width                     = 30
$FinaliseMoveText2.height                    = 10
$FinaliseMoveText2.location                  = New-Object System.Drawing.Point(20,40)
$FinaliseMoveText2.ForeColor                 = "#ff0000"
$FinaliseMoveText2.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$CompleteMoveForm.controls.AddRange(@($FinaliseMoveText2))



#Result Buttons
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#026075"
$ExecBtn.text              = "Complete Move"
$ExecBtn.width             = 120
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,90)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$CompleteMoveForm.CancelButton   = $cancelBtn4
$CompleteMoveForm.Controls.Add($ExecBtn)
$ExecBtn.Add_Click({
	Write-host $Id.name " ID.Name"
	$MC = $Null
	$MC =  $moves |where {$_.name -eq $ID.name}
	Write-Host "Value: $MC"
	If (!$MC){NoMoveForm
	Write-Host "No move associated with user"}
	If ($MC) {set-moverequest $ID.name -completeafter 2020-01-01
	MoveConfirmForm}
	})

#Cancel Button
$cancelBtn4                       = New-Object system.Windows.Forms.Button
$cancelBtn4.BackColor             = "#ffffff"
$cancelBtn4.text                  = "Close"
$cancelBtn4.width                 = 120
$cancelBtn4.height                = 30
$cancelBtn4.location              = New-Object System.Drawing.Point(20,130)
$cancelBtn4.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn4.ForeColor             = "#000fff"
$cancelBtn4.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$CompleteMoveForm.CancelButton   = $cancelBtn4
$CompleteMoveForm.Controls.Add($cancelBtn4)
$cancelBtn4.Add_Click({ $CompleteMoveForm.close() })

# Display the form
$result = $CompleteMoveForm.ShowDialog()
}

Function ActionForm {
Write-host "ActionForm" -fore Yellow
Add-Type -AssemblyName System.Windows.Forms
# Create a new form
$ActionForm                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$ActionForm.ClientSize         = '500,300'
$ActionForm.text               = "Mailbox Move Management"
$ActionForm.BackColor          = "#ffffff"

# Create a Title for our form. We will use a label for it.
$TitleOperationChoice                           = New-Object system.Windows.Forms.Label
$TitleOperationChoice.text                      = "Mailbox Move Management"
$TitleOperationChoice.AutoSize                  = $true
$TitleOperationChoice.width                     = 25
$TitleOperationChoice.height                    = 10
$TitleOperationChoice.location                  = New-Object System.Drawing.Point(20,20)
$TitleOperationChoice.Font                      = 'Microsoft Sans Serif,13'

# Other elemtents

#Dropdown Text Box
#TextBoxLable
$OperationChoiceLabel                = New-Object system.Windows.Forms.Label
$OperationChoiceLabel.text           = "Select Option:"
$OperationChoiceLabel.AutoSize       = $true
$OperationChoiceLabel.width          = 25
$OperationChoiceLabel.height         = 20
$OperationChoiceLabel.location       = New-Object System.Drawing.Point(20,130)
$OperationChoiceLabel.Font           = 'Microsoft Sans Serif,13,style=Bold'
$OperationChoiceLabel.Visible        = $True
$ActionForm.Controls.Add($OperationChoiceLabel)

$OperationChoice                     = New-Object system.Windows.Forms.ComboBox
$OperationChoice.text                = "Choose"
$OperationChoice.width               = 200
$OperationChoice.autosize            = $true
$OperationChoice.Visible             = $true        
# Add the items in the dropdown list
@("Choose an option",$NewMoveForm,$CompleteMoveForm,$ViewMoveForm,$RemoveMoveForm) | ForEach-Object {[void] $OperationChoice.Items.Add($_)}
# Select the default value
$OperationChoice.SelectedIndex       = 0
$OperationChoice.location            = New-Object System.Drawing.Point(150,130)
$OperationChoice.ForeColor           = "#016113"
$OperationChoice.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {    
if ($OperationChoice.text -eq $NewMoveForm) {NewMoveForm}
if ($OperationChoice.text -eq $CompleteMoveForm) {FinaliseMoveForm}
if ($OperationChoice.text -eq $ViewMoveForm) {MoveConfirmForm}
if ($OperationChoice.text -eq $RemoveMoveForm) {RemoveMoveForm}
#if ($OperationChoice.text -eq $sub5) {Sub5}
    }
})
$ActionForm.Controls.Add($OperationChoice)

$Name = $ID.name
If (!$name) {$Name = "No User Selected"}
#TextBoxLable
$SearchNameLabel                = New-Object system.Windows.Forms.Label
$SearchNameLabel.text           = "You are managing user:"
$SearchNameLabel.AutoSize       = $true
$SearchNameLabel.width          = 25
$SearchNameLabel.height         = 20
$SearchNameLabel.location       = New-Object System.Drawing.Point(20,80)
$SearchNameLabel.Font           = 'Microsoft Sans Serif,14,style=Bold'
$SearchNameLabel.Visible        = $True
$ActionForm.Controls.Add($SearchNameLabel)

#TextBoxLable
$NameLable                = New-Object system.Windows.Forms.Label
$NameLable.text           = $ID
$NameLable.AutoSize       = $true
$NameLable.width          = 25
$NameLable.height         = 20
$NameLable.ForeColor 	  = "#0000ff"
$NameLable.location       = New-Object System.Drawing.Point(240,80)
$NameLable.Font           = 'Microsoft Sans Serif,14,style=Bold'
$NameLable.Visible        = $True
$ActionForm.Controls.Add($NameLable)


#Buttons
$ExecuteBtn                   = New-Object system.Windows.Forms.Button
$ExecuteBtn.BackColor         = "#026075"
$ExecuteBtn.text              = "Execute"
$ExecuteBtn.width             = 90
$ExecuteBtn.height            = 30
$ExecuteBtn.location          = New-Object System.Drawing.Point(150,250)
$ExecuteBtn.Font              = 'Microsoft Sans Serif,10,style=bold'
$ExecuteBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($ExecuteBtn)


$ExecuteBtn.Add_Click({ 
if ($OperationChoice.text -eq $NewMoveForm) {NewMoveForm}
if ($OperationChoice.text -eq $CompleteMoveForm) {FinaliseMoveForm}
if ($OperationChoice.text -eq $ViewMoveForm) {MoveConfirmForm}
if ($OperationChoice.text -eq $RemoveMoveForm) {CheckEmailForm}
 })



#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Close"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(260,250)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000fff"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($cancelBtn)

$ActionForm.controls.AddRange(@($TitleOperationChoice,$Description,$Status))


# Display the form
$result = $ActionForm.ShowDialog()
}

Function BulkActionForm { #Incomplete
Write-host "BulkActionForm" -fore Yellow
Add-Type -AssemblyName System.Windows.Forms
# Create a new form
$BulkActionForm                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$BulkActionForm.ClientSize         = '500,300'
$BulkActionForm.text               = "Mailbox Move Management"
$BulkActionForm.BackColor          = "#ffffff"

# Create a Title for our form. We will use a label for it.
$TitleOperationChoice                           = New-Object system.Windows.Forms.Label
$TitleOperationChoice.text                      = "Mailbox Move Management"
$TitleOperationChoice.AutoSize                  = $true
$TitleOperationChoice.width                     = 25
$TitleOperationChoice.height                    = 10
$TitleOperationChoice.location                  = New-Object System.Drawing.Point(20,20)
$TitleOperationChoice.Font                      = 'Microsoft Sans Serif,13'

# Other elemtents

#Dropdown Text Box
#TextBoxLable
$OperationChoiceLabel                = New-Object system.Windows.Forms.Label
$OperationChoiceLabel.text           = "Select Option:"
$OperationChoiceLabel.AutoSize       = $true
$OperationChoiceLabel.width          = 25
$OperationChoiceLabel.height         = 20
$OperationChoiceLabel.location       = New-Object System.Drawing.Point(20,130)
$OperationChoiceLabel.Font           = 'Microsoft Sans Serif,13,style=Bold'
$OperationChoiceLabel.Visible        = $True
$BulkActionForm.Controls.Add($OperationChoiceLabel)

$OperationChoice                     = New-Object system.Windows.Forms.ComboBox
$OperationChoice.text                = "Choose"
$OperationChoice.width               = 200
$OperationChoice.autosize            = $true
$OperationChoice.Visible             = $true        
# Add the items in the dropdown list
@("Choose an option",$NewBulkMoveForm,$CompleteBulkMoveForm,$ViewBulkMoveForm,$RemoveBulkMoveForm) | ForEach-Object {[void] $OperationChoice.Items.Add($_)}
# Select the default value
$OperationChoice.SelectedIndex       = 0
$OperationChoice.location            = New-Object System.Drawing.Point(150,130)
$OperationChoice.ForeColor           = "#016113"
$OperationChoice.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {    
if ($OperationChoice.text -eq $NewBulkMoveForm) {NewBulkMoveForm}
if ($OperationChoice.text -eq $CompleteBulkMoveForm) {FinaliseBulkMoveForm}
if ($OperationChoice.text -eq $ViewBulkMoveForm) {BulkMoveConfirmForm}
if ($OperationChoice.text -eq $RemoveBulkMoveForm) {RemoveBulkMoveForm}
#if ($OperationChoice.text -eq $sub5) {Sub5}
    }
})
$BulkActionForm.Controls.Add($OperationChoice)

$Name = $ID.name
If (!$name) {$Name = "No User Selected"}
#TextBoxLable
$BulkNameLable                = New-Object system.Windows.Forms.Label
$BulkNameLable.text           = "Bulk Move Operations"
$BulkNameLable.AutoSize       = $true
$BulkNameLable.width          = 25
$BulkNameLable.height         = 20
$BulkNameLable.location       = New-Object System.Drawing.Point(20,80)
$BulkNameLable.Font           = 'Microsoft Sans Serif,14,style=Bold'
$BulkNameLable.Visible        = $True
$BulkActionForm.Controls.Add($BulkNameLable)


#Buttons
$ExecuteBtn                   = New-Object system.Windows.Forms.Button
$ExecuteBtn.BackColor         = "#026075"
$ExecuteBtn.text              = "Execute"
$ExecuteBtn.width             = 90
$ExecuteBtn.height            = 30
$ExecuteBtn.location          = New-Object System.Drawing.Point(150,250)
$ExecuteBtn.Font              = 'Microsoft Sans Serif,10,style=bold'
$ExecuteBtn.ForeColor         = "#ffffff"
$BulkActionForm.CancelButton   = $cancelBtn
$BulkActionForm.Controls.Add($ExecuteBtn)


$ExecuteBtn.Add_Click({ 
if ($OperationChoice.text -eq $NewBulkMoveForm) {NewMoveForm}
if ($OperationChoice.text -eq $CompleteBulkMoveForm) {FinaliseMoveForm}
if ($OperationChoice.text -eq $ViewMoveForm) {MoveConfirmForm}
if ($OperationChoice.text -eq $RemoveAllCompleteForm) {RemoveAllCompleteForm}
 })



#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Close"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(260,250)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000fff"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$BulkActionForm.CancelButton   = $cancelBtn
$BulkActionForm.Controls.Add($cancelBtn)

$BulkActionForm.controls.AddRange(@($TitleOperationChoice,$Description,$Status))


# Display the form
$result = $BulkActionForm.ShowDialog()


}

Function InvalidUserForm {
	Write-host "InvalidUserForm" -fore Yellow
# Ivalid User Form
$InvalidUserForm                    = New-Object system.Windows.Forms.Form
$InvalidUserForm.ClientSize         = '400,100'
$InvalidUserForm.text               = "Invalid User"
$InvalidUserForm.BackColor          = "#bababa"

#Account Name Heading
$InvalidUserText                           = New-Object system.Windows.Forms.Label
$InvalidUserText.text                      = 'No User has been selected'
$InvalidUserText.AutoSize                  = $true
$InvalidUserText.width                     = 25
$InvalidUserText.height                    = 10
$InvalidUserText.ForeColor                 = "#ff0000"
$InvalidUserText.location                  = New-Object System.Drawing.Point(20,10)
$InvalidUserText.Font                      = 'Microsoft Sans Serif,13'
$InvalidUserForm.controls.AddRange(@($InvalidUserText))

$InvalidUserForm.ShowDialog()

}
 
Function NoMoveForm {
write-host "NoMoveForm" -fore Yellow
$NoMoveForm                    = New-Object system.Windows.Forms.Form
$NoMoveForm.ClientSize         = '650,100'
$NoMoveForm.text               = "Invalid User"
$NoMoveForm.BackColor          = "#bababa"

#Account Name Heading
$NoMoveText                           = New-Object system.Windows.Forms.Label
$NoMoveText.text                      = "User " + $ID.Name  + "is has no Move Request associated with it."
$NoMoveText.AutoSize                  = $true
$NoMoveText.width                     = 25
$NoMoveText.height                    = 10
$NoMoveText.ForeColor                 = "#ff0000"
$NoMoveText.location                  = New-Object System.Drawing.Point(20,10)
$NoMoveText.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$NoMoveForm.controls.AddRange(@($NoMoveText))

#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Close"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(20,50)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000fff"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$NoMoveForm.CancelButton   = $cancelBtn
$NoMoveForm.Controls.Add($cancelBtn)

$NoMoveForm.ShowDialog()

}

function FindUserForm { 
  Write-host "FindUserForm" -fore Yellow
  #Username to be used in code
  #$ID=get-adobject -filter 'cn -like $searchstring' |Out-GridView -PassThru
  $ID=$allusers   |select name|Out-GridView -PassThru
  $CN = $ID.Name
Write-host "ID is $ID"
  #Get user data
Write-host "User selected: " $CN
If ($CN -eq $true){actionform}
if (!$CN) {write-host "Username invalid:" $SearchName.Text -ForegroundColor Red
$Valid = 0
InvalidUserForm 
}
if ($CN) {write-host "Username Valid:" $SearchName.Text -ForegroundColor Cyan
$Valid = 0
ActionForm 
}

}

Function MoveConfirmForm{
	# Ivalid User Form
	Write-host "MoveConfirmForm"
$MoveConfirmForm                    = New-Object system.Windows.Forms.Form
$MoveConfirmForm.ClientSize         = '400,100'
$MoveConfirmForm.text               = "Move Status"
$MoveConfirmForm.BackColor          = "#bababa"

#Account Name Heading
$MoveConfirmText                           = New-Object system.Windows.Forms.Label
$MoveConfirmText.text                      = "Mailbox: " + $ID.name
$MoveConfirmText.AutoSize                  = $true
$MoveConfirmText.width                     = 25
$MoveConfirmText.height                    = 10
#$MoveConfirmText.ForeColor                 = "#ff0000"
$MoveConfirmText.location                  = New-Object System.Drawing.Point(20,10)
$MoveConfirmText.Font                      = 'Microsoft Sans Serif,13'
$MoveConfirmForm.controls.AddRange(@($MoveConfirmText))

Write-host $ID

$MRS = Get-moverequest $ID.name |get-moverequestStatistics
$MRT = $MRS |select DisplayName,StatusDetail,PercentComplete

$MoveConfirmDetail                           = New-Object system.Windows.Forms.Label
$MoveConfirmDetail.text                      = "Status: " + $MRT.StatusDetail.Value + ", "  + $MRT.PercentComplete + "% Complete"
$MoveConfirmDetail.AutoSize                  = $true
$MoveConfirmDetail.width                     = 300
$MoveConfirmDetail.height                    = 10
#$MoveConfirmDetail.ForeColor                 = "#ff0000"
$MoveConfirmDetail.location                  = New-Object System.Drawing.Point(20,40)
$MoveConfirmDetail.Font                      = 'Microsoft Sans Serif,13'
$MoveConfirmForm.controls.AddRange(@($MoveConfirmDetail))



$MoveConfirmForm.ShowDialog()
}

Function RemoveMoveForm {
Write-Host "FinaliseMoveForm" -fore yellow
Add-Type -AssemblyName System.Windows.Forms
# Add Alias form
$RemoveMoveForm                    = New-Object system.Windows.Forms.Form
$RemoveMoveForm.ClientSize         = '600,200'
$RemoveMoveForm.text               = "Complete a Mailbox Move"
$RemoveMoveForm.BackColor          = "#bababa"


<# if ($Valid -eq 1)  { [void]$ResultForm2.Close() } #>

########### Result Form cont.
#Account Name Heading
$RemoveMoveText                           = New-Object system.Windows.Forms.Label
$RemoveMoveText.text                      = 'You are removing the Move Request for ' + $ID.name
$RemoveMoveText.AutoSize                  = $true
$RemoveMoveText.width                     = 30
$RemoveMoveText.height                    = 10
$RemoveMoveText.location                  = New-Object System.Drawing.Point(20,10)
$RemoveMoveText.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$RemoveMoveForm.controls.AddRange(@($RemoveMoveText))

$RemoveMoveText2                           = New-Object system.Windows.Forms.Label
$RemoveMoveText2.text                      = "(This may take a few seconds)"
$RemoveMoveText2.AutoSize                  = $true
$RemoveMoveText2.width                     = 30
$RemoveMoveText2.height                    = 10
$RemoveMoveText2.location                  = New-Object System.Drawing.Point(20,40)
$RemoveMoveText2.ForeColor                 = "#ff0000"
$RemoveMoveText2.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$RemoveMoveForm.controls.AddRange(@($RemoveMoveText2))



#Result Buttons
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#026075"
$ExecBtn.text              = "Remove Move Request"
$ExecBtn.width             = 120
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,90)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$RemoveMoveForm.CancelButton   = $cancelBtn4
$RemoveMoveForm.Controls.Add($ExecBtn)
$ExecBtn.Add_Click({
	Write-host $Id.name " ID.Name"
	$MC = $Null
	$MC =  $moves |where {$_.name -eq $ID.name}
	Write-Host "Value: $MC"
	If (!$MC){NoMoveForm
	Write-Host "No move associated with user"}
	If ($MC) {Remove-moverequest $ID.name -confirm $False 
	MoveConfirmForm}
	})

#Cancel Button
$cancelBtn4                       = New-Object system.Windows.Forms.Button
$cancelBtn4.BackColor             = "#ffffff"
$cancelBtn4.text                  = "Close"
$cancelBtn4.width                 = 120
$cancelBtn4.height                = 30
$cancelBtn4.location              = New-Object System.Drawing.Point(20,130)
$cancelBtn4.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn4.ForeColor             = "#000fff"
$cancelBtn4.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$RemoveMoveForm.CancelButton   = $cancelBtn4
$RemoveMoveForm.Controls.Add($cancelBtn4)
$cancelBtn4.Add_Click({ $RemoveMoveForm.close() })

# Display the form
$result = $RemoveMoveForm.ShowDialog()
	}

Function RemoveAllCompleteForm {
Write-Host "RemoveAllCompleteForm" -fore yellow
Add-Type -AssemblyName System.Windows.Forms
# Add Alias form
$RemoveAllCompleteForm                    = New-Object system.Windows.Forms.Form
$RemoveAllCompleteForm.ClientSize         = '600,200'
$RemoveAllCompleteForm.text               = "Complete a Mailbox Move"
$RemoveAllCompleteForm.BackColor          = "#bababa"


<# if ($Valid -eq 1)  { [void]$ResultForm2.Close() } #>

########### Result Form cont.
#Account Name Heading
$RemoveAllMoveText                           = New-Object system.Windows.Forms.Label
$RemoveAllMoveText.text                      = 'You are removing the Move Request for ' + $ID.name
$RemoveAllMoveText.AutoSize                  = $true
$RemoveAllMoveText.width                     = 30
$RemoveAllMoveText.height                    = 10
$RemoveAllMoveText.location                  = New-Object System.Drawing.Point(20,10)
$RemoveAllMoveText.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$RemoveAllCompleteForm.controls.AddRange(@($RemoveAllMoveText))

$RemoveAllMoveText2                           = New-Object system.Windows.Forms.Label
$RemoveAllMoveText2.text                      = "(This may take a few seconds)"
$RemoveAllMoveText2.AutoSize                  = $true
$RemoveAllMoveText2.width                     = 30
$RemoveAllMoveText2.height                    = 10
$RemoveAllMoveText2.location                  = New-Object System.Drawing.Point(20,40)
$RemoveAllMoveText2.ForeColor                 = "#ff0000"
$RemoveAllMoveText2.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$RemoveAllCompleteForm.controls.AddRange(@($RemoveAllMoveText2))



#Result Buttons
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#026075"
$ExecBtn.text              = "Remove Move Request"
$ExecBtn.width             = 120
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,90)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$RemoveAllCompleteForm.CancelButton   = $cancelBtn4
$RemoveAllCompleteForm.Controls.Add($ExecBtn)
$ExecBtn.Add_Click({
	Write-host $Id.name " ID.Name"
	$MC = $Null
	$MC =  $moves |where {$_.name -eq $ID.name}
	Write-Host "Value: $MC"
	If (!$MC){NoMoveForm
	Write-Host "No move associated with user"}
	If ($MC) {$Moves |where {$_.Status -eq "Completed"} |Remove-moverequest -confirm $False -
	#MoveConfirmForm2
	#$moves = Get-moverequest
	}
	})

#Cancel Button
$cancelBtn4                       = New-Object system.Windows.Forms.Button
$cancelBtn4.BackColor             = "#ffffff"
$cancelBtn4.text                  = "Close"
$cancelBtn4.width                 = 120
$cancelBtn4.height                = 30
$cancelBtn4.location              = New-Object System.Drawing.Point(20,130)
$cancelBtn4.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn4.ForeColor             = "#000fff"
$cancelBtn4.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$RemoveAllCompleteForm.CancelButton   = $cancelBtn4
$RemoveAllCompleteForm.Controls.Add($cancelBtn4)
$cancelBtn4.Add_Click({ $RemoveAllCompleteForm.close() })

# Display the form
$result = $RemoveAllCompleteForm.ShowDialog()
	}

Function BulkUserImport { #Incomplete
	
	
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('MyDocuments') }
$null = $FileBrowser.ShowDialog()	
$List = Import-csv $FileBrowser.FileName	

}

Function ClockCode {
	do {
      $clock = Get-Date -format 'HH:mm:ss'
      $clock
      Start-Sleep -Seconds 1
    } while ($true)}

############Form Functions End

# Init PowerShell GUI
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$script =  $scriptPath + "\" + $MyInvocation.MyCommand.name 

#Initiate GUI
Write-Host "Office 365 Move Management GUI Initiated" -for green
chooseForm
