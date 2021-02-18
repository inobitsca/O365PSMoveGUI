# GUI to Manage Exchange Online PowerShell Based Move Requests .
#
#
#MUST be run from Exange online PowerShell
##Requires ACTIVEDIRECTORY PowerShell module. 
#You will need User Admin rights.
#Created by Cedric Abrahams - cedric@inobits.com
#
#Version 1.3 2021-01-05


#Connect-EXOPSSession
Write-Host "Getting Mailbox and Move details.
This may take some time." -fore green
#$allUsers = Get-MailUser
#$moves = Get-moverequest
#$credential = get-credential

#Variables
$result = ""
$res = $false
$IT = 0
$user = 'None'
$NP =''
$NewEmail =''
$check = $null

#Menu Options
$NewMoveForm = 'Create New Move Requests'
$CompleteMoveForm = 'Complete Move Request'
$ViewMoveForm = 'View Existing Move Requests'
$RemoveMoveForm = 'Remove Move Requests'

$NewBulkMoveForm = "New Bulk Move Request"
$CompleteBulkMoveForm = "Finalise Bulk Move Requests"
$ViewBulkMoveForm = "View Bulk Moves Status"
$RemoveBulkMoveForm = "Remove Bulk Move Requests"

############Form Functions Start

Function ChooseForm {
	Add-Type -AssemblyName System.Windows.Forms
	Write-host "ChooseForm" -Fore Yellow
# Create a new form
$ActionForm                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$ActionForm.ClientSize         = '300,240'
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


#Buttons
$SingleUserBtn                   = New-Object system.Windows.Forms.Button
$SingleUserBtn.BackColor         = "#a4ba67"
$SingleUserBtn.text              = "Individual User Operations"
$SingleUserBtn.width             = 200
$SingleUserBtn.height            = 30
$SingleUserBtn.location          = New-Object System.Drawing.Point(20,80)
$SingleUserBtn.Font              = 'Microsoft Sans Serif,10'
$SingleUserBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($SingleUserBtn)
$SingleUserBtn.Add_Click({FindUserForm})

#Buttons
$BulkUserBtn                   = New-Object system.Windows.Forms.Button
$BulkUserBtn.BackColor         = "#a4ba67"
$BulkUserBtn.text              = "Bulk User Operations"
$BulkUserBtn.width             = 200
$BulkUserBtn.height            = 30
$BulkUserBtn.location          = New-Object System.Drawing.Point(20,110)
$BulkUserBtn.Font              = 'Microsoft Sans Serif,10'
$BulkUserBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($BulkUserBtn)
$BulkUserBtn.Add_Click({BulkActionForm})

#Buttons
$CheckMovesBtn                  = New-Object system.Windows.Forms.Button
$CheckMovesBtn.BackColor         = "#32a852"
$CheckMovesBtn.text              = "Check Existing Moves"
$CheckMovesBtn.width             = 200
$CheckMovesBtn.height            = 30
$CheckMovesBtn.location          = New-Object System.Drawing.Point(20,150)
$CheckMovesBtn.Font              = 'Microsoft Sans Serif,10'
$CheckMovesBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($CheckMovesBtn)
$CheckMovesBtn.Add_Click({$Moves|Select displayname,Status|Out-GridView -PassThru})

#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Cancel"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(20,190)
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
$ExecBtn.BackColor         = "#a4ba67"
$ExecBtn.text              = "Start Move - Auto complete"
$ExecBtn.width             = 220
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,60)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$NewMoveForm.CancelButton   = $cancelBtn3
$NewMoveForm.Controls.Add($ExecBtn)

$ExecBtn.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 2020-01-01})

#Execute Button
$ExecBtn1                   = New-Object system.Windows.Forms.Button
$ExecBtn1.BackColor         = "#a4ba67"
$ExecBtn1.text              = "Start Move - Manual complete"
$ExecBtn1.width             = 220
$ExecBtn1.height            = 30
$ExecBtn1.location          = New-Object System.Drawing.Point(20,100)
$ExecBtn1.Font              = 'Microsoft Sans Serif,10'
$ExecBtn1.ForeColor         = "#ffffff"
$NewMoveForm.CancelButton   = $cancelBtn3
$NewMoveForm.Controls.Add($ExecBtn1)

$ExecBtn1.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 9999-01-01})

#Cancel Button
$cancelBtn3                       = New-Object system.Windows.Forms.Button
$cancelBtn3.BackColor             = "#ffffff"
$cancelBtn3.text                  = "Cancel"
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
$NewBulkMoveForm.ClientSize         = '500,200'
$NewBulkMoveForm.text               = "New Move Request"
$NewBulkMoveForm.BackColor          = "#bababa"


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
$NewBulkMoveForm.controls.AddRange(@($NewPrimaryText))

# 

Write-host "NewBulkMoveForm" $NewBulkMove.text

#Execute Button
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#a4ba67"
$ExecBtn.text              = "Start Move - Auto complete"
$ExecBtn.width             = 220
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,60)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$NewBulkMoveForm.CancelButton   = $cancelBtn3
$NewBulkMoveForm.Controls.Add($ExecBtn)

$ExecBtn.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 2020-01-01})

#Execute Button
$ExecBtn1                   = New-Object system.Windows.Forms.Button
$ExecBtn1.BackColor         = "#a4ba67"
$ExecBtn1.text              = "Start Move - Manual complete"
$ExecBtn1.width             = 220
$ExecBtn1.height            = 30
$ExecBtn1.location          = New-Object System.Drawing.Point(20,100)
$ExecBtn1.Font              = 'Microsoft Sans Serif,10'
$ExecBtn1.ForeColor         = "#ffffff"
$NewBulkMoveForm.CancelButton   = $cancelBtn3
$NewBulkMoveForm.Controls.Add($ExecBtn1)

$ExecBtn1.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 9999-01-01})

#Cancel Button
$cancelBtn3                       = New-Object system.Windows.Forms.Button
$cancelBtn3.BackColor             = "#ffffff"
$cancelBtn3.text                  = "Cancel"
$cancelBtn3.width                 = 90
$cancelBtn3.height                = 30
$cancelBtn3.location              = New-Object System.Drawing.Point(20,170)
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
$ExecBtn.BackColor         = "#a4ba67"
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
	If ($MC) {set-moverequest $ID.name -completeafter 2020-01-01 -whatif
	MoveConfirmForm1}
	})

#Cancel Button
$cancelBtn4                       = New-Object system.Windows.Forms.Button
$cancelBtn4.BackColor             = "#ffffff"
$cancelBtn4.text                  = "Cancel"
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

<# Function ResultsForm {
Write-host "Resultsform" -fore Yellow
if ($user -ne "None") {Write-host "Check Address" $NewEmail -ForegroundColor Green
$res = $true

$user = get-adobject $CN -Properties mail,proxyaddresses


# Process form
$ResultsForm                    = New-Object system.Windows.Forms.Form
$ResultsForm.ClientSize         = '500,600'
$ResultsForm.text               = "Email Process Management"
$ResultsForm.BackColor          = "#abdbff"


#Create a Title for our form. We will use a label for it.
$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Email Addresses for ' + $user.name + " (" +$SearchName.Text +  ')'
$EmailTitle.AutoSize                  = $true
$EmailTitle.location                  = New-Object System.Drawing.Point(20,25)
$EmailTitle.Font                      = 'Microsoft Sans Serif,13,style=bold'
$EmailTitle.ForeColor                 = "#151051"
$ResultsForm.controls.AddRange(@($EmailTitle))


$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Primary Email Address:'
$EmailTitle.AutoSize                  = $true
$EmailTitle.width                     = 25
$EmailTitle.height                    = 10
$EmailTitle.location                  = New-Object System.Drawing.Point(20,75)
$EmailTitle.Font                      = 'Microsoft Sans Serif,13'
$ResultsForm.controls.AddRange(@($EmailTitle))


# Show Email Address
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $user.mail
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
$EmailChoice.location                  = New-Object System.Drawing.Point(20,100)
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$ResultsForm.controls.AddRange(@($EmailChoice))

# Show Proxy Addresses Heading.
$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Alias Addresses:'
$EmailTitle.AutoSize                  = $true
$EmailTitle.width                     = 25
$EmailTitle.height                    = 10
$EmailTitle.location                  = New-Object System.Drawing.Point(20,150)
$EmailTitle.Font                      = 'Microsoft Sans Serif,13'
$ResultsForm.controls.AddRange(@($EmailTitle))

$VL = 175
Foreach ($P in $user.proxyaddresses) {
If ($P -clike "*smtp*") {$Pr = $P 
$PR =$PR -replace 'smtp:',''
# Show Proxy Addresses
$EmailChoice                          = New-Object system.Windows.Forms.Label

$EmailChoice.text                      = $PR
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
$EmailChoice.location                  = New-Object System.Drawing.Point(20,$VL)
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$ResultsForm.controls.AddRange(@($EmailChoice))
$VL = $VL +25}}

#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Close"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(220,560)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000fff"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$ResultsForm.CancelButton   = $cancelBtn
$ResultsForm.Controls.Add($cancelBtn)





# Display the form
[void]$ResultsForm.ShowDialog()
}}
 #>

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
$ExecuteBtn.BackColor         = "#a4ba67"
$ExecuteBtn.text              = "Execute"
$ExecuteBtn.width             = 90
$ExecuteBtn.height            = 30
$ExecuteBtn.location          = New-Object System.Drawing.Point(150,250)
$ExecuteBtn.Font              = 'Microsoft Sans Serif,10'
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
$cancelBtn.text                  = "Cancel"
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

Function BulkActionForm {
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
if ($OperationChoice.text -eq $NewBulkMoveForm) {NewMoveForm}
if ($OperationChoice.text -eq $CompleteBulkMoveForm) {FinaliseMoveForm}
if ($OperationChoice.text -eq $ViewBulkMoveForm) {MoveConfirmForm}
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
$ExecuteBtn.BackColor         = "#a4ba67"
$ExecuteBtn.text              = "Execute"
$ExecuteBtn.width             = 90
$ExecuteBtn.height            = 30
$ExecuteBtn.location          = New-Object System.Drawing.Point(150,250)
$ExecuteBtn.Font              = 'Microsoft Sans Serif,10'
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
$cancelBtn.text                  = "Cancel"
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

<# Function InvalidUserForm {
	Write-host "nvalidUserForm" -fore Yellow
# Ivalid User Form
$InvalidUserForm                    = New-Object system.Windows.Forms.Form
$InvalidUserForm.ClientSize         = '400,100'
$InvalidUserForm.text               = "Invalid User"
$InvalidUserForm.BackColor          = "#bababa"

#Account Name Heading
$InvalidUserText                           = New-Object system.Windows.Forms.Label
$InvalidUserText.text                      = 'User $ID.Name is has been migrated'
$InvalidUserText.AutoSize                  = $true
$InvalidUserText.width                     = 25
$InvalidUserText.height                    = 10
$InvalidUserText.ForeColor                 = "#ff0000"
$InvalidUserText.location                  = New-Object System.Drawing.Point(20,10)
$InvalidUserText.Font                      = 'Microsoft Sans Serif,13'
$InvalidUserForm.controls.AddRange(@($InvalidUserText))

Write-host" $ID.Name ID.Name $CN CN"
$InvalidUserForm.ShowDialog()

}
#>
 
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
$cancelBtn.text                  = "Cancel"
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

<#Function StartForm {
Write-host "StartForm" -fore yellow
Add-Type -AssemblyName System.Windows.Forms
# Create a new form
$StartForm                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$StartForm.ClientSize         = '300,200'
$StartForm.text               = "Mailbox Move Management"
$StartForm.BackColor          = "#cceeff"

# Create a Title for our form. We will use a label for it.
$TitleOperationChoice                           = New-Object system.Windows.Forms.Label
$TitleOperationChoice.text                      = "Mailbox Move Management"
$TitleOperationChoice.AutoSize                  = $true
$TitleOperationChoice.width                     = 25
$TitleOperationChoice.height                    = 10
$TitleOperationChoice.location                  = New-Object System.Drawing.Point(20,20)
$TitleOperationChoice.Font                      = 'Microsoft Sans Serif,13'

# Other elemtents
$Description                     = New-Object system.Windows.Forms.Label
$Description.text                = "Select a user."
$Description.AutoSize            = $false
$Description.width               = 450
$Description.height              = 35
$Description.location            = New-Object System.Drawing.Point(20,50)
$Description.Font                = 'Microsoft Sans Serif,10'
<# $Status                   = New-Object system.Windows.Forms.Label
$Status.text              = "Please enter the username below"
$Status.AutoSize          = $true
$Status.location          = New-Object System.Drawing.Point(20,170)
$Status.Font              = 'Microsoft Sans Serif,10' 


TextBoxLable
$SearchNameLabel                = New-Object system.Windows.Forms.Label
$SearchNameLabel.text           = "Search for name: "
$SearchNameLabel.AutoSize       = $true
$SearchNameLabel.width          = 25
$SearchNameLabel.height         = 20
$SearchNameLabel.location       = New-Object System.Drawing.Point(20,200)
$SearchNameLabel.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SearchNameLabel.Visible        = $True
$StartForm.Controls.Add($SearchNameLabel) 

TextBox
$SearchName                     = New-Object system.Windows.Forms.TextBox
$SearchName.multiline           = $false
$SearchName.width               = 314
$SearchName.height              = 20
$SearchName.location            = New-Object System.Drawing.Point(150,200)
$SearchName.Font                = 'Microsoft Sans Serif,10'
$SearchName.Visible             = $True
$SearchName.Add_KeyDown({ 
    if ($_.KeyCode -eq "Enter") 
    {    
    FinduserForm
    }


$StartForm.Controls.Add($SearchName)


#Buttons
$FinduserBtn                   = New-Object system.Windows.Forms.Button
$FinduserBtn.BackColor         = "#a4ba67"
$FinduserBtn.text              = "Find User"
$FinduserBtn.width             = 90
$FinduserBtn.height            = 30
$FinduserBtn.location          = New-Object System.Drawing.Point(75,100)
$FinduserBtn.Font              = 'Microsoft Sans Serif,10'
$FinduserBtn.ForeColor         = "#ffffff"
$StartForm.CancelButton   = $cancelBtn
$StartForm.Controls.Add($FinduserBtn)

$FinduserBtn.Add_Click({ FinduserForm })

#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Cancel"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(75,200)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000fff"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$StartForm.CancelButton   = $cancelBtn
$StartForm.Controls.Add($cancelBtn)

$StartForm.controls.AddRange(@($TitleOperationChoice,$Description,$Status))
# Display the form
$result = $StartForm.ShowDialog()
}
#>

function FindUserForm { 
  Write-host "FindUserForm" -fore Yellow
  #Username to be used in code
  #$ID=get-adobject -filter 'cn -like $searchstring' |Out-GridView -PassThru
  $ID=$allusers  |Out-GridView -PassThru
  $CN = $ID.Name

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
$CompleteMoveForm.close() 
}

Function MoveConfirmForm1{
	# Move Status from FinaliseMoveForm
	Write-host "MoveConfirmForm"
$MoveConfirmForm1                    = New-Object system.Windows.Forms.Form
$MoveConfirmForm1.ClientSize         = '400,100'
$MoveConfirmForm1.text               = "Move Status"
$MoveConfirmForm1.BackColor          = "#bababa"

#Account Name Heading
$MoveConfirmText1                           = New-Object system.Windows.Forms.Label
$MoveConfirmText1.text                      = "Mailbox: " + $ID.name
$MoveConfirmText1.AutoSize                  = $true
$MoveConfirmText1.width                     = 25
$MoveConfirmText1.height                    = 10
#$MoveConfirmText.ForeColor                 = "#ff0000"
$MoveConfirmText1.location                  = New-Object System.Drawing.Point(20,10)
$MoveConfirmText1.Font                      = 'Microsoft Sans Serif,13'
$MoveConfirmForm1.controls.AddRange(@($MoveConfirmText1))

Write-host $ID

$MRS = Get-moverequest $ID.name |get-moverequestStatistics
$MRT = $MRS |select DisplayName,StatusDetail,PercentComplete

$MoveConfirmDetail1                           = New-Object system.Windows.Forms.Label
$MoveConfirmDetail1.text                      = "Status: " + $MRT.StatusDetail.Value + ", "  + $MRT.PercentComplete + "% Complete"
$MoveConfirmDetail1.AutoSize                  = $true
$MoveConfirmDetail1.width                     = 300
$MoveConfirmDetail1.height                    = 10
#$MoveConfirmDetail1.ForeColor                 = "#ff0000"
$MoveConfirmDetail1.location                  = New-Object System.Drawing.Point(20,40)
$MoveConfirmDetail1.Font                      = 'Microsoft Sans Serif,13'
$MoveConfirmForm1.controls.AddRange(@($MoveConfirmDetail1))



$MoveConfirmForm1.ShowDialog()
$CompleteMoveForm.close() 
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
$ExecBtn.BackColor         = "#a4ba67"
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
	If ($MC) {Remove-moverequest $ID.name -confirm $False -whatif
	MoveConfirmForm1}
	})

#Cancel Button
$cancelBtn4                       = New-Object system.Windows.Forms.Button
$cancelBtn4.BackColor             = "#ffffff"
$cancelBtn4.text                  = "Cancel"
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
$ExecBtn.BackColor         = "#a4ba67"
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
	If ($MC) {$Moves |where {$_.Status -eq "Completed"} |Remove-moverequest -confirm $False -whatif
	#MoveConfirmForm2
	#$moves = Get-moverequest
	}
	})

#Cancel Button
$cancelBtn4                       = New-Object system.Windows.Forms.Button
$cancelBtn4.BackColor             = "#ffffff"
$cancelBtn4.text                  = "Cancel"
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

Function BulkUserForm {
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('MyDocuments') }
$null = $FileBrowser.ShowDialog()	
$List = Import-csv $FileBrowser.FileName	
}

############Form Functions End

# Init PowerShell GUI
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$script =  $scriptPath + "\" + $MyInvocation.MyCommand.name 
chooseForm
