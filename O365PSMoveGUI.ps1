﻿# GUI to Manage Exchange Online PowerShell Based Move Requests .
#
#
#MUST be run from Exange online PowerShell
##Requires ACTIVEDIRECTORY PowerShell module. 
#You will need User Admin rights.
#Created by Cedric Abrahams - cedric@inobits.com
#
#Version 1.3 2021-01-05


#Connect-EXOPSSession
Write-Host "Getting Mailbox and Move details" -fore green
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


############Code Functions Start
#Set new Primary Address
function NewMove {
$SourceFunction = 1

  Write-host "SNP" $NewMove.text
  Write-host $ID.DistinguishedName
  
$NewEmail = $NewMove.Text
$checkmail = "*" + $NewEmail + "*"
Write-host "Checkmail: " $checkmail
$Check = $null

ResultsForm
}

<# Function RemoveAlias {
$SourceFunction = 3
Write-host $RemoveAlias.text
$NewEmail = $RemoveAlias.Text
If (!$RemoveAlias) {NoStartForm
}
Else {
if ( $Obj -eq "User") {
set-ADuser -Identity $CN -Remove @{Proxyaddresses="smtp:" + $RemoveAlias.text }
Write-Host "Alias Removed"
Write-host "Refreshing User Data"

$user = Get-ADObject $CN -Properties mail, proxyaddresses
ResultsForm
}

if ( $Obj -eq "Group"){
set-ADGroup -Identity $CN -remove @{Proxyaddresses="smtp:"+ $RemoveAlias.text}
Write-Host "Alias Removed"
Write-host "Refreshing User Data"

$user = Get-ADObject $CN -Properties mail, proxyaddresses
ResultsForm
}
}
$RemoveAliasForm.close()  
}
 #>
 
############Code Functions End 

############Form Functions Start

Function ChooseForm {
	Add-Type -AssemblyName System.Windows.Forms
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
$BulkUserBtn.Add_Click({FindUserForm})

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

$ExecBtn.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 2020-01-01 -whatif})

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

$ExecBtn1.Add_Click({New-MoveRequest -identity $ID.name -Remote -RemoteHostName autodiscover.medikredit.co.za -RemoteCredential $credential -TargetDeliveryDomain "medikredit.co.za" -AcceptLargeDataLoss -BadItemLimit 1000 -CompleteAfter 9999-01-01 -whatif})

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

function FinaliseMoveForm {
Write-Host "FinaliseMoveForm"
Add-Type -AssemblyName System.Windows.Forms
# Add Alias form
$CompleteMoveForm                    = New-Object system.Windows.Forms.Form
$CompleteMoveForm.ClientSize         = '600,200'
$CompleteMoveForm.text               = "Complete a Mailbox Move"
$CompleteMoveForm.BackColor          = "#bababa"


if ($Valid -eq 1)  { [void]$ResultForm2.Close() }
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


#Result Buttons
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#a4ba67"
$ExecBtn.text              = "Complete Move"
$ExecBtn.width             = 120
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,40)
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
	If ($MC) {set-moverequest $ID.name -completeafter 2020-01-01 -whatif}
	})

#Cancel Button
$cancelBtn4                       = New-Object system.Windows.Forms.Button
$cancelBtn4.BackColor             = "#ffffff"
$cancelBtn4.text                  = "Cancel"
$cancelBtn4.width                 = 120
$cancelBtn4.height                = 30
$cancelBtn4.location              = New-Object System.Drawing.Point(20,80)
$cancelBtn4.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn4.ForeColor             = "#000fff"
$cancelBtn4.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$CompleteMoveForm.CancelButton   = $cancelBtn4
$CompleteMoveForm.Controls.Add($cancelBtn4)
$cancelBtn4.Add_Click({ $CompleteMoveForm.close() })

# Display the form
$result = $CompleteMoveForm.ShowDialog()
}

<# 
function DeleteAliasForm {
	Add-Type -AssemblyName System.Windows.Forms
$user = Get-ADObject $CN -Properties mail, proxyaddresses
# Result form
$RemoveAliasForm                    = New-Object system.Windows.Forms.Form
$RemoveAliasForm.ClientSize         = '500,600'
$RemoveAliasForm.text               = "Email Result Management"
$RemoveAliasForm.BackColor          = "#bababa"

if ($Valid -eq 1)  { [void]$ResultForm2.Close() }
########### Result Form cont.
#Account Name Heading
$RemoveAliasText                           = New-Object system.Windows.Forms.Label
$RemoveAliasText.text                      = 'You have chosen to edit user:'
$RemoveAliasText.AutoSize                  = $true
$RemoveAliasText.width                     = 25
$RemoveAliasText.height                    = 10
$RemoveAliasText.location                  = New-Object System.Drawing.Point(20,10)
$RemoveAliasText.Font                      = 'Microsoft Sans Serif,13'
$RemoveAliasForm.controls.AddRange(@($RemoveAliasText))

# Account Name.
$ResultChoice                           = New-Object system.Windows.Forms.Label
$ResultChoice.text                      = $SearchName.Text + " ("+ $user.Name +")"
$ResultChoice.AutoSize                  = $true
$ResultChoice.width                     = 25
$ResultChoice.height                    = 10
$ResultChoice.location                  = New-Object System.Drawing.Point(20,35)
$ResultChoice.Font                      = 'Microsoft Sans Serif,13,style=bold'
$RemoveAliasForm.controls.AddRange(@($ResultChoice))

# Show Email Address Heading.
$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Current Primary Email Address:'
$EmailTitle.AutoSize                  = $true
$EmailTitle.width                     = 25
$EmailTitle.height                    = 10
$EmailTitle.location                  = New-Object System.Drawing.Point(20,75)
$EmailTitle.Font                      = 'Microsoft Sans Serif,13'
$RemoveAliasForm.controls.AddRange(@($EmailTitle))



# Show Email Address
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $user.mail
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
# Position the element
$EmailChoice.location                  = New-Object System.Drawing.Point(20,100)
# Define the font type and size
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$RemoveAliasForm.controls.AddRange(@($EmailChoice))



# Show Proxy Addresses Heading.
$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Current Alias Addresses:'
$EmailTitle.AutoSize                  = $true
$EmailTitle.width                     = 25
$EmailTitle.height                    = 10
# Position the element
$EmailTitle.location                  = New-Object System.Drawing.Point(20,125)
# Define the font type and size
$EmailTitle.Font                      = 'Microsoft Sans Serif,13'
$RemoveAliasForm.controls.AddRange(@($EmailTitle))

$VL = 150
$PRcount = 0
Foreach ($P in $user.proxyaddresses) {
If ($P -clike "*smtp*") {
$PRcount = $PRcount + 1
$Pr = $P 
$PR =$PR -replace 'smtp:',''
if ($PRcount -eq 1) {$prx1 = $PR}
if ($PRcount -eq 2) {$prx2 = $PR}
if ($PRcount -eq 3) {$prx3 = $PR}
if ($PRcount -eq 4) {$prx4 = $PR}
if ($PRcount -eq 5) {$prx5 = $PR}
if ($PRcount -eq 6) {$prx6 = $PR}
if ($PRcount -eq 7) {$prx7 = $PR}
if ($PRcount -eq 8) {$prx8 = $PR}
if ($PRcount -eq 9) {$prx9 = $PR}
if ($PRcount -eq 10) {$prx10 = $PR}
if ($PRcount -eq 11) {$prx11 = $PR}
if ($PRcount -eq 12) {$prx12 = $PR}
if ($PRcount -gt 3) {$prxExtra1 = "Only first 12 Addresses Shown. Type address to remove"}


# Show Proxy Addresses
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $PR
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
# Position the element
$EmailChoice.location                  = New-Object System.Drawing.Point(20,$VL)
# Define the font type and size
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$RemoveAliasForm.controls.AddRange(@($EmailChoice))
$VL = $VL +25
}
else {}
}
#NO Email TextBoxLable
$NoEmailLable               = New-Object system.Windows.Forms.Label
$NoEmailLable.text           = $NP
$NoEmailLable.AutoSize       = $true
$NoEmailLable.width          = 25
$NoEmailLable.height         = 20
$NoEmailLable.location       = New-Object System.Drawing.Point(20,475)
$NoEmailLable.Font           = 'Microsoft Sans Serif,16,style=Bold'
$NoEmailLable.ForeColor      = "#ff0000"
$NoEmailLable.Visible        = $True
$RemoveAliasForm.Controls.Add($NoEmailLable)

#New Email TextBoxLable
$RemoveAliasLabel                = New-Object system.Windows.Forms.Label
$RemoveAliasLabel.text           = "Select the Alias Email Address to be removed:"
$RemoveAliasLabel.AutoSize       = $true
$RemoveAliasLabel.width          = 250
$RemoveAliasLabel.height         = 20
$RemoveAliasLabel.location       = New-Object System.Drawing.Point(20,500)
$RemoveAliasLabel.Font           = 'Microsoft Sans Serif,10,style=Bold'
$RemoveAliasLabel.Visible        = $True
$RemoveAliasForm.Controls.Add($RemoveAliasLabel)

#New Email TextBox
$RemoveAlias                     = New-Object system.Windows.Forms.ComboBox
$RemoveAlias.text                = "Choose"
$RemoveAlias.width               = 400
$RemoveAlias.autosize            = $true
$RemoveAlias.Visible             = $true  
      
# Add the items in the dropdown list
@($prx1,$prx2,$Prx3,$Prx4,$prx5,$prx6,$prx7,$prx8,$prx9,$prx10,$prx11,$prx12,$prxExtra1) | ForEach-Object {if($_) {[void] $RemoveAlias.Items.Add($_)}}
# Select the default value
$RemoveAlias.SelectedIndex       = 0
$RemoveAlias.location            = New-Object System.Drawing.Point(20,525)
$RemoveAlias.Font                = 'Microsoft Sans Serif,10'
$RemoveAliasForm.Controls.Add($RemoveAlias)
Write-host $RemoveAlias.text
$RemoveAlias.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {    
    RemoveAlias
    }
    })

#Result Buttons
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#FF0000"
$ExecBtn.text              = "Remove Alias"
$ExecBtn.width             = 150
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,560)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$RemoveAliasForm.CancelButton   = $cancelBtn5
$RemoveAliasForm.Controls.Add($ExecBtn)
$ExecBtn.Add_Click({ RemoveAlias })

#Cancel Button
$cancelBtn5                       = New-Object system.Windows.Forms.Button
$cancelBtn5.BackColor             = "#ffffff"
$cancelBtn5.text                  = "Cancel"
$cancelBtn5.width                 = 90
$cancelBtn5.height                = 30
$cancelBtn5.location              = New-Object System.Drawing.Point(400,560)
$cancelBtn5.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn5.ForeColor             = "#000fff"
$cancelBtn5.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$RemoveAliasForm.CancelButton   = $cancelBtn5
$RemoveAliasForm.Controls.Add($cancelBtn5)
$cancelBtn5.Add_Click({ $RemoveAliasForm.close() })

# Display the form
$result = $RemoveAliasForm.ShowDialog()
}

function CheckEmailForm {
ResultsForm
Write-host "Check Existing Addresses"
}
 #>

Function ResultsForm {
if ($user -ne "None") {Write-host "Check Address" $NewEmail -ForegroundColor Green
$res = $true

$user = get-adobject $CN -Properties mail,proxyaddresses
Write-host "Resultsform"

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

Function ActionForm {

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
if ($OperationChoice.text -eq $ViewMoveForm) {DeleteAliasForm}
if ($OperationChoice.text -eq $RemoveMoveForm) {CheckEmailForm}
#if ($OperationChoice.text -eq $sub5) {Sub5}
    }
})
$ActionForm.Controls.Add($OperationChoice)

$Name = $ID.name
If (!$name) {$Name = "No User Selected"}
#TextBoxLable
$SearchNameLabel                = New-Object system.Windows.Forms.Label
$SearchNameLabel.text           = "You are managing user"
$SearchNameLabel.AutoSize       = $true
$SearchNameLabel.width          = 25
$SearchNameLabel.height         = 20
$SearchNameLabel.location       = New-Object System.Drawing.Point(20,80)
$SearchNameLabel.Font           = 'Microsoft Sans Serif,14,style=Bold'
$SearchNameLabel.Visible        = $True
$ActionForm.Controls.Add($SearchNameLabel)

#TextBoxLable
$NameLable                = New-Object system.Windows.Forms.Label
$NameLable.text           = $Name
$NameLable.AutoSize       = $true
$NameLable.width          = 25
$NameLable.height         = 20
#$NameLable.ForeColor 	  = "#0000ff"
$NameLable.location       = New-Object System.Drawing.Point(40,80)
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
if ($OperationChoice.text -eq $ViewMoveForm) {DeleteAliasForm}
if ($OperationChoice.text -eq $RemoveMoveForm) {CheckEmailForm}
#if ($OperationChoice.text -eq $sub5) {Sub5}
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

<# 
Function UsedStartForm {
Add-Type -AssemblyName System.Windows.Forms
# Result form2
$UsedEmailForm                    = New-Object system.Windows.Forms.Form
$UsedEmailForm.ClientSize         = '400,100'
$UsedEmailForm.text               = "Invalid User"
$UsedEmailForm.BackColor          = "#bababa"

#Account Name Heading
$UsedEmailText                          = New-Object system.Windows.Forms.Label
$UsedEmailText.text                      = 'Email address is already in use.'
$UsedEmailText.AutoSize                  = $true
$UsedEmailText.width                     = 25
$UsedEmailText.height                    = 10
$UsedEmailText.ForeColor                 = "#ff0000"
$UsedEmailText.location                  = New-Object System.Drawing.Point(20,10)
$UsedEmailText.Font                      = 'Microsoft Sans Serif,13'
$UsedEmailForm.controls.AddRange(@($UsedEmailText))
$UsedEmailForm.ShowDialog()
$OperationChoice.text =''
}

Function NoStartForm {
Add-Type -AssemblyName System.Windows.Forms
# Result form2
$NoEmailForm                    = New-Object system.Windows.Forms.Form
$NoEmailForm.ClientSize         = '400,100'
$NoEmailForm.text               = "Invalid User"
$NoEmailForm.BackColor          = "#bababa"

#Account Name Heading
$NoEmailText                          = New-Object system.Windows.Forms.Label
$NoEmailText.text                      = 'No Email Address was Entered.'
$NoEmailText.AutoSize                  = $true
$NoEmailText.width                     = 25
$NoEmailText.height                    = 10
$NoEmailText.ForeColor                 = "#ff0000"
$NoEmailText.location                  = New-Object System.Drawing.Point(20,10)
$NoEmailText.Font                      = 'Microsoft Sans Serif,13'
$NoEmailForm.controls.AddRange(@($NoEmailText))
$NoEmailForm.ShowDialog()
$OperationChoice.text =''
}
 #>

Function InvalidUserForm {
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

Function NoMoveForm {
write-host "NoMoveForm"
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

Function StartForm {
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
$Status.Font              = 'Microsoft Sans Serif,10' #>


<# #TextBoxLable
$SearchNameLabel                = New-Object system.Windows.Forms.Label
$SearchNameLabel.text           = "Search for name: "
$SearchNameLabel.AutoSize       = $true
$SearchNameLabel.width          = 25
$SearchNameLabel.height         = 20
$SearchNameLabel.location       = New-Object System.Drawing.Point(20,200)
$SearchNameLabel.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SearchNameLabel.Visible        = $True
$StartForm.Controls.Add($SearchNameLabel) #>

<# #TextBox
$SearchName                     = New-Object system.Windows.Forms.TextBox
$SearchName.multiline           = $false
$SearchName.width               = 314
$SearchName.height              = 20
$SearchName.location            = New-Object System.Drawing.Point(150,200)
$SearchName.Font                = 'Microsoft Sans Serif,10'
$SearchName.Visible             = $True
$SearchName.Add_KeyDown({ #>
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

 ################ Start Window ###############
function FindUserForm { 
  Write-host "Select User" -fore Yellow
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

############Form Functions End


# Init PowerShell GUI
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$script =  $scriptPath + "\" + $MyInvocation.MyCommand.name 
chooseForm