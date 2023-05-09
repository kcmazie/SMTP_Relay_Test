<#==============================================================================
         File Name : SMTP_Relay_Test.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com )
                   :
       Description : This script will send an email to the selected user,
                   : SMTP host, and SMTP port.
                   :
                   : The target email and server IP are stored and recalled for
                   : subsequent runs.
                   :
             Notes : Normal operation is with no command line options.
                   :
          Warnings : None
                   :
    Last Update by : Kenneth C. Mazie
                   : #>
$CurrentVersion = "3.00"
                <# :
   Version History : v1.00 - 06-26-12 - Original
    Change History : v2.00 - 06-29-18 - Rewrite. Changed save file format. Edited GUI.
                   : v3.00 - 05-09-23 - Cleaned up older code.  Adjusted font.  Updated for publishing.
                   :
===============================================================================#>
<#PSScriptInfo
.VERSION 3.00
.AUTHOR Kenneth C. Mazie (kcmjr AT kcmjr.com)
.DESCRIPTION
GUI driven SMTP test. Sends a quick and dirty test email to the recipient of your choice via the SMTP server of your choice.
The target email and server are stored to pre-populate the script for future runs.
#>

Clear-Host

#--[ Functions ]--
Function ProcessMessage {
    New-Item "$PSScriptRoot\eMailRecipient.txt" -type file -Value "$eMailRecipient@$eMailDomain" -Force -Confirm:$false         #--[ Who to send status email to.
    New-Item "$PSScriptRoot\eMailServer.txt" -type file -Value "$SMTPserver" -Force -Confirm:$false                             #--[ What server to send through.
    
    $MailMessage = New-Object System.Net.Mail.MailMessage
    $MailMessage.IsBodyHtml = $true
    $SMTPserverClient = New-Object System.Net.Mail.smtpClient
    $SMTPserverClient.host = $SMTPserver
    If ($P465RadioButton.checked){
        $SMTPport = 465
        $SMTPserverClient.EnableSsl = $true
        $SMTPserverClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
    }
    If ($P25RadioButton.checked){
        $SMTPport = 25
        $SMTPserverClient.EnableSsl = $false
        $SMTPserverClient.Credentials = $null 
    }
        
    #--[ Build Message Body ]--
    $strBody = "
    <html><body>
    <p>This is a SMTP test email. Below are the details specific to this test email:
    <p>
    <table border='1'>
        <TR>
            <th WIDTH='50%'>Originating Computer</th>
            <th WIDTH='50%'>$localComputer</th>
        </TR>
        <TR>
            <th WIDTH='50%'>SMTP Server Address</th>
            <th WIDTH='50%'>$SMTPserver</th>
        </TR>
        <TR>
            <th WIDTH='50%'>SMTP Server Port</th>
            <th WIDTH='50%'>$SMTPport</th>
        </TR>
        <TR>
            <th WIDTH='50%'>Sender</th>
            <th WIDTH='50%'>SMTPtest@$eMailDomain</th>
        </TR>
        <TR>
            <th WIDTH='50%'>Recipient</th>
            <th WIDTH='50%'>$eMailRecipient@$eMailDomain</th>
        </TR>
            <TR>
            <th WIDTH='50%'>Date / Time</th>
            <th WIDTH='50%'>$date</th>
        </TR>
    </table>
    <p> The SMTP test has <b>COMPLETED SUCCESSFULLY</b> if you receive this email!
    </body></html>
    "
    
    $Recipient = New-Object System.Net.Mail.MailAddress("$eMailRecipient@$eMailDomain")
    $Sender = New-Object System.Net.Mail.MailAddress("SMTPtest@$eMailDomain")
    $SMTPserverClient.port = $SMTPport
    $MailMessage.Subject = "Test SMTP email over port $SMTPport from $SMTPserver"
    $MailMessage.Sender = $Sender
    $MailMessage.From = $Sender
    $MailMessage.To.add($Recipient)
    $MailMessage.Body = $strBody
    
    Try {
        $SMTPserverClient.Send($MailMessage)
    } Catch {
        #[System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”)
#        [Windows.Forms.MessageBox]::Show($Error[0], “Error Sending Email”, [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
    }
    $objForm.Close()
    }

Function IsThereText ($TargetBox){
    if (($TargetBox.Text.Length -ne 0)){ # -or ($Script:FileNameTextBox.Text.Length -ne 0)){
        Return $true
    }else{
        Return $false
    }
  }

[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-null  
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

#--[ Variables ]--
[string]$Username = "user@domain.com"                                            #--[ SSL authentication user. Adjust per environment
[string]$Password = "password"                                                   #--[ SSL password. Adjust per environment

$ErrorActionPreference = "Stop" #ilentlyContinue"

if (Test-Path "$PSScriptRoot\SMTPconfig.txt") {
    $SMTPconfig = Get-Content "$PSScriptRoot\SMTPconfig.txt"              #--[ This will be stored in the script folder location ]--
    [string]$eMailRecipient = $SMTPconfig.Split(',')[0]
    [string]$eMailDomain = $SMTPconfig.Split(',')[1] 
    [string]$SMTPserver = $SMTPconfig.Split(',')[2] 
}else{
    [string]$eMailRecipient = ""
    [string]$eMailDomain = ""
    [string]$SMTPserver = ""
}

$date = Get-Date
[int]$FormWidth = 360
[int]$FormHeight = 300
[int]$ButtonLeft = 55
[int]$BoxHeight = 22

#--[ FQDN of the local computer ]--
$objCompSys = Get-WmiObject win32_computersystem
$localComputer = $objCompSys.name+"."+$objCompSys.domain
Remove-Variable objCompSys

#-----------------------------[ Main Process ]----------------------------------
#--[ Create Form ]--
$objForm = new-object System.Windows.Forms.form
$objForm.Text = "SMTP Test Script"
$objForm.size = new-object System.Drawing.Size($FormWidth,$FormHeight)
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $true
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$Recipient=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$objForm.Close();$Stop = $true}})

#-------------------------------------------------------------
#--[ Add Form Lable ]--
$RowLocation = 5 
$objFormLabelBox = new-object System.Windows.Forms.Label
$objFormLabelBox.Font = new-object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Bold)
$objFormLabelBox.Location = new-object System.Drawing.Size(18,$RowLocation)
$objFormLabelBox.size = new-object System.Drawing.Size(400,$BoxHeight)
$objFormLabelBox.Text = "Select SMTP test parameters to use:"
$objForm.Controls.Add($objFormLabelBox)

#--[ Add Email Form Label ]--
$RowLocation = $RowLocation+29 
$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Point(20,$RowLocation) 
$objLabel1.Size = New-Object System.Drawing.Size(280,29) 
$objLabel1.Text = "Enter the email address where results should be sent:"
$objLabel1.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$objForm.Controls.Add($objLabel1) 

#-------------------------------------------------------------
#--[ Add Email Text Input Box ]--
$RowLocation = $RowLocation+36 
$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Location = New-Object System.Drawing.Size(45,$RowLocation) 
$objTextBox1.Size = New-Object System.Drawing.Size(100,$BoxHeight) 
$objTextBox1.Text = $eMailRecipient
$objTextBox1.add_TextChanged({
    If (IsThereText $objTextBox1){
        $objProcessButton.Enabled = $True
    }Else{
        $objProcessButton.Enabled = $False
    }
})    
$objForm.Controls.Add($objTextBox1) 

#--[ Add Email @ Label ]--
$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Point(153,$RowLocation) 
$objLabel2.Size = New-Object System.Drawing.Size(24,25) 
$objLabel2.Text = "@"
$objForm.Controls.Add($objLabel2) 

#--[ Add Email Domain Input Box ]--
$objTextBox2 = New-Object System.Windows.Forms.TextBox
$objTextBox2.Location = New-Object System.Drawing.Point(180,$RowLocation) 
$objTextBox2.Size = New-Object System.Drawing.Size(100,20) 
$objTextBox2.Text = $eMailDomain
$objForm.Controls.Add($objTextBox2) 

#-------------------------------------------------------------
#--[ Add SMTP Host Form Label ]--
$RowLocation = $RowLocation+27 
$objLabel3 = New-Object System.Windows.Forms.Label
$objLabel3.Location = New-Object System.Drawing.Point(20,$RowLocation) 
$objLabel3.Size = New-Object System.Drawing.Size(280,20) 
$objLabel3.Text = "Enter the SMTP server to use:"
$objLabel3.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$objForm.Controls.Add($objLabel3) 

#--[ Add SMTP Host Input Box ]--
$RowLocation = $RowLocation+27 
$objTextBox3 = New-Object System.Windows.Forms.TextBox 
$objTextBox3.Location = New-Object System.Drawing.Size(110,$RowLocation)
$objTextBox3.Size = New-Object System.Drawing.Size(110,50) 
$objTextBox3.Text = $SMTPserver 
$objForm.Controls.Add($objTextBox3)

#-------------------------------------------------------------
#--[ Add port 25 Radio Button ]--
$RowLocation = $RowLocation+25 
$P25RadioButton = New-Object Windows.Forms.radiobutton
$P25RadioButton.text = "Use Port 25 (anonymous connection)"
$P25RadioButton.height = 26
$P25RadioButton.width = 270
$P25RadioButton.top = $RowLocation
$P25RadioButton.left = 38
$P25RadioButton.checked = $true
$objForm.controls.add($P25RadioButton)

#--[ Add SSL Radio Button ]--
$RowLocation = $RowLocation+27
$P465RadioButton = New-Object Windows.Forms.radiobutton
$P465RadioButton.text = "Use Port 465 (SSL connection)"
$P465RadioButton.height = 26
$P465RadioButton.width = 270
$P465RadioButton.top = $RowLocation
$P465RadioButton.left = 58
$objForm.controls.add($P465RadioButton)

#-------------------------------------------------------------
#--[ Add Send Button ]--
$RowLocation = $RowLocation+32
$RowTop = $FormHeight-90
$objProcessButton = new-object System.Windows.Forms.Button
If (!([string]::IsNullOrEmpty($eMailRecipient))){
    write-host "not null"    
        $objProcessButton.Enabled = $True    
    }Else{
$objProcessButton.Enabled = $false}
$objProcessButton.Location = new-object System.Drawing.Size($ButtonLeft,$RowLocation)
$objProcessButton.Size = new-object System.Drawing.Size(100,25)
$objProcessButton.Text = "Send Email"
$objProcessButton.Add_Click({
    $SMTPserver = $objTextBox3.Text
    $eMailRecipient = $objTextBox1.Text
    $eMailDomain = $objTextBox2.Text
    "$eMailRecipient,$eMailDomain,$SMTPserver" | Out-File -FilePath "$PSScriptRoot\SMTPconfig.txt" 
    ProcessMessage
    $objForm.Close()
})
$objForm.Controls.Add($objProcessButton)

#--[ Add Exit Button ]--
$objCloseButton = new-object System.Windows.Forms.Button
$objCloseButton.Location = new-object System.Drawing.Size(($ButtonLeft+125),$RowLocation)
$objCloseButton.Size = new-object System.Drawing.Size(100,25)
$objCloseButton.Text = "Cancel/Close"
$objCloseButton.Add_Click({$objForm.close()})
$objForm.Controls.Add($objCloseButton)

#--[ Open Form ]--
$objForm.topmost = $true
$objForm.Add_Shown({$objForm.Activate()})
$objForm.ShowDialog()

