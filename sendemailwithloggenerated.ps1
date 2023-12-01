Add-Type -AssemblyName System.Windows.Forms
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# Declare $sentToday as a global variable
$global:sentToday = @{}

# Get the PC name
$pcName = [System.Environment]::MachineName

# Declare the log file path globally
$global:logFileName = "Email_Log_" + (Get-Date -Format "yyyyMMdd") + ".txt"
$global:logFilePath = "C:\$global:logFileName"

if (-not (Test-Path -Path $global:logFilePath)) {
    # If log file doesn't exist, create a new log file
    $global:logFilePath = New-Item -Path $global:logFilePath -ItemType "file" -Force
}

# Default sender's email
$defaultSenderEmail = "attachment-only@hotmail.com"

# Default email Subject
$defaultEmailSubject = "Attachment files from {0}" -f $pcName # Subject now includes PC name

# Create a form for user input
$form = New-Object System.Windows.Forms.Form
$form.Text = "Send Email from $pcName"
$form.Size = New-Object System.Drawing.Size(400, 480)
$form.StartPosition = "CenterScreen"

# Function to create label controls
function CreateLabel($text, $x, $y) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point($x, $y)
    $label.Font = New-Object System.Drawing.Font("Arial", 9)
    $form.Controls.Add($label)
}

# Function to create textbox controls
function CreateTextBox($x, $y, $readonly = $false, $defaultText = "") {
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(360, 25)
    $textBox.Font = New-Object System.Drawing.Font("Arial", 9)
    $textBox.Location = New-Object System.Drawing.Point($x, $y)
    $textBox.ReadOnly = $readonly
    $textBox.TabStop = $false  # Disable tab focus
    $textBox.Text = $defaultText
    $form.Controls.Add($textBox)
    return $textBox
}

# Create controls for sender, recipient, subject, attachments
CreateLabel "Sender's Email Address:" 10 10
$textBoxSender = CreateTextBox 10 30 -readonly $true -defaultText $defaultSenderEmail

CreateLabel "Recipient's Email Address:" 10 60
$textBoxRecipient = CreateTextBox 10 80

CreateLabel "Email Subject:" 10 110
$textBoxSubject = CreateTextBox 10 130 -readonly $true -defaultText $defaultEmailSubject

$buttonAttachment = New-Object System.Windows.Forms.Button
$buttonAttachment.Location = New-Object System.Drawing.Point(10, 180)
$buttonAttachment.Size = New-Object System.Drawing.Size(180, 30)
$buttonAttachment.Text = "Select Attachments"
$form.Controls.Add($buttonAttachment)

$listBoxAttachments = New-Object System.Windows.Forms.ListBox
$listBoxAttachments.Location = New-Object System.Drawing.Point(10, 230)
$listBoxAttachments.Size = New-Object System.Drawing.Size(360, 100)
$listBoxAttachments.Font = New-Object System.Drawing.Font("Arial", 9)
$listBoxAttachments.SelectionMode = "MultiExtended"
$form.Controls.Add($listBoxAttachments)

$buttonRemoveAttachment = New-Object System.Windows.Forms.Button
$buttonRemoveAttachment.Location = New-Object System.Drawing.Point(200, 180)
$buttonRemoveAttachment.Size = New-Object System.Drawing.Size(170, 30)
$buttonRemoveAttachment.Text = "Remove Selected Attachments"
$form.Controls.Add($buttonRemoveAttachment)

$buttonSend = New-Object System.Windows.Forms.Button
$buttonSend.Location = New-Object System.Drawing.Point(10, 350)
$buttonSend.Size = New-Object System.Drawing.Size(180, 30)
$buttonSend.Text = "Send Email"
$form.Controls.Add($buttonSend)

$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(200, 350)
$buttonCancel.Size = New-Object System.Drawing.Size(170, 30)
$buttonCancel.Text = "Cancel"
$buttonCancel.Add_Click({
    $form.Close()
})
$form.Controls.Add($buttonCancel)

# Handler for attachment button click
$buttonAttachment.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Multiselect = $true
    $openFileDialog.Filter = "Docs (*.pdf *.csv *.zip *.rar *.xlsx *.xls) | *.pdf;*.csv;*.zip;*.rar;*.xlsx;*.xls"
    
    $result = $openFileDialog.ShowDialog()
    
    if ($result -eq "OK") {
        $attachmentPaths = $openFileDialog.FileNames
        $listBoxAttachments.Items.AddRange($attachmentPaths)
    }
})

# Handler for remove attachment button click
$buttonRemoveAttachment.Add_Click({
    $selectedIndices = $listBoxAttachments.SelectedIndices
    if ($selectedIndices.Count -gt 0) {
        for ($i = $selectedIndices.Count - 1; $i -ge 0; $i--) {
            $listBoxAttachments.Items.RemoveAt($selectedIndices[$i])
        }
    }
})

# Declare $sentToday as a global variable
$global:sentToday = @{}

# Handler for send button click
$buttonSend.Add_Click({
    $recipient = $textBoxRecipient.Text.Trim()
    
    if (-not $recipient) {
        [System.Windows.Forms.MessageBox]::Show("Please provide a valid recipient email address.", "Recipient Missing", "OK", [System.Windows.Forms.MessageBoxIcon]::Warning)
        return  # Stop further execution
    }

    if ($recipient -notmatch '@') {
        [System.Windows.Forms.MessageBox]::Show("Recipient email address must contain the '@' symbol.", "Invalid Format", "OK", [System.Windows.Forms.MessageBoxIcon]::Warning)
        return  # Stop further execution
    }

    if (-not $global:sentToday.ContainsKey($recipient)) {
        $global:sentToday[$recipient] = 0
    }

    # Read log file and count occurrences of the recipient's email
    $logContent = Get-Content $logFilePath
    $occurrences = ($logContent | Where-Object {$_ -match $recipient}).Count

    # Update count based on occurrences in the log
    $global:sentToday[$recipient] += $occurrences

    if ($global:sentToday[$recipient] -lt 3) {  
        try {
            $smtpClient = New-Object System.Net.Mail.SmtpClient
            $smtpClient.Host = "smtp.office365.com"
            $smtpClient.Port = 587
            $smtpClient.EnableSsl = $true
            $credentials = New-Object System.Net.NetworkCredential($defaultSenderEmail, "attachment2008")
            $smtpClient.Credentials = $credentials

            $mail = New-Object System.Net.Mail.MailMessage
            $mail.From = $defaultSenderEmail
            $mail.To.Add($recipient)
            $mail.Subject = $defaultEmailSubject
            $mail.Body = @"
Please find multiple attachments.

This message and any attachments are confidential and may be privileged or otherwise protected from disclosure. It is intended solely for the named recipient(s) and may not be used, copied, disclosed, or distributed in any way. If you are not an intended recipient, please notify the sender and delete this email immediately.

Please note that attachments to this email may contain viruses or other malicious software. It is the recipient's responsibility to scan for any viruses or malware. The sender does not accept liability for any damage caused by any virus or malware transmitted by this email.

Thank you.
"@

            foreach ($item in $listBoxAttachments.Items) {
                $attachment = New-Object System.Net.Mail.Attachment($item)
                $mail.Attachments.Add($attachment)
            }

            $attachmentList = ($listBoxAttachments.Items | ForEach-Object { $_ }) -join ","
            
            $smtpClient.Send($mail)
            
            # Increment the count only after successful email sending
            $global:sentToday[$recipient]++
            
            Write-Host "Email sent successfully! Emails sent today to $recipient : $($global:sentToday[$recipient])"

            $emailStatus = "Email sent successfully!"
            $logEntry = "`nSent Date & Time: {0}`nRecipient Email: {1}`nRunning On: {2}`nStatus: {3}`nAttachments:{4}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $recipient, $pcName, $emailStatus, $attachmentList
            $logEntry | Out-File -FilePath $global:logFilePath -Append
            [System.Windows.Forms.MessageBox]::Show("Email sent successfully!", "Success", "OK", [System.Windows.Forms.MessageBoxIcon]::Information)
            
            $listBoxAttachments.Items.Clear()
            $textBoxRecipient.Text = ""
        
        } catch {
            Write-Host "Failed to send email to $recipient. Emails sent today: $($global:sentToday[$recipient])"
            $emailStatus = "Failed to send email. Please check the SMTP configuration and try again."
            $logEntry = "`nSent Date & Time: {0}`nRecipient Email: {1}`nRunning On: {2}`nStatus: {3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $recipient, $pcName, $emailStatus
            $logEntry | Out-File -FilePath $global:logFilePath -Append
            [System.Windows.Forms.MessageBox]::Show("Failed to send email. Please check the SMTP configuration and try again.", "Error", "OK", [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        # When limit is reached for a recipient
        $emailStatus = "Limit reached! You can send a maximum of 3 emails to this recipient today."
        $logEntry = "`nSent Date & Time: {0}`nRecipient Email: {1}`nRunning On: {2}`nStatus: {3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $recipient, $pcName, $emailStatus
        $logEntry | Out-File -FilePath $global:logFilePath -Append

        # Display a warning about the limit being reached
        [System.Windows.Forms.MessageBox]::Show("Limit reached! You can send a maximum of 3 emails to this recipient today.", "Limit Exceeded", "OK", [System.Windows.Forms.MessageBoxIcon]::Warning)

        # Clear attachments and recipient box after reaching the limit
        $listBoxAttachments.Items.Clear()
        $textBoxRecipient.Text = ""
    }
})


# Function to calculate total size of attachments
function GetTotalAttachmentSize($attachmentPaths) {
    $totalSize = 0
    foreach ($file in $attachmentPaths) {
        $totalSize += (Get-Item $file).Length
    }
    return $totalSize
}

# Show the form
$form.ShowDialog()
