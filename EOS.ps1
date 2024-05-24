# Adding a reference to the .NET WPF assembly [PresentationFramework]. More modern UI than previous design where we used [System.Windows.Forms] which is from 2002.
# [Microsoft.Windows.AppSDK] would be even more modern, but it would require using separate XAML for the visual GUI elements while we prefer to keep all our eggs in one henhouse for this script.
Add-Type -AssemblyName PresentationFramework

# Some global scoped variables as well as initializing the variable to hold attachments. 
# itemNames[] is an array of the headers for each EOS item.
# descriptionList{} is a hashtable to map the descriptions to the associated item key using itemNames[].
$attachment = $null
$itemNames = @("OEE/Downtime", "Spot Repair", "PDI", "VES", "Unit Review", "New Material/New Equipment", "Special Request")
$descriptionList = @{}

# Get today's date
$today = Get-Date -Format "MM/dd/yyyy"

# DEFAULT CONTACTS LIST: UPDATE AS NEEDED
# For temporary additional contacts, use the "Add New..." button on the dialog window and select "Contact". 
$defaults = defaultsfile

# Function to prompt user for EOS item descriptions for the day
function Get-Descriptions {
    param ([string[]]$itemNames)

    # Building the initial window
    $window = New-Object -TypeName System.Windows.Window
    $window.Title = "PL1 Engineering End-of-Shift Reporting"
    $window.SizeToContent = "WidthAndHeight"
    $window.WindowStartupLocation = "CenterScreen"
    $window.ResizeMode = "NoResize"
    # $window.Background = [Windows.Media.Brushes]::ColorName

    $grid = New-Object -TypeName System.Windows.Controls.Grid
    $grid.Width = 575
    $grid.Height = 315
    $window.Content = $grid

    $tabControl = New-Object -TypeName System.Windows.Controls.TabControl
    $tabControl.FontSize = 12
    # $tabControl.Background = [Windows.Media.Brushes]::ColorName
    $grid.Children.Add($tabControl)

    # Generating some tabs to hold description textboxes for each EOS follow-up item
    foreach ($itemName in $itemNames) {
        $tabItem = New-Object -TypeName System.Windows.Controls.TabItem
        $tabItem.Header = $itemName
        $tabControl.Items.Add($tabItem)

        $stackPanel = New-Object -TypeName System.Windows.Controls.StackPanel
        $stackPanel.Orientation = "Vertical"
        $tabItem.Content = $stackPanel

        $label = New-Object -TypeName System.Windows.Controls.Label
        $label.Content = "Description for $($itemName):"
        # $label.Background = [Windows.Media.Brushes]::ColorName
        $stackPanel.Children.Add($label)

        $textBox = New-Object -TypeName System.Windows.Controls.TextBox
        $textBox.AcceptsReturn = $true
        $textBox.Width = 567
        $textBox.Height = 220
        $textBox.VerticalScrollBarVisibility = "Auto"
        # $textBox.Background = [Windows.Media.Brushes]::ColorName
        $stackPanel.Children.Add($textBox)

    }

    # Stack panel for our buttons
    $buttonPanel = New-Object -TypeName System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Center"
    $buttonPanel.VerticalAlignment = "Bottom"
    $buttonPanel.Height = 35
    $buttonPanel.Margin = "5"
    # $buttonPanel.Background = [Windows.Media.Brushes]::ColorName
    $grid.Children.Add($buttonPanel)

    # Add New... button logic allowing the user to either choose to add attachments or new contacts
    $newButton = New-Object -TypeName System.Windows.Controls.Button
    $newButton.Content = "Add New..."
    $newButton.Width = 75
    $newButton.Margin = "5"

    # Define click event handler
    $newButton.Add_Click({
        $newDialog = New-Object -TypeName System.Windows.Window
        $newDialog.Title = "Add New..."
        $newDialog.Height = 86
        $newDialog.Width = 182
        $newDialog.WindowStartupLocation = "CenterScreen"
        $newDialog.ResizeMode = "NoResize"

        # Create a StackPanel to hold the buttons
        $newButtonPanel = New-Object -TypeName System.Windows.Controls.StackPanel
        $newButtonPanel.Orientation = "Horizontal"
        $newButtonPanel.Margin = "5"

        # Add Attachment button
        $attachmentButton = New-Object -TypeName System.Windows.Controls.Button
        $attachmentButton.Content = "Attachment"
        $attachmentButton.Height = 25
        $attachmentButton.Width = 75
        $attachmentButton.Margin = "5"
        $attachmentButton.Add_Click({
            # Open file dialog
            $openFileDialog = New-Object -TypeName Microsoft.Win32.OpenFileDialog
            $openFileDialog.Title = "Select File to Attach"
            $result = $openFileDialog.ShowDialog()

            if ($result -eq "OK") {
                $global:attachment = $openFileDialog.FileName
                [System.Windows.MessageBox]::Show("Attachment selected: $attachment", "Add Attachment")
                $newDialog.Close()    
            }
        })

        # Add Contact button
        $contactButton = New-Object -TypeName System.Windows.Controls.Button
        $contactButton.Content = "Contact"
        $contactButton.Height = 25
        $contactButton.Width = 70
        $contactButton.Margin = "5"
        $contactButton.Add_Click({
            # Create a dialog window
            $contactDialog = New-Object System.Windows.Window
            $contactDialog.Title = "Add Contacts"
            $contactDialog.Height = 160
            $contactDialog.Width = 320
            $contactDialog.ResizeMode = "NoResize"
            $contactDialog.WindowStartupLocation = "CenterScreen"
    
            # Create a panel to hold the controls
            $contactPanel = New-Object System.Windows.Controls.StackPanel
            $contactPanel.Margin = "5"
    
            # Create a prompt TextBlock
            $prompt = New-Object System.Windows.Controls.TextBlock
            $prompt.TextWrapping = "Wrap"
            $prompt.Text = "Please enter valid Nissan e-mail addresses separated by semicolons (for example: `"Jim.Parsons@Nissan-Usa.com; Sheldon.Cooper@Nissan-Usa.com`"): "
            $prompt.Margin = "5"
    
            # Create an input TextBox
            $contactTextBox = New-Object System.Windows.Controls.TextBox
            $contactTextBox.Width = 280
            $contactTextBox.Margin = "5"
    
            # Create an OK Button
            $contactOKButton = New-Object System.Windows.Controls.Button
            $contactOKButton.Content = "OK"
            $contactOKButton.Height = 25
            $contactOKButton.Width = 65
            $contactOKButton.Margin = "5"
            $contactOKButton.Add_Click({
                # Check the added contact email addresses to confirm validity here.
                $validEmail = "@nissan-usa.com"
                if ($contactTextBox -match "(?i)$validEmail") {
                    # Add the entered email addresses to the $defaults string
                    $contactDialog.DialogResult = $true
                    $global:defaults += $contactTextBox.Text + ";"
                    [System.Windows.MessageBox]::Show("Contacts successfully added: $($contactTextBox.Text)", "Add Contacts")
                    $contactDialog.Close()                
                    $newDialog.Close()
                } else {
                    # Warn user of invalid email.
                    [System.Windows.MessageBox]::Show("One or more e-mail addresses is invalid. Please try again.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            })

            # Create cancel button
            $contactCancelButton = New-Object System.Windows.Controls.Button
            $contactCancelButton.Content = "OK"
            $contactCancelButton.Height = 25
            $contactCancelButton.Width = 75
            $contactCancelButton.Margin = "5"
            $contactCancelButton.Add_Click({
                # Close the form
                $contactDialog.Close()
            })
                
            # Add controls to the panel
            $contactPanel.Children.Add($prompt)
            $contactPanel.Children.Add($contactTextBox)
            $contactPanel.Children.Add($contactOKButton)
            $contactPanel.Children.Add($contactCancelButton)
    
            # Set the dialog content
            $contactDialog.Content = $contactPanel
    
            # Show the dialog
            $contactDialog.ShowDialog() | Out-Null
        })

        # Add New dialog and panel setup
        $newButtonPanel.Children.Add($attachmentButton)
        $newButtonPanel.Children.Add($contactButton)
        $newDialog.Content = $newButtonPanel
        $newDialog.ShowDialog() | Out-Null
    })


    $okButton = New-Object -TypeName System.Windows.Controls.Button
    $okButton.Content = "OK"
    $okButton.Width = 70
    $okButton.Margin = "5"
    $okButton.IsDefault = $true
    $okButton.Add_Click({
        $window.DialogResult = $true
    })

    $cancelButton = New-Object -TypeName System.Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Width = 70
    $cancelButton.Margin = "5"
    $cancelButton.IsCancel = $true

    # Adding main button panel children
    $buttonPanel.Children.Add($newButton)
    $buttonPanel.Children.Add($okButton)
    $buttonPanel.Children.Add($cancelButton)

    # Hashtable logic here. This handles storing the strings entered in the Description fields in a dictionary object.
    if ($window.ShowDialog() -eq $true) {
        foreach ($tabItem in $tabControl.Items) {
            # Access the content of the tab item
            $content = $tabItem.Content

            # Find the text box within the content of each tab item
            $textBox = $content.Children | Where-Object { $_ -is [System.Windows.Controls.TextBox] }

            # Store the description with the tab item header as the key
            $global:descriptionList[$tabItem.Header] = $textBox.Text
        }

    } else {
        $window.Close()
        Exit
    }

    # Pre-call debug
    #foreach ($itemName in $itemNames) {
    #    Write-Host "Pre call test: $($descriptionList[$itemName])"
    #}

}

# Prompt user for description
Get-Descriptions -itemNames $itemNames

# Post-call debug
#foreach ($itemName in $itemNames) {
#    Write-Host "Post call test: $($descriptionList[$itemName])"
#}

# Conditionals to handle OEE and Spot Repair descriptions 
if ($descriptionList['OEE/Downtime'] -eq "") {
    $emailDowntime = " - No items for today."
} else {
    $emailDowntime = $descriptionList['OEE/Downtime']  
}

if ($descriptionList['Spot Repair'] -eq "") {
    $emailSpotRepair = " - No items for today."
} else {
    $emailSpotRepair = $descriptionList['OEE/Downtime']  
} 

# Create HTML body
$emailBody = @"
<style>
table {
    border-collapse: collapse;
    border: 1px solid black;
	width: 520px;
}
td {
    border: 1px solid black;
    padding: 2px;
	height: 72px;
    text-align: left;
	font-family: 'Calibri';
    font-size: 11pt;
}
th {
    border: 1px solid black;
	padding: 2px;
    text-align: left;
	font-family: 'Calibri';
    font-size: 11pt;
	font-weight: bold;		
    background-color: lightgrey;
}
</style>
<table>
<tr><th colspan='2'> Internal KPI Follow-Up Items</th></tr>
<tr><th style='background-color:#FFF'> OEE/Equipment/Downtime</th><th style='background-color:#FFF'>Spot Repair/Repaint</th></tr>
<tr><td>$($emailDowntime)</td><td>$($emailSpotRepair)</td></tr>
"@

$remainingItems = @("PDI", "VES", "Unit Review", "New Material/New Equipment", "Special Request")

foreach ($itemName in $remainingItems) {
    if ($descriptionList[$itemName] -eq "") {
        $emailBody += "<tr><th colspan='2'> $itemName Items</th></tr>"
        $emailBody += "<tr><td colspan='2'> - No items for today.</td></tr>"
    } else {
        $emailBody += "<tr><th colspan='2'> $itemName Items</th></tr>"
        $emailBody += "<tr><td colspan='2'>$($descriptionList[$itemName])</td></tr>"
    }
}

$emailBody += "</table>"

# Split the defaults string into an array of email addresses and display names
$contacts = $defaults -split '; '

# Create Outlook application object
$outlook = New-Object -ComObject Outlook.Application

# Create a new email
$mail = $outlook.CreateItem(0)

# Set email properties
$mail.Display()
$mail.To = $contacts
$mail.Subject = "PL1 Topcoat Engineering EOS for $today"
$mail.HTMLBody = $emailBody + "<br>" + $mail.HTMLBody 

if ($null -ne $attachment) {
    $attachmentObject = $mail.Attachments.Add($attachment)
}
