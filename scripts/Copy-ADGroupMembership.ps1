#region Assemblies
Add-Type -AssemblyName PresentationFramework
#endregion

#region Global Variables
$Global:adUserTarget = $null
#endregion

#region Function - Get AD Group Memberships for a User
function Get-UserGroups {
    param (
        [string]$Username
    )
    # Ensure the Active Directory module is available
    Import-Module ActiveDirectory -ErrorAction Stop

    try {
        # Get user groups
        $groups = Get-ADUser -Identity $Username -Properties MemberOf | Select-Object -ExpandProperty MemberOf
        $groups | ForEach-Object {
            ($_ -split ',')[0] -replace 'CN='
        }
    } catch {
        "User not found or error: $_"
    }
}

#endregion

#region Window Setup
# Create WPF Window
$window = New-Object System.Windows.Window
$window.Title = "AD Group Membership Copy"
$window.Width = 800
$window.Height = 600
#endregion

#region Layout - Grid Setup
$grid = New-Object System.Windows.Controls.Grid

# Column Definitions
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "3*" })) # Left
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "2*" })) # Middle
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "3*" })) # Right

# Row Definitions
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) # Row 0: Inputs
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "*" }))    # Row 1: ListBoxes
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) # Row 2: Arrows
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) # Row 3: Actions
#endregion

#region UI Elements - Source User (Left Pane)
$textBox1 = New-Object System.Windows.Controls.TextBox
$textBox1.Margin = "10"
$textBox1.Height = "20"
$textBox1.TextAlignment = 'Center'
$textBox1.Text = "source user"
$textBox1.Foreground = [System.Windows.Media.Brushes]::Gray
[System.Windows.Controls.Grid]::SetRow($textBox1, 0)
[System.Windows.Controls.Grid]::SetColumn($textBox1, 0)
$grid.Children.Add($textBox1)

$button1 = New-Object System.Windows.Controls.Button
$button1.Content = "Fetch Group(s)"
$button1.Margin = "10"
$button1.Height = "20"
$button1.HorizontalAlignment = "Right"
[System.Windows.Controls.Grid]::SetRow($button1, 0)
[System.Windows.Controls.Grid]::SetColumn($button1, 0)
$grid.Children.Add($button1)

$buttonClearSource = New-Object System.Windows.Controls.Button
$buttonClearSource.Content = [char]0x2716
$buttonClearSource.FontFamily = "Segoe UI Symbol"
$buttonClearSource.Height = 20
$buttonClearSource.Margin = '10'
$buttonClearSource.HorizontalAlignment = "Left"
$buttonClearSource.ToolTip = "Clear source input"
[System.Windows.Controls.Grid]::SetRow($buttonClearSource, 0)
[System.Windows.Controls.Grid]::SetColumn($buttonClearSource, 0)
$grid.Children.Add($buttonClearSource)
#endregion

#region UI Elements - Target User (Right Pane)
$textBox2 = New-Object System.Windows.Controls.TextBox
$textBox2.Margin = "10"
$textBox2.Height = "20"
$textBox2.TextAlignment = 'Center'
$textBox2.Text = "target user"
$textBox2.Foreground = [System.Windows.Media.Brushes]::Gray
[System.Windows.Controls.Grid]::SetRow($textBox2, 0)
[System.Windows.Controls.Grid]::SetColumn($textBox2, 2)
$grid.Children.Add($textBox2)

$button2 = New-Object System.Windows.Controls.Button
$button2.Content = "Fetch Group(s)"
$button2.Margin = "10"
$button2.Height = "20"
$button2.HorizontalAlignment = "Right"
[System.Windows.Controls.Grid]::SetRow($button2, 0)
[System.Windows.Controls.Grid]::SetColumn($button2, 2)
$grid.Children.Add($button2)

$buttonClearTarget = New-Object System.Windows.Controls.Button
$buttonClearTarget.Content = [char]0x2716
$buttonClearTarget.FontFamily = "Segoe UI Symbol"
$buttonClearTarget.Height = 20
$buttonClearTarget.Margin = '10'
$buttonClearTarget.HorizontalAlignment = "Left"
$buttonClearTarget.ToolTip = "Clear source input"
[System.Windows.Controls.Grid]::SetRow($buttonClearTarget, 0)
[System.Windows.Controls.Grid]::SetColumn($buttonClearTarget, 2)
$grid.Children.Add($buttonClearTarget)
#endregion

#region UI Elements - Middle Buttons (Top + Confirm + Clear)
$buttonClear1 = New-Object System.Windows.Controls.Button
$buttonClear1.Content = "Clear"
$buttonClear1.Margin = "0, 0, 0, 10"
$buttonClear1.Width = 50
$buttonClear1.Height = 20
$buttonClear1.HorizontalAlignment = "Left"
[System.Windows.Controls.Grid]::SetRow($buttonClear1, 2)
[System.Windows.Controls.Grid]::SetColumn($buttonClear1, 1)
$grid.Children.Add($buttonClear1)

$buttonConfirm1 = New-Object System.Windows.Controls.Button
$buttonConfirm1.Content = "Add * To Group(s)"
$buttonConfirm1.Margin = "0, 0, 0, 10"
$buttonConfirm1.Width = "120"
$buttonConfirm1.HorizontalAlignment = "Right"
$buttonConfirm1.ToolTip = "Add target user to selected groups"
[System.Windows.Controls.Grid]::SetRow($buttonConfirm1, 2)
[System.Windows.Controls.Grid]::SetColumn($buttonConfirm1, 1)
$grid.Children.Add($buttonConfirm1)
#endregion

#region UI Elements - Arrow Button
$buttonArrow1 = New-Object System.Windows.Controls.Button
$buttonArrow1.Content = [char]0x2192
$buttonArrow1.Margin = "0, 0, 0, 10"
$buttonArrow1.Width = "50"
$buttonArrow1.HorizontalAlignment = "Center"
$buttonArrow1.ToolTip = "Move selected groups to transfer list"
[System.Windows.Controls.Grid]::SetRow($buttonArrow1, 2)
[System.Windows.Controls.Grid]::SetColumn($buttonArrow1, 0)
$grid.Children.Add($buttonArrow1)
#endregion

#region UI Elements - ListBoxes
$listBox1 = New-Object System.Windows.Controls.ListBox
$listBox1.Margin = "10, 10, 10, 0"
$listBox1.SelectionMode = 'Extended'
[System.Windows.Controls.Grid]::SetRow($listBox1, 1)
[System.Windows.Controls.Grid]::SetColumn($listBox1, 0)
$grid.Children.Add($listBox1)

$listBox2 = New-Object System.Windows.Controls.ListBox
$listBox2.Margin = "10, 10, 10, 0"
$listBox2.SelectionMode = 'Extended'
[System.Windows.Controls.Grid]::SetRow($listBox2, 1)
[System.Windows.Controls.Grid]::SetColumn($listBox2, 2)
$grid.Children.Add($listBox2)

$listBoxMiddle = New-Object System.Windows.Controls.ListBox
$listBoxMiddle.Margin = "0, 10, 0, 0"
$listBoxMiddle.SelectionMode = 'Extended'
[System.Windows.Controls.Grid]::SetRow($listBoxMiddle, 1)
[System.Windows.Controls.Grid]::SetColumn($listBoxMiddle, 1)
$grid.Children.Add($listBoxMiddle)
#endregion

#region events
# Add Event Handlers for Buttons
$button1.Add_Click({
    $username = $textBox1.Text.Trim()
    if ($username) {
        $groups = Get-UserGroups -Username $username
        $listBox1.Items.Clear()
        $groups | ForEach-Object { $listBox1.Items.Add($_) }
    } else {
        [System.Windows.MessageBox]::Show("Please enter a source username (Initials).", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$buttonClearSource.Add_Click({
    $listBox1.Items.Clear()
    $textBox1.Text = "source user"
    $textBox1.Foreground = [System.Windows.Media.Brushes]::Gray
})

$buttonClearTarget.Add_Click({
    $listBox2.Items.Clear()
    $textBox2.Text = "target user"
    $textBox2.Foreground = [System.Windows.Media.Brushes]::Gray
    $buttonConfirm1.Content = "Add * To Group(s)"
})

$buttonArrow1.Add_Click({
    $listBoxMiddle.Items.Clear()
    if (($listBox1.SelectedItems).count -gt 0) {
        $listBox1.SelectedItems | ForEach-Object { $listBoxMiddle.Items.Add($_) }
        $listBox1.UnselectAll()
    } else {
        [System.Windows.MessageBox]::Show("Please select at least one item.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$buttonClear1.Add_Click({
    $listBoxMiddle.Items.Clear()
    $listBox1.UnselectAll()
    $listBox2.UnselectAll()
})

$buttonConfirm1.Add_Click({
    if ($listBoxMiddle.items.count -gt 0 -and $Global:adUserTarget) {
        $listBoxMiddle.items | ForEach-Object {
            $group = New-Object System.Windows.Controls.ListBoxItem
            $group.Foreground = [System.Windows.Media.Brushes]::Green
            $group.Content = $_.ToString()
            $listBox2.Items.Add($group)
            try {
                Add-ADGroupMember -Identity $_ -Members $Global:adUserTarget -ErrorAction Stop
            } catch {
                Write-Error "Failed to add $Global:adUserTarget to group $_"
            }
        }
    } elseif ($listBoxMiddle.items.count -le 0) {
        [System.Windows.MessageBox]::Show("No groups have been selected.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } elseif (-not $Global:adUserTarget) {
        [System.Windows.MessageBox]::Show("The target user is not valid.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }

})

$button2.Add_Click({
    $username = $textBox2.Text.Trim()
    if ($username) {
        $buttonConfirm1.Content = "Add $($textBox2.Text) To Group(s)"
        $Global:adUserTarget = Get-ADUser $username
        $groups = Get-UserGroups -Username $username
        $listBox2.Items.Clear()
        $groups | ForEach-Object { $listBox2.Items.Add($_) }
    } else {
        [System.Windows.MessageBox]::Show("Please enter a target username (Initials).", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Add Event Handler to Remove Placeholder Text on Focus (Click)
$textBox1.Add_GotFocus({
    if ($textBox1.Text -eq "source user") {
        $textBox1.Text = ""
        $textBox1.Foreground = [System.Windows.Media.Brushes]::Black  # Change text color to black when user types
    }
})

# Add Event Handler to Restore Placeholder Text if TextBox is Empty on Lost Focus
$textBox1.Add_LostFocus({
    if ($textBox1.Text -eq "") {
        $textBox1.Text = "source user"
        $textBox1.Foreground = [System.Windows.Media.Brushes]::Gray  # Change text color back to grey
    }
})

# Add Event Handler to Remove Placeholder Text on Focus (Click)
$textBox2.Add_GotFocus({
    if ($textBox2.Text -eq "target user") {
        $textBox2.Text = ""
        $textBox2.Foreground = [System.Windows.Media.Brushes]::Black  # Change text color to black when user types
    }
})

# Add Event Handler to Restore Placeholder Text if TextBox is Empty on Lost Focus
$textBox2.Add_LostFocus({
    if ($textBox2.Text -eq "") {
        $textBox2.Text = "target user"
        $textBox2.Foreground = [System.Windows.Media.Brushes]::Gray  # Change text color back to grey
    }
})

#endregion

#region main script flow

# Show Window
$window.Content = $grid
$window.ShowDialog()

Exit

#endregion
