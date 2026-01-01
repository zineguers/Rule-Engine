Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

function Pump-UI { [System.Windows.Forms.Application]::DoEvents() }

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="[Rule Engine]" Height="450" Width="550" 
        AllowsTransparency="True" WindowStyle="None" Background="Transparent" WindowStartupLocation="CenterScreen">
    <Border x:Name="MainBorder" CornerRadius="20" Background="#1C1C1E" BorderBrush="#3A3A3C" BorderThickness="1" Opacity="0" RenderTransformOrigin="0.5,0.5">
        <Border.Effect>
            <DropShadowEffect x:Name="WindowShadow" BlurRadius="15" Color="#000000" Opacity="0" ShadowDepth="5"/>
        </Border.Effect>
        <Border.RenderTransform>
            <TransformGroup>
                <ScaleTransform ScaleX="0.95" ScaleY="0.95"/>
                <TranslateTransform Y="25"/>
            </TransformGroup>
        </Border.RenderTransform>
        <Grid x:Name="ContentGrid">
            <Grid.Effect>
                <BlurEffect x:Name="ContentBlur" Radius="10"/>
            </Grid.Effect>
            <Grid VerticalAlignment="Top" Margin="30,25,30,0">
                <StackPanel>
                    <TextBlock Text="[Rule Engine]" FontSize="26" FontWeight="Bold" Foreground="#FFFFFF"/>
                    <TextBlock Name="StatusTxt" Text="System Ready" FontSize="11" Foreground="#AEAEB2" Margin="2,0,0,0"/>
                </StackPanel>
                <Button Name="BtnClose" Content="Exit" HorizontalAlignment="Right" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Foreground="#AEAEB2" Cursor="Hand" FontSize="12">
                    <Button.Style>
                        <Style TargetType="Button">
                            <Setter Property="Foreground" Value="#FF0000"/>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Foreground" Value="#FFFFFF"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </Button.Style>
                </Button>
            </Grid>

            <StackPanel Margin="30,85,30,0">
                <TextBlock Text="TENANT ID" FontSize="9" FontWeight="Bold" Foreground="#AEAEB2" Margin="5,0,0,3"/>
                <TextBox Name="TxtTenant" Height="28" Background="#2C2C2E" BorderBrush="#3A3A3C" VerticalContentAlignment="Center" Margin="0,0,0,6" Foreground="#FFFFFF"/>

                <TextBlock Text="CLIENT ID" FontSize="9" FontWeight="Bold" Foreground="#AEAEB2" Margin="5,0,0,3"/>
                <TextBox Name="TxtClient" Height="28" Background="#2C2C2E" BorderBrush="#3A3A3C" VerticalContentAlignment="Center" Margin="0,0,0,6" Foreground="#FFFFFF"/>

                <TextBlock Text="CLIENT SECRET" FontSize="9" FontWeight="Bold" Foreground="#AEAEB2" Margin="5,0,0,3"/>
                <PasswordBox Name="TxtSecret" Height="28" Background="#2C2C2E" BorderBrush="#3A3A3C" VerticalContentAlignment="Center" Margin="0,0,0,8" Foreground="#FFFFFF"/>

                <Grid Margin="0,0,0,10">
                    <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                    <Button Name="BtnClear" Grid.Column="0" Content="Clear" Height="22" BorderThickness="0" Margin="0,0,4,0" Cursor="Hand" Foreground="#FFFFFF" FontSize="11">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Background" Value="#3A3A3C"/>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#333333"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>
                    <Button Name="BtnTest" Grid.Column="1" Content="Verify" Height="22" BorderThickness="0" Margin="4,0,0,0" Cursor="Hand" Foreground="White" FontSize="11">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Background" Value="#5e5cff"/>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#4c49ff"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>
                </Grid>

                <Rectangle Height="1" Fill="#3A3A3C" Margin="0,5,0,10"/>

                <TextBlock Text="TARGET MAILBOXES" FontSize="9" FontWeight="Bold" Foreground="#AEAEB2" Margin="5,0,0,3"/>
                <Grid Margin="0,0,0,8">
                    <Grid.ColumnDefinitions><ColumnDefinition Width="3*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <TextBox Name="TxtEmail" Grid.Column="0" Height="28" Background="#2C2C2E" BorderBrush="#3A3A3C" VerticalContentAlignment="Center" Margin="0,0,4,0" Foreground="#FFFFFF"/>
                    <Button Name="BtnLoad" Grid.Column="1" Content="Import .txt" Height="28" BorderThickness="0" Margin="4,0,0,0" Cursor="Hand" Foreground="White" FontSize="11">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Background" Value="#5e5cff"/>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#4c49ff"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>
                </Grid>

                <ProgressBar Name="ProgBar" Height="3" Foreground="#0A84FF" Background="#3A3A3C" BorderThickness="0" Margin="0,10,0,10"/>
            </StackPanel>

            <Grid VerticalAlignment="Bottom" Margin="30,0,30,30">
                <Button Name="BtnCheck" Height="40" Cursor="Hand">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="border" Background="#0A84FF" CornerRadius="10">
                                <TextBlock Text="SEARCH NOW" Foreground="White" FontWeight="Bold" VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="14"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="border" Property="Background" Value="#4c49ff"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </Grid>
        </Grid>
    </Border>
    <Window.Triggers>
        <EventTrigger RoutedEvent="Loaded">
            <BeginStoryboard>
                <Storyboard>
                    <DoubleAnimation Storyboard.TargetName="MainBorder" Storyboard.TargetProperty="Opacity" To="1" Duration="0:0:0.75">
                        <DoubleAnimation.EasingFunction>
                            <CubicEase EasingMode="EaseOut"/>
                        </DoubleAnimation.EasingFunction>
                    </DoubleAnimation>
                    <DoubleAnimation Storyboard.TargetName="MainBorder" Storyboard.TargetProperty="(UIElement.RenderTransform).(TransformGroup.Children)[0].(ScaleTransform.ScaleX)" To="1" Duration="0:0:1.1">
                        <DoubleAnimation.EasingFunction>
                            <BackEase EasingMode="EaseOut" Amplitude="0.18"/>
                        </DoubleAnimation.EasingFunction>
                    </DoubleAnimation>
                    <DoubleAnimation Storyboard.TargetName="MainBorder" Storyboard.TargetProperty="(UIElement.RenderTransform).(TransformGroup.Children)[0].(ScaleTransform.ScaleY)" To="1" Duration="0:0:1.1">
                        <DoubleAnimation.EasingFunction>
                            <BackEase EasingMode="EaseOut" Amplitude="0.18"/>
                        </DoubleAnimation.EasingFunction>
                    </DoubleAnimation>
                    <DoubleAnimation Storyboard.TargetName="MainBorder" Storyboard.TargetProperty="(UIElement.RenderTransform).(TransformGroup.Children)[1].(TranslateTransform.Y)" To="0" Duration="0:0:0.95">
                        <DoubleAnimation.EasingFunction>
                            <QuinticEase EasingMode="EaseOut"/>
                        </DoubleAnimation.EasingFunction>
                    </DoubleAnimation>
                    <DoubleAnimation Storyboard.TargetName="ContentBlur" Storyboard.TargetProperty="Radius" To="0" Duration="0:0:0.65">
                        <DoubleAnimation.EasingFunction>
                            <ExponentialEase EasingMode="EaseOut"/>
                        </DoubleAnimation.EasingFunction>
                    </DoubleAnimation>
                    <DoubleAnimation Storyboard.TargetName="WindowShadow" Storyboard.TargetProperty="Opacity" To="0.1" Duration="0:0:1.1"/>
                </Storyboard>
            </BeginStoryboard>
        </EventTrigger>
    </Window.Triggers>
</Window>
"@

try {
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    [System.Windows.Forms.MessageBox]::Show("XAML Load Error: $($_.Exception.Message)")
    exit
}

$status = $window.FindName("StatusTxt")
$prog   = $window.FindName("ProgBar")
$txtTenant = $window.FindName("TxtTenant")
$txtClient = $window.FindName("TxtClient")
$txtSecret = $window.FindName("TxtSecret")
$txtEmail  = $window.FindName("TxtEmail")
$btnCheck  = $window.FindName("BtnCheck")
$btnLoad   = $window.FindName("BtnLoad")
$btnClear  = $window.FindName("BtnClear")
$btnTest   = $window.FindName("BtnTest")

function Get-GraphToken($t, $c, $s) {
    $body = @{ grant_type="client_credentials"; scope="https://graph.microsoft.com/.default"; client_id=$c; client_secret=$s }
    return (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$t/oauth2/v2.0/token" -Method Post -Body $body).access_token
}

$window.FindName("BtnClose").Add_Click({ $window.Close() })
$window.Add_MouseLeftButtonDown({ $window.DragMove() })

$btnClear.Add_Click({
    $txtTenant.Text = ""
    $txtClient.Text = ""
    $txtSecret.Password = ""
    $txtEmail.Text = ""
    $global:emailList = @()
    $status.Text = "Form reset."
})

$btnTest.Add_Click({
    $tid = $txtTenant.Text.Trim()
    $cid = $txtClient.Text.Trim()
    $sec = $txtSecret.Password.Trim()
    if (!$tid -or !$cid -or !$sec) { [System.Windows.Forms.MessageBox]::Show("Credentials required."); return }
    
    $status.Text = "Testing Connection..."
    Pump-UI
    try {
        $tk = Get-GraphToken $tid $cid $sec
        if ($tk) {
            $status.Text = "Connection Successful!"
            [System.Windows.Forms.MessageBox]::Show("Connection Successful!", "Success")
        }
    } catch {
        $status.Text = "Connection Failed."
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failed")
    }
})

$btnLoad.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Text Files (*.txt)|*.txt"
    if ($ofd.ShowDialog() -eq "OK") {
        $global:emailList = Get-Content $ofd.FileName | ForEach-Object { $_.Split(',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique
        
        $txtEmail.Text = $global:emailList -join ", "
        $status.Text = "Loaded $($global:emailList.Count) emails from file."
    }
})

$global:emailList = @()
$global:rawResults = {}

$btnCheck.Add_Click({
    Show-Results
})

function Show-Results {
    $tid = $txtTenant.Text.Trim()
    $cid = $txtClient.Text.Trim()
    $sec = $txtSecret.Password.Trim()
    if (!$tid -or !$cid -or !$sec) { [System.Windows.Forms.MessageBox]::Show("Credentials required."); return }
    
    $status.Text = "Authenticating..."; Pump-UI

    try {
        $tk = Get-GraphToken $tid $cid $sec

        $emails = if ($global:emailList.Count -gt 0) { $global:emailList } else { $txtEmail.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } }

        if ($emails.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No mailboxes specified.")
            $status.Text = "Ready"; $prog.Value = 0
            return
        }

        $resultsList = New-Object System.Collections.ArrayList
        $global:rawResults = @{}

        $c = 0
        foreach ($e in $emails) {
            $c++
            if ($emails.Count -gt 0) { $prog.Value = ($c / $emails.Count) * 100 }
            $status.Text = "Scanning: $e"; Pump-UI

            $hasRules = $false
            $mailboxAccessed = $false

            $headers = @{ Authorization = "Bearer $tk" }
            try {
                $st = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$e/mailboxSettings" -Headers $headers
                $mailboxAccessed = $true
                if ($st.forwardingSmtpAddress) {
                    $uniqueKey = [guid]::NewGuid().ToString()
                    $resultsList.Add([PSCustomObject]@{ Mailbox=$e; Type="Forwarding"; RuleName="SMTP Setting"; Condition="Forwarding"; From="N/A"; To=$st.forwardingSmtpAddress; Key=$uniqueKey }) | Out-Null
                    $global:rawResults[$uniqueKey] = @{ ID="SMTP"; Raw=$st }
                    $hasRules = $true
                }
            } catch {}
            try {
                $rules = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$e/mailFolders/inbox/messageRules" -Headers $headers
                $mailboxAccessed = $true
                foreach ($rule in $rules.value) {
                    $foundCondition = "Other"; $fromVal = "Any"; $toVal = "None"
                    if ($rule.actions.redirectTo) { $foundCondition = "Redirect"; $toVal = $rule.actions.redirectTo.emailAddress.address -join "; " }
                    elseif ($rule.actions.forwardTo) { $foundCondition = "Forward"; $toVal = $rule.actions.forwardTo.emailAddress.address -join "; " }
                    elseif ($rule.actions.moveToFolder) { $foundCondition = "Move to Folder"; $toVal = "Folder: " + $rule.actions.moveToFolder }
                    elseif ($rule.actions.delete -eq $true) { $foundCondition = "Delete"; $toVal = "Deleted Items" }
                    if ($rule.conditions.fromAddresses) { $fromVal = $rule.conditions.fromAddresses.emailAddress.address -join "; " }
                    $uniqueKey = [guid]::NewGuid().ToString()
                    $resultsList.Add([PSCustomObject]@{ Mailbox=$e; Type="Inbox Rule"; RuleName=$rule.displayName; Condition=$foundCondition; From=$fromVal; To=$toVal; Key=$uniqueKey }) | Out-Null
                    $global:rawResults[$uniqueKey] = @{ ID=$rule.id; Raw=$rule }
                    $hasRules = $true
                }
            } catch {}

            if ($mailboxAccessed -and -not $hasRules) {
                $resultsList.Add([PSCustomObject]@{ Mailbox=$e; Type="No Rules Found"; RuleName="No rules or forwarding configured"; Condition="-"; From="-"; To="-"; Key="dummy" }) | Out-Null
            }

            if (-not $mailboxAccessed) {
                $resultsList.Add([PSCustomObject]@{ Mailbox=$e; Type="Access Error"; RuleName="Unable to access mailbox"; Condition="Insufficient permissions or mailbox does not exist"; From="-"; To="-"; Key="error" }) | Out-Null
            }
        }

        $sortedResults = @($resultsList | Sort-Object Mailbox)
        $realCount = ($sortedResults | Where-Object { $_.Type -notin @("No Rules Found", "Access Error") }).Count

        Show-ResultsWindow -Results $sortedResults -AccessToken $tk -Mailboxes $emails -RealCount $realCount

    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message)
    }
    $status.Text = "Ready"; $prog.Value = 0
}

function Get-SelectedMailboxes {
    param([string[]]$Mailboxes)

    if ($Mailboxes.Count -le 1) {
        return $Mailboxes
    }

    $selXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Mailboxes for Rule Title Rules" Height="500" Width="600" 
        AllowsTransparency="True" WindowStyle="None" Background="Transparent" WindowStartupLocation="CenterScreen">
    <Border CornerRadius="20" Background="#1C1C1E" BorderBrush="#3A3A3C" BorderThickness="1">
        <Border.Effect><DropShadowEffect BlurRadius="15" Color="#000000" Opacity="0.1" ShadowDepth="5"/></Border.Effect>
        <Grid>
            <Grid VerticalAlignment="Top" Margin="30,25,30,0">
                <StackPanel>
                    <TextBlock Text="Select Mailboxes" FontSize="26" FontWeight="Bold" Foreground="#FFFFFF"/>
                    <TextBlock Text="Hold Ctrl/Shift for multiple selection" FontSize="11" Foreground="#AEAEB2" Margin="2,10,0,0"/>
                </StackPanel>
                <Button Name="SelClose" Content="X" HorizontalAlignment="Right" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Foreground="#AEAEB2" Cursor="Hand" FontSize="12">
                    <Button.Style>
                        <Style TargetType="Button">
                            <Setter Property="Foreground" Value="#AEAEB2"/>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Foreground" Value="#FF0000"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </Button.Style>
                </Button>
            </Grid>

            <ListBox Name="LstMailboxes" Margin="30,100,30,100" Background="#2C2C2E" Foreground="#FFFFFF" BorderBrush="#3A3A3C" SelectionMode="Extended"/>

            <Grid VerticalAlignment="Bottom" Margin="30,0,30,30">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Name="BtnSelCancel" Grid.Column="0" Content="Cancel" Height="40" Margin="0,0,10,0" Background="#3A3A3C" Foreground="#FFFFFF" Cursor="Hand"/>
                <Button Name="BtnSelOK" Grid.Column="1" Content="Apply to Selected" Height="40" Margin="10,0,0,0" Background="#5e5cff" Foreground="#FFFFFF" Cursor="Hand"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

    $selReader = New-Object System.Xml.XmlNodeReader ([xml]$selXaml)
    $selWindow = [Windows.Markup.XamlReader]::Load($selReader)

    $lst = $selWindow.FindName("LstMailboxes")
    $lst.ItemsSource = $Mailboxes | Sort-Object

    $btnOK = $selWindow.FindName("BtnSelOK")
    $btnCancel = $selWindow.FindName("BtnSelCancel")
    $selClose = $selWindow.FindName("SelClose")

    $script:selectedMailboxes = $null

    $btnOK.Add_Click({
        if ($lst.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one mailbox.")
            return
        }
        $script:selectedMailboxes = @($lst.SelectedItems)
        $selWindow.Close()
    })

    $btnCancel.Add_Click({ $selWindow.Close() })
    $selClose.Add_Click({ $selWindow.Close() })
    $selWindow.Add_MouseLeftButtonDown({ $selWindow.DragMove() })

    $selWindow.ShowDialog() | Out-Null

    return $script:selectedMailboxes
}

function Show-ResultsWindow {
    param($Results, $AccessToken, $Mailboxes, $RealCount)

    $resXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Rule Scan Results" Height="600" Width="1100" 
        AllowsTransparency="True" WindowStyle="None" Background="Transparent" WindowStartupLocation="CenterScreen">
    <Border CornerRadius="20" Background="#1C1C1E" BorderBrush="#3A3A3C" BorderThickness="1">
        <Border.Effect><DropShadowEffect BlurRadius="15" Color="#000000" Opacity="0.1" ShadowDepth="5"/></Border.Effect>
        <Grid>
            <Grid VerticalAlignment="Top" Margin="30,25,30,0">
                <StackPanel>
                    <TextBlock Text="Rule Scan Results" FontSize="26" FontWeight="Bold" Foreground="#FFFFFF"/>
                    <TextBlock Name="ResStatus" Text="" FontSize="11" Foreground="#AEAEB2" Margin="2,0,0,0"/>
                </StackPanel>
                <Button Name="ResClose" Content="Exit" HorizontalAlignment="Right" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Foreground="#AEAEB2" Cursor="Hand" FontSize="12">
                    <Button.Style>
                        <Style TargetType="Button">
                            <Setter Property="Foreground" Value="#FF0000"/>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Foreground" Value="#FFFFFF"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </Button.Style>
                </Button>
            </Grid>

            <DataGrid Name="DgvResults" Margin="30,85,30,100" AutoGenerateColumns="False" Background="#2C2C2E" Foreground="#FFFFFF" BorderBrush="#3A3A3C" GridLinesVisibility="None" HeadersVisibility="Column" RowBackground="#2C2C2E" AlternatingRowBackground="#1C1C1E" CanUserAddRows="False" IsReadOnly="True" SelectionMode="Extended">
                <DataGrid.Resources>
                    <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightBrushKey}" Color="#5e5cff"/>
                    <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightTextBrushKey}" Color="#FFFFFF"/>
                </DataGrid.Resources>
                <DataGrid.ColumnHeaderStyle>
                    <Style TargetType="DataGridColumnHeader">
                        <Setter Property="Background" Value="#3A3A3C"/>
                        <Setter Property="Foreground" Value="#FFFFFF"/>
                        <Setter Property="FontWeight" Value="Bold"/>
                        <Setter Property="Padding" Value="10,5"/>
                    </Style>
                </DataGrid.ColumnHeaderStyle>
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Mailbox" Binding="{Binding Mailbox}" Width="*"/>
                    <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="Auto"/>
                    <DataGridTextColumn Header="RuleName" Binding="{Binding RuleName}" Width="Auto"/>
                    <DataGridTextColumn Header="Condition" Binding="{Binding Condition}" Width="Auto"/>
                    <DataGridTextColumn Header="From" Binding="{Binding From}" Width="Auto"/>
                    <DataGridTextColumn Header="To" Binding="{Binding To}" Width="*"/>
                </DataGrid.Columns>
                <DataGrid.GroupStyle>
                    <GroupStyle>
                        <GroupStyle.ContainerStyle>
                            <Style TargetType="GroupItem">
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="GroupItem">
                                            <Expander IsExpanded="True" Background="#3A3A3C" Margin="0,5,0,5">
                                                <Expander.Header>
                                                    <StackPanel Orientation="Horizontal">
                                                        <TextBlock Text="{Binding Name, StringFormat='Email account: {0}'}" FontWeight="Bold" Foreground="#FFFFFF" Margin="10,5"/>
                                                        <TextBlock Text="{Binding ItemCount, StringFormat=' ({0} items)'}" Foreground="#AEAEB2" Margin="0,5"/>
                                                    </StackPanel>
                                                </Expander.Header>
                                                <ItemsPresenter Margin="20,0,0,0"/>
                                            </Expander>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </GroupStyle.ContainerStyle>
                    </GroupStyle>
                </DataGrid.GroupStyle>
            </DataGrid>

            <Grid VerticalAlignment="Bottom" Margin="30,0,30,30" Height="40">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition/>
                    <ColumnDefinition/>
                    <ColumnDefinition/>
                    <ColumnDefinition/>
                </Grid.ColumnDefinitions>
                <Button Name="BtnRefresh" Grid.Column="0" Content="REFRESH" Margin="0,0,5,0" Cursor="Hand" Foreground="#FFFFFF" FontSize="11">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Grid>
                                <Border x:Name="border" Background="#3A3A3C" CornerRadius="8"/>
                                <Border x:Name="glow" Opacity="0" CornerRadius="8">
                                    <Border.Effect>
                                        <DropShadowEffect x:Name="glowEffect" Color="#5e5cff" BlurRadius="0" ShadowDepth="0" Opacity="0"/>
                                    </Border.Effect>
                                </Border>
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Grid>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Trigger.EnterActions>
                                        <BeginStoryboard>
                                            <Storyboard>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="BlurRadius" To="25" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="Opacity" To="0.8" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glow" Storyboard.TargetProperty="Opacity" To="1" Duration="0:0:0.3"/>
                                                <ColorAnimation Storyboard.TargetName="border" Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)" To="#5e5cff" Duration="0:0:0.3"/>
                                            </Storyboard>
                                        </BeginStoryboard>
                                    </Trigger.EnterActions>
                                    <Trigger.ExitActions>
                                        <BeginStoryboard>
                                            <Storyboard>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="BlurRadius" To="0" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glow" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.3"/>
                                                <ColorAnimation Storyboard.TargetName="border" Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)" To="#3A3A3C" Duration="0:0:0.3"/>
                                            </Storyboard>
                                        </BeginStoryboard>
                                    </Trigger.ExitActions>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
                <Button Name="BtnPreview" Grid.Column="1" Content="PREVIEW" Margin="5,0,5,0" Cursor="Hand" Foreground="#FFFFFF" FontSize="11">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Grid>
                                <Border x:Name="border" Background="#3A3A3C" CornerRadius="8"/>
                                <Border x:Name="glow" Opacity="0" CornerRadius="8">
                                    <Border.Effect>
                                        <DropShadowEffect x:Name="glowEffect" Color="#5e5cff" BlurRadius="0" ShadowDepth="0" Opacity="0"/>
                                    </Border.Effect>
                                </Border>
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Grid>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Trigger.EnterActions>
                                        <BeginStoryboard>
                                            <Storyboard>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="BlurRadius" To="25" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="Opacity" To="0.8" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glow" Storyboard.TargetProperty="Opacity" To="1" Duration="0:0:0.3"/>
                                                <ColorAnimation Storyboard.TargetName="border" Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)" To="#5e5cff" Duration="0:0:0.3"/>
                                            </Storyboard>
                                        </BeginStoryboard>
                                    </Trigger.EnterActions>
                                    <Trigger.ExitActions>
                                        <BeginStoryboard>
                                            <Storyboard>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="BlurRadius" To="0" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glow" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.3"/>
                                                <ColorAnimation Storyboard.TargetName="border" Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)" To="#3A3A3C" Duration="0:0:0.3"/>
                                            </Storyboard>
                                        </BeginStoryboard>
                                    </Trigger.ExitActions>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
                <Button Name="BtnDelete" Grid.Column="3" Content="DELETE RULE" Margin="5,0,5,0" Cursor="Hand" Foreground="White" FontSize="11">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Grid>
                                <Border x:Name="border" Background="#FF6961" CornerRadius="8"/>
                                <Border x:Name="glow" Opacity="0" CornerRadius="8">
                                    <Border.Effect>
                                        <DropShadowEffect x:Name="glowEffect" Color="#FF0000" BlurRadius="0" ShadowDepth="0" Opacity="0"/>
                                    </Border.Effect>
                                </Border>
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Grid>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Trigger.EnterActions>
                                        <BeginStoryboard>
                                            <Storyboard>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="BlurRadius" To="30" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="Opacity" To="0.9" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glow" Storyboard.TargetProperty="Opacity" To="1" Duration="0:0:0.3"/>
                                                <ColorAnimation Storyboard.TargetName="border" Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)" To="#FF0000" Duration="0:0:0.3"/>
                                            </Storyboard>
                                        </BeginStoryboard>
                                    </Trigger.EnterActions>
                                    <Trigger.ExitActions>
                                        <BeginStoryboard>
                                            <Storyboard>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="BlurRadius" To="0" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glow" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.3"/>
                                                <ColorAnimation Storyboard.TargetName="border" Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)" To="#FF6961" Duration="0:0:0.3"/>
                                            </Storyboard>
                                        </BeginStoryboard>
                                    </Trigger.ExitActions>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
                <Button Name="BtnAddRule Title" Grid.Column="4" Content="Run Engine" Margin="5,0,0,0" Cursor="Hand" Foreground="White" FontSize="11">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Grid>
                                <Border x:Name="border" Background="#5e5cff" CornerRadius="8"/>
                                <Border x:Name="glow" Opacity="0" CornerRadius="8">
                                    <Border.Effect>
                                        <DropShadowEffect x:Name="glowEffect" Color="#7068ff" BlurRadius="0" ShadowDepth="0" Opacity="0"/>
                                    </Border.Effect>
                                </Border>
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Grid>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Trigger.EnterActions>
                                        <BeginStoryboard>
                                            <Storyboard>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="BlurRadius" To="30" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="Opacity" To="0.9" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glow" Storyboard.TargetProperty="Opacity" To="1" Duration="0:0:0.3"/>
                                                <ColorAnimation Storyboard.TargetName="border" Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)" To="#7068ff" Duration="0:0:0.3"/>
                                            </Storyboard>
                                        </BeginStoryboard>
                                    </Trigger.EnterActions>
                                    <Trigger.ExitActions>
                                        <BeginStoryboard>
                                            <Storyboard>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="BlurRadius" To="0" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glowEffect" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.3"/>
                                                <DoubleAnimation Storyboard.TargetName="glow" Storyboard.TargetProperty="Opacity" To="0" Duration="0:0:0.3"/>
                                                <ColorAnimation Storyboard.TargetName="border" Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)" To="#5e5cff" Duration="0:0:0.3"/>
                                            </Storyboard>
                                        </BeginStoryboard>
                                    </Trigger.ExitActions>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

    $resReader = New-Object System.Xml.XmlNodeReader ([xml]$resXaml)
    $resWindow = [Windows.Markup.XamlReader]::Load($resReader)

    $resWindow.FindName("ResClose").Add_Click({ $resWindow.Close() })
    $resWindow.Add_MouseLeftButtonDown({ $resWindow.DragMove() })

    $dgv = $resWindow.FindName("DgvResults")
    $dgv.ItemsSource = $Results

    $collectionView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($dgv.ItemsSource)
    if ($collectionView) {
        $groupDesc = New-Object System.Windows.Data.PropertyGroupDescription("Mailbox")
        $collectionView.GroupDescriptions.Add($groupDesc)
    }

    $resStatus = $resWindow.FindName("ResStatus")
    $resStatus.Text = "$($Mailboxes.Count) mailboxes scanned - $RealCount items found"

    $btnRefresh = $resWindow.FindName("BtnRefresh")
    $btnPreview = $resWindow.FindName("BtnPreview")
    $btnDelete = $resWindow.FindName("BtnDelete")
    $btnAddRule Title  = $resWindow.FindName("BtnAddRule Title")

    $btnRefresh.Add_Click({
        $resWindow.Close()
        Show-Results
    })

    $btnPreview.Add_Click({
        $selected = $dgv.SelectedItem
        if (-not $selected -or $selected.Type -in @("No Rules Found", "Access Error")) { return }
        $key = $selected.Key
        if ($key -eq "dummy" -or $key -eq "error") { return }
        $data = $global:rawResults[$key]
        Show-PreviewWindow -Row $selected -Data $data
    })

    $btnDelete.Add_Click({
        $selectedItems = $dgv.SelectedItems | Where-Object { $_.Type -notin @("No Rules Found", "Access Error") }
        if ($selectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No rules selected.")
            return
        }

        $toDelete = @()
        $backupLines = @()

        $backupLines += "Backup created at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $backupLines += ""

        foreach ($item in $selectedItems) {
            $key = $item.Key
            $data = $global:rawResults[$key]
            if ($data.ID -eq "SMTP") {
                [System.Windows.Forms.MessageBox]::Show("Cannot delete SMTP forwarding setting for $($item.Mailbox). Skipping.")
                continue
            }

            $rawRule = $data.Raw

            $subjectContains = if ($rawRule.conditions.subjectContains) { $rawRule.conditions.subjectContains -join "; " } else { "" }
            $actionType = if ($rawRule.actions.forwardTo) { "Forward" }
                         elseif ($rawRule.actions.redirectTo) { "Redirect" }
                         elseif ($rawRule.actions.moveToFolder) { "MoveToFolder" }
                         elseif ($rawRule.actions.delete) { "Delete" }
                         else { "Other" }
            $actionTarget = if ($rawRule.actions.forwardTo) { $rawRule.actions.forwardTo.emailAddress.address -join "; " }
                           elseif ($rawRule.actions.redirectTo) { $rawRule.actions.redirectTo.emailAddress.address -join "; " }
                           elseif ($rawRule.actions.moveToFolder) { $rawRule.actions.moveToFolder }
                           else { "" }

            $backupLines += "Deleted rules for mailbox: $($item.Mailbox)"
            $backupLines += "================================================"
            $backupLines += "Rule Name: $($item.RuleName)"
            $backupLines += "From: $($item.From)"
            $backupLines += "Subject Contains: $subjectContains"
            $backupLines += "Action: $actionType"
            if ($actionTarget) { $backupLines += "To: $actionTarget" }
            $backupLines += ""

            $toDelete += [PSCustomObject]@{ Item = $item; Data = $data }
        }

        if ($toDelete.Count -eq 0) { return }

        if ($backupLines.Count -gt 2) {
            $backupFolder = "C:\Temp\Rule Engine Vault"
            if (-not (Test-Path $backupFolder)) {
                New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null
            }
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $backupPath = "$backupFolder\RuleBackup_$timestamp.txt"

            try {
                $backupLines | Out-File -FilePath $backupPath -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Backup saved to:`n$backupPath", "Backup Created")
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to save backup:`n$($_.Exception.Message)", "Backup Error")
            }
        }

        $confirmMsg = "Delete $($toDelete.Count) selected rule(s)?"
        $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Confirm Deletion", "YesNo", "Warning")
        if ($confirm -ne "Yes") { return }

        foreach ($del in $toDelete) {
            try {
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($del.Item.Mailbox)/mailFolders/inbox/messageRules/$($del.Data.ID)" -Method Delete -Headers @{ Authorization = "Bearer $AccessToken" }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to delete '$($del.Item.RuleName)' on $($del.Item.Mailbox)`n$($_.Exception.Message)")
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Selected rules deleted successfully.")
        $resWindow.Close()
        Show-Results
    })

    $btnAddRule Title.Add_Click({
        $targetMailboxes = Get-SelectedMailboxes -Mailboxes $Mailboxes
        if (-not $targetMailboxes -or $targetMailboxes.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No mailboxes selected.")
            return
        }

        $createdCount = 0
        $skippedCount = 0

        foreach ($mailbox in $targetMailboxes) {
            $headers = @{ Authorization = "Bearer $AccessToken" }

            try {
                $existingRules = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messageRules" -Headers $headers
                $existingNames = $existingRules.value.displayName
            } catch {
                $existingNames = @()
            }

            $rulesToCreate = @(
                @{ name = "Rule Title";      subject = "email address here"; from = "email address here"; to = "email address here" },
                @{ name = "Rule Title";     subject = $null; from = "email address here"; to = "notification@yourcompany.com.ca" },
                @{ name = "Rule Title"; subject = $null; from = "email address here"; to = "notification@yourcompany.com.ca" },
                @{ name = "Rule Title";  subject = "email address here"; from = "email address here"; to = "email address here" },
                @{ name = "Rule Title";                subject = "word"; from = "email address here"; to = "email address here" },
                @{ name = "Rule Title";     subject = $null; from = "email address here"; to = "email address here" },
                @{ name = "Rule Title";         subject = $null; from = "email address here"; to = "email address here" },
                @{ name = "Rule Title"; subject = @("subject keyword 1", "subject keyword 2"); from = "email address"; to = "email address here" }
            )

            foreach ($r in $rulesToCreate) {
                if ($existingNames -contains $r.name) {
                    $skippedCount++
                    continue
                }

                $body = @{
                    displayName = $r.name
                    sequence = 10
                    isEnabled = $true
                    conditions = @{ fromAddresses = @(@{ emailAddress = @{ address = $r.from } }) }
                    actions = @{
                        forwardTo = @(@{ emailAddress = @{ address = $r.to } })
                        stopProcessingRules = $true
                    }
                }
                if ($r.subject) {
                    if ($r.subject -is [array]) {
                        $body.conditions.subjectContains = $r.subject
                    } else {
                        $body.conditions.subjectContains = @($r.subject)
                    }
                }

                try {
                    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messageRules" `
                                      -Method Post `
                                      -Headers @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" } `
                                      -Body ($body | ConvertTo-Json -Depth 10)
                    $createdCount++
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Failed to create '$($r.name)' on $mailbox`n$($_.Exception.Message)", "Error")
                }
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Rule Title rules creation complete.`nCreated: $createdCount`nSkipped (already exists): $skippedCount", "Finished")
        $resWindow.Close()
        Show-Results
    })

    $resWindow.ShowDialog() | Out-Null
}

function Show-PreviewWindow {
    param($Row, $Data)

    $isSmtp = $Data.ID -eq "SMTP"
    $raw = $Data.Raw

    $previewText = ""

    if ($isSmtp) {
        $previewText += "SMTP Forwarding Configuration`r`n`r`n"
        foreach ($prop in $raw.PSObject.Properties) {
            if ($prop.Value) {
                $val = if ($prop.Value -is [array]) { $prop.Value -join "; " } else { $prop.Value }
                $previewText += "$($prop.Name): $val`r`n"
            }
        }
    } else {
        $rule = $raw
        $previewText += "Inbox Rule Configuration`r`n`r`n"
        $previewText += "Display Name: $($rule.displayName)`r`n"
        $previewText += "Enabled: $($rule.isEnabled)`r`n"
        $previewText += "Sequence: $($rule.sequence)`r`n`r`n"

        $previewText += "Conditions:`r`n"
        $hasCond = $false
        foreach ($prop in $rule.conditions.PSObject.Properties) {
            if ($prop.Value) {
                $hasCond = $true
                $val = if ($prop.Value -is [array]) {
                    if ($prop.Value[0].emailAddress) { $prop.Value.emailAddress.address -join "; " } else { $prop.Value -join "; " }
                } else { $prop.Value }
                $previewText += "  $($prop.Name): $val`r`n"
            }
        }
        if (-not $hasCond) { $previewText += "  (No specific conditions)`r`n" }

        $previewText += "`r`nActions:`r`n"
        $hasAct = $false
        foreach ($prop in $rule.actions.PSObject.Properties) {
            if ($null -ne $prop.Value -and ($prop.Value -isnot [bool] -or $prop.Value)) {
                $hasAct = $true
                $val = if ($prop.Value -is [array]) {
                    if ($prop.Value[0].emailAddress) { $prop.Value.emailAddress.address -join "; " } else { $prop.Value -join "; " }
                } else { $prop.Value }
                $previewText += "  $($prop.Name): $val`r`n"
            }
        }
        if (-not $hasAct) { $previewText += "  (No actions)`r`n" }
    }

    $prevXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Rule Preview" Height="400" Width="800" 
        AllowsTransparency="True" WindowStyle="None" Background="Transparent" WindowStartupLocation="CenterScreen">
    <Border CornerRadius="20" Background="#1C1C1E" BorderBrush="#3A3A3C" BorderThickness="1">
        <Border.Effect><DropShadowEffect BlurRadius="15" Color="#000000" Opacity="0.1" ShadowDepth="5"/></Border.Effect>
        <Grid>
            <Grid VerticalAlignment="Top" Margin="30,25,30,0">
                <TextBlock Text="Rule Preview" FontSize="26" FontWeight="Bold" Foreground="#FFFFFF"/>
                <Button Name="PrevClose" Content="X" HorizontalAlignment="Right" VerticalAlignment="Top" Background="Transparent" BorderThickness="0" Foreground="#AEAEB2" Cursor="Hand" FontSize="12">
                    <Button.Style>
                        <Style TargetType="Button">
                            <Setter Property="Foreground" Value="#AEAEB2"/>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Foreground" Value="#FF0000"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </Button.Style>
                </Button>
            </Grid>

            <ScrollViewer Margin="30,85,30,30" VerticalScrollBarVisibility="Auto">
                <TextBox Name="TxtPreview" IsReadOnly="True" TextWrapping="Wrap" Background="#2C2C2E" Foreground="#FFFFFF" BorderThickness="0" FontSize="14" Padding="15" FontFamily="Segoe UI"/>
            </ScrollViewer>
        </Grid>
    </Border>
</Window>
"@

    $prevReader = New-Object System.Xml.XmlNodeReader ([xml]$prevXaml)
    $prevWindow = [Windows.Markup.XamlReader]::Load($prevReader)

    $prevWindow.Title = if ($isSmtp) { "SMTP Forwarding Preview" } else { "Inbox Rule Preview: $($Row.RuleName)" }

    $txtPreview = $prevWindow.FindName("TxtPreview")
    $txtPreview.Text = $previewText

    $prevWindow.FindName("PrevClose").Add_Click({ $prevWindow.Close() })
    $prevWindow.Add_MouseLeftButtonDown({ $prevWindow.DragMove() })

    $prevWindow.ShowDialog() | Out-Null
}

$window.ShowDialog() | Out-Null