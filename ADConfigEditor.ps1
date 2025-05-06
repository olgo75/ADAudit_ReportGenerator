# Charger les assemblys nécessaires
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Définir l'interface XAML
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Configuration AD Report" Height="300" Width="600">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Margin" Value="5"/>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal">
            <Label x:Name="SMTPServerLabel" Content="Serveur SMTP (SMTP Server):" Grid.Row="0" Grid.Column="0" Margin="5"/>
            <TextBox x:Name="SMTPServer" Grid.Row="0" Grid.Column="1" Margin="5" ToolTip="Adresse du serveur SMTP" Width="250" VerticalAlignment="Center"/>
        </StackPanel>
        
        <StackPanel Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal">
            <Label x:Name="PortLabel" Content="Port SMTP:" Grid.Row="1" Grid.Column="0" Margin="5"/>
            <TextBox x:Name="Port" Grid.Row="1" Grid.Column="1" Margin="5" ToolTip="Port utilisé pour la connexion SMTP" Width="100" VerticalAlignment="Center" Text="587"/>
        </StackPanel>
        
        <StackPanel Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal">
            <Label x:Name="FromLabel" Content="Adresse email expéditeur (From):" Grid.Row="2" Grid.Column="0" Margin="5"/>
            <TextBox x:Name="From" Grid.Row="2" Grid.Column="1" Margin="5" ToolTip="Adresse email utilisée pour l'envoi des rapports" Width="200" VerticalAlignment="Center"/>
        </StackPanel>   
        
        <StackPanel Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal">
            <Label x:Name="ToLabel" Content="Adresse(s) email destinataires (To):" Grid.Row="3" Grid.Column="0" Margin="5"/>
            <TextBox x:Name="To" Grid.Row="3" Grid.Column="1" Margin="5" ToolTip="Adresse(s) email qui recevra les rapports (dans le cas de valeurs multiples, séparez-les par des virgules)" Width="300" VerticalAlignment="Center"/>
        </StackPanel>
        
        <StackPanel Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal">
            <Label x:Name="OutputPathLabel" Content="Répertoire de sortie (Output Path):" Grid.Row="4" Grid.Column="0" Margin="5"/>
            <TextBox x:Name="OutputPath" Grid.Row="4" Grid.Column="1" Margin="5" ToolTip="Dossier où seront stockés les rapports et les données" Width="250" VerticalAlignment="Center"/>
            <Button x:Name="BrowsePath" Content="Parcourir..." Grid.Row="4" Grid.Column="2"/>
        </StackPanel>
              
        <StackPanel Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal">
            <Label Content="" x:Name="PathLabel" Grid.Row="4" Grid.Column="1" Margin="5" Foreground="Gray"/>
        </StackPanel>
        
        <StackPanel Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="5">
            <Button x:Name="ConfigCredentials" Content="Configurer identifiants"/>
            <Button x:Name="CancelButton" Content="Annuler"/>
            <Button x:Name="SaveButton" Content="Enregistrer"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Créer la fenêtre
$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))

$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object { 
    New-Variable -Name $_.Name -Value $window.FindName($_.Name) -Force 
}

# Charger la configuration existante
$configPath = "$PSScriptRoot\ADReportConfig.xml"

if (Test-Path $configPath) {
    $config = [xml](Get-Content $configPath)
    $SMTPServer.Text = $config.Configuration.EmailSettings.SMTPServer
    $Port.Text = $config.Configuration.EmailSettings.Port
    $From.Text = $config.Configuration.EmailSettings.From
    $To.Text = $config.Configuration.EmailSettings.To
    $OutputPath.Text = $config.Configuration.Paths.OutputPath
    
}

[xml]$authXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Authentification | SMTP" Height="200" Width="400">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Margin" Value="5"/>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Orientation="Horizontal" HorizontalAlignment="Right" Margin="5">
            <Label Content="Nom d'utilisateur:" Margin="5"/>
            <TextBox x:Name="Username" Margin="5" Width="200" VerticalAlignment="Center"/>
        </StackPanel>

        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="5">
            <Label Content="Mot de passe:" Margin="5"/>
            <PasswordBox x:Name="Password" Margin="5" Width="200" VerticalAlignment="Center"/>
        </StackPanel>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="5">
            <Button x:Name="IDCancelButton" Content="Annuler"/>
            <Button x:Name="IDOKButton" Content="OK"/>
        </StackPanel>
    </Grid>
</Window>
"@

$ConfigCredentials = $window.FindName("ConfigCredentials")
$ConfigCredentials.Add_Click({
    # Créer la fenêtre d'authentification
    $authWindow = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $authXaml))
    $authXaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object { 
        New-Variable -Name $_.Name -Value $authWindow.FindName($_.Name) -Force 
    }
    
    # Gestionnaires d'événements
    $IDOKButton.Add_Click({
        $authWindow.DialogResult = $true
        $authWindow.Close()
    })
    
    $IDCancelButton.Add_Click({
        $authWindow.DialogResult = $false
        $authWindow.Close()
    })
    
    # Afficher la fenêtre
    if ($authWindow.ShowDialog()) {
        # Créer les identifiants
        $securePassword = $password.SecurePassword
        $credential = New-Object System.Management.Automation.PSCredential ($username.Text, $securePassword)
        
        # Exporter les identifiants
        $credentialPath = "$PSScriptRoot\SMTPCredentials.xml"
        $credential | Export-Clixml -Path $credentialPath
        
        [System.Windows.MessageBox]::Show("Identifiants sauvegardés avec succès", "Succès", "OK", "Information")
    }
})

$SaveButton.Add_Click({
    # Sauvegarder la configuration
    $config = "<?xml version='1.0' encoding='utf-8'?>\n<Configuration>\n" +
              "    <EmailSettings>\n" +
              "        <SMTPServer>$($SMTPServer.Text)</SMTPServer>\n" +
              "        <From>$($From.Text)</From>\n" +
              "        <To>$($To.Text)</To>\n" +
              "        <Subject>AD Audit Report - {0}</Subject>\n" +
              "    </EmailSettings>\n" +
              "    <Paths>\n" +
              "        <OutputPath>$($OutputPath.Text)</OutputPath>\n" +
              "    </Paths>\n" +
              "</Configuration>"
    
    $config | Set-Content "$PSScriptRoot\ADReportConfig.xml"
    $window.Close()
})

# Créer les gestionnaires d'événements
$BrowsePath.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $OutputPath.Text = $folderBrowser.SelectedPath
    }
})

$CancelButton.Add_Click({
    $window.Close() | out-null
})

$SaveButton.Add_Click({
    Save-Configuration
    $window.Close() | out-null
})

# Fonction pour sauvegarder la configuration
function Save-Configuration {
    $xml = New-Object System.Xml.XmlDocument
    $emailSettings = $xml.CreateElement("EmailSettings")
    $emailSettings.AppendChild($xml.CreateElement("SMTPServer")).InnerText = $SMTPServer.Text
    $emailSettings.AppendChild($xml.CreateElement("Port")).InnerText = $Port.Text
    $emailSettings.AppendChild($xml.CreateElement("From")).InnerText = $From.Text
    
    # Créer un élément To avec un attribut pour indiquer que c'est une liste
    $toElement = $xml.CreateElement("To")
    $toElement.SetAttribute("type", "list")
    $toElement.InnerText = $To.Text
    $emailSettings.AppendChild($toElement)
    
    $emailSettings.AppendChild($xml.CreateElement("Subject")).InnerText = "AD Audit Report - {0}"

    $paths = $xml.CreateElement("Paths")
    $paths.AppendChild($xml.CreateElement("OutputPath")).InnerText = $OutputPath.Text

    $schedule = $xml.CreateElement("Schedule")
    $schedule.AppendChild($xml.CreateElement("Time")).InnerText = $ExecutionTime.Text

    $config = $xml.CreateElement("Configuration")
    $config.AppendChild($emailSettings)
    $config.AppendChild($paths)
    $config.AppendChild($schedule)

    $xml.AppendChild($config)
    $xml.Save("$PSScriptRoot\ADReportConfig.xml")
}

# Charger la configuration existante
$configPath = "$PSScriptRoot\ADReportConfig.xml"
if (Test-Path $configPath) {
    $config = [xml](Get-Content $configPath)
    $SMTPServer.Text = $config.Configuration.EmailSettings.SMTPServer
    $Port.Text = $config.Configuration.EmailSettings.Port
    $From.Text = $config.Configuration.EmailSettings.From
    
    # Récupérer le champ To
    $toElement = $config.Configuration.EmailSettings.SelectSingleNode("To")
    if ($toElement -and $toElement.Attributes["type"] -eq "list") {
        $To.Text = $toElement.InnerText
    } else {
        $To.Text = $config.Configuration.EmailSettings.To.InnerText
    }
    
    $OutputPath.Text = $config.Configuration.Paths.OutputPath
}

# Afficher la fenêtre
$window.ShowDialog() | out-null
