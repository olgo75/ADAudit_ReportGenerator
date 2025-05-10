# Ajout du paramètre Verbose au début du script
[CmdletBinding()]
param()

# Configuration
$configPath = Join-Path $PSScriptRoot "ADReportConfig.xml"
if (-not (Test-Path $configPath)) {
    Write-Error "Fichier de configuration non trouvé : $configPath"
    exit 1
}

$config = [xml](Get-Content $configPath)
$OutputPath = $config.Configuration.Paths.OutputPath
$SMTPServer = $config.Configuration.EmailSettings.SMTPServer
$Port = $config.Configuration.EmailSettings.Port
$From = $config.Configuration.EmailSettings.From
$To = $config.Configuration.EmailSettings.To

# Charger les credentials SMTP
$credentialsPath = Join-Path $PSScriptRoot "SMTPCredentials.xml"
if (-not (Test-Path $credentialsPath)) {
    Write-Error "Fichier de credentials SMTP non trouvé : $credentialsPath"
    exit 1
}

try {
    $SMTPCredential = Import-Clixml -Path $credentialsPath
}
catch {
    Write-Error "Erreur lors de l'import des credentials SMTP : $_"
    exit 1
}

# Créer le dossier de sortie s'il n'existe pas
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force
}

# Fonction pour exporter les données AD
function Export-ADData {
    param($Path)
    
    # Utilisation de -Properties spécifiques au lieu de *
    $userProperties = @(
        'Name', 'SamAccountName', 'UserPrincipalName', 'Enabled', 
        'LastLogonDate', 'Description', 'DistinguishedName', 'memberOf'
    )
    
    $computerProperties = @(
        'Name', 'SamAccountName', 'DistinguishedName', 
        'OperatingSystem', 'OperatingSystemVersion'
    )
    
    $groupProperties = @(
        'Name', 'SamAccountName', 'Description', 
        'DistinguishedName'
    )

    # Utilisation de -LDAPFilter pour optimiser les requêtes
    $data = @{
        Users = Get-ADUser -LDAPFilter "(objectClass=user)" -Properties $userProperties | 
            Select-Object $userProperties
        Computers = Get-ADComputer -LDAPFilter "(objectClass=computer)" -Properties $computerProperties | 
            Select-Object $computerProperties
        Groups = Get-ADGroup -LDAPFilter "(objectClass=group)" -Properties $groupProperties | 
            ForEach-Object {
                $group = $_
                $membersResult = Get-GroupMembers -GroupSamAccountName $group.SamAccountName -GroupName $group.Name
                
                [PSCustomObject]@{
                    Name = $group.Name
                    SamAccountName = $group.SamAccountName
                    Description = $group.Description
                    DistinguishedName = $group.DistinguishedName
                    Members = $membersResult.Members
                    MembersAccessError = if (-not $membersResult.Success) { $membersResult.Error } else { $null }
                }
            }
        OUs = Get-ADOrganizationalUnit -Filter * | 
            Select-Object Name, DistinguishedName
    }
    
    # Compression des données avant l'export
    $jsonData = $data | ConvertTo-Json -Depth 10 -Compress
    $jsonData | Set-Content "$Path\ADData-$ReportDate.json"
}

# Fonction pour récupérer le jour ouvré précédent
function Get-PreviousWorkingDay {
    param($CurrentDate)
    
    $date = $CurrentDate
    do {
        $date = $date.AddDays(-1)
        # Ignorer les week-ends
    } while ($date.DayOfWeek -eq "Saturday" -or $date.DayOfWeek -eq "Sunday")
    
    return $date.ToString("yyyy-MM-dd")
}

# Fonction pour comparer les données
function Compare-ADData {
    param($CurrentPath, $PreviousPath)
    
    # Chargement des données avec un timeout
    $currentData = Get-Content "$CurrentPath\ADData-$ReportDate.json" -Raw | 
        ConvertFrom-Json -AsHashtable
    $previousData = Get-Content "$PreviousPath\ADData-$PreviousDate.json" -Raw | 
        ConvertFrom-Json -AsHashtable

    # Utilisation de hashtables pour les comparaisons
    $currentUsers = @{}
    $previousUsers = @{}
    $currentGroups = @{}
    $previousGroups = @{}

    # Préparation des hashtables
    foreach ($user in $currentData.Users) {
        $currentUsers[$user.SamAccountName] = $user
    }
    foreach ($user in $previousData.Users) {
        $previousUsers[$user.SamAccountName] = $user
    }
    foreach ($group in $currentData.Groups) {
        $currentGroups[$group.SamAccountName] = $group
    }
    foreach ($group in $previousData.Groups) {
        $previousGroups[$group.SamAccountName] = $group
    }

    # Comparaison optimisée
    $changes = @{
        Users = @{
            Created = $currentUsers.Keys | Where-Object { -not $previousUsers.ContainsKey($_) } | 
                ForEach-Object { $currentUsers[$_] }
            Modified = $currentUsers.Keys | Where-Object { 
                $previousUsers.ContainsKey($_) -and 
                ($currentUsers[$_] | ConvertTo-Json) -ne ($previousUsers[$_] | ConvertTo-Json)
            } | ForEach-Object { $currentUsers[$_] }
            Deleted = $previousUsers.Keys | Where-Object { -not $currentUsers.ContainsKey($_) } | 
                ForEach-Object { $previousUsers[$_] }
        }
        Groups = @{
            Created = $currentGroups.Keys | Where-Object { -not $previousGroups.ContainsKey($_) } | 
                ForEach-Object { $currentGroups[$_] }
            Modified = $currentGroups.Keys | Where-Object { 
                $previousGroups.ContainsKey($_) -and 
                ($currentGroups[$_] | ConvertTo-Json) -ne ($previousGroups[$_] | ConvertTo-Json)
            } | ForEach-Object { $currentGroups[$_] }
            Deleted = $previousGroups.Keys | Where-Object { -not $currentGroups.ContainsKey($_) } | 
                ForEach-Object { $previousGroups[$_] }
        }
    }
    
    return $changes
}

# Fonction pour générer le rapport HTML
function Generate-HTMLReport {
    param($Changes)
    
    $html = @"
    <!DOCTYPE html>
    <html>
    <head>
        <title>Rapport d'Audit AD - $ReportDate</title>
        <style>
            body { 
                font-family: Arial, sans-serif; 
                margin: 20px;
            }
            table { 
                border-collapse: collapse; 
                margin: 20px 0;
                display: inline-block;
            }
            th, td { 
                border: 1px solid #ddd; 
                padding: 8px; 
                text-align: left;
                word-wrap: break-word;
                max-width: 300px;
            }
            th { 
                background-color: #084C64; 
                color: white; 
            }
            .created { 
                background-color: #DB504A; 
            }
            .modified { 
                background-color: #E3B505; 
            }
            .deleted { 
                background-color: #4F6D7A; 
            }
            h1 {
                color: #000;
                font-weight: bold;
                margin-bottom: 20px;
            }
            h2 {
                color: #000;
                margin: 20px 0;
            }
            .changes {
                margin: 10px 0;
            }
            .table-container {
                max-width: 100%;
                overflow-x: auto;
            }
            .no-changes {
                color: #666;
                font-style: italic;
                margin: 20px 0;
                text-align: center;
            }
            .date-info {
                color: #666;
                font-style: italic;
                margin: 10px 0;
            }
        </style>
    </head>
    <body>
        <h1>Rapport d'Audit AD - $DisplayReportDate</h1>
        <p class="date-info">Comparaison avec le $DisplayPreviousDate</p>
        
        <h2>Changements des utilisateurs</h2>
        $(if ($Changes.Users.Created.Count -eq 0 -and $Changes.Users.Modified.Count -eq 0 -and $Changes.Users.Deleted.Count -eq 0) {
            "<div class='no-changes'>Aucun changement observé</div>"
        } else {
            "<div class='table-container'>" +
            "<table>" +
            "<tr>" +
            "<th>Nom</th>" +
            "<th>Nature du changement</th>" +
            "<th>Informations complémentaires</th>" +
            "</tr>"
        })

            # Créations
            $(foreach ($user in $Changes.Users.Created) {
                "<tr class='created'>" +
                "<td>$($user.Name) ($($user.SamAccountName))</td>" +
                "<td>Création</td>" +
                "<td></td>" +
                "</tr>"
            })

            # Modifications
            $(foreach ($user in $Changes.Users.Modified) {
                $current = $user.InputObject
                $previous = $Changes.Users.Previous | Where-Object SamAccountName -eq $current.SamAccountName
                
                $changesHTML = @()
                if ($current.Enabled -ne $previous.Enabled) {
                    $status = if ($current.Enabled) { 'Activé' } else { 'Désactivé' }
                    $changesHTML += "<div style='margin-left: 20px;'>Statut: $status</div>"
                }
                if ($current.Description -ne $previous.Description) {
                    $changesHTML += "<div style='margin-left: 40px;'>Description: '$($previous.Description)' → '$($current.Description)'</div>"
                }
                if ($current.UserPrincipalName -ne $previous.UserPrincipalName) {
                    $changesHTML += "<div style='margin-left: 40px;'>UPN: '$($previous.UserPrincipalName)' → '$($current.UserPrincipalName)'</div>"
                }
                
                $changesString = $changesHTML -join ''
                "<tr class='modified'>" +
                "<td>$($current.Name) ($($current.SamAccountName))</td>" +
                "<td>Modification</td>" +
                "<td>$changesString</td>" +
                "</tr>"
            })
            # Suppressions
            $(foreach ($user in $Changes.Users.Deleted) {
                "<tr class='deleted'>" +
                "<td>$($user.Name) ($($user.SamAccountName))</td>" +
                "<td>Suppression</td>" +
                "<td></td>" +
                "</tr>"
            })

                "<td></td>" +
                "</tr>"
            })
        </table>
        
        <h2>Changements des groupes</h2>
        $(if ($Changes.Groups.Created.Count -eq 0 -and $Changes.Groups.Modified.Count -eq 0 -and $Changes.Groups.Deleted.Count -eq 0) {
            "<div class='no-changes'>Aucun changement observé</div>"
        } else {
            "<div class='table-container'>" +
            "<table>" +
            "<tr>" +
            "<th>Nom</th>" +
            "<th>Nature du changement</th>" +
            "<th>Informations complémentaires</th>" +
            "</tr>"
        })

            # Créations
            $(foreach ($group in $Changes.Groups.Created) {
                "<tr class='created'>" +
                "<td>$($group.Name) ($($group.SamAccountName))</td>" +
                "<td>Création</td>" +
                "<td></td>" +
                "</tr>"
            })
            # Modifications
            $(foreach ($group in $Changes.Groups.Modified) {
                $current = $group.InputObject
                $previous = $Changes.Groups.Previous | Where-Object SamAccountName -eq $current.SamAccountName
                
                $changes = @()
                
                # Comparer les membres
                $currentMembers = Get-ADGroupMember -Identity $current.SamAccountName | Select-Object Name, SamAccountName
                $previousMembers = Get-ADGroupMember -Identity $previous.SamAccountName | Select-Object Name, SamAccountName
                
                $memberChanges = @()
                
                # Nouveaux membres
                $newMembers = Compare-Object $previousMembers $currentMembers -Property SamAccountName -PassThru | Where-Object SideIndicator -eq "=>"
                foreach ($member in $newMembers) {
                    $memberChanges += "Ajout: $($member.Name) ($($member.SamAccountName))"
                }
                
                # Membres supprimés
                $removedMembers = Compare-Object $previousMembers $currentMembers -Property SamAccountName -PassThru | Where-Object SideIndicator -eq "<="
                foreach ($member in $removedMembers) {
                    $memberChanges += "Suppression: $($member.Name) ($($member.SamAccountName))"
                }
                
                $changesHTML = @()
                if ($current.Name -ne $previous.Name) {
                    $changesHTML += @'
<div style="margin-left: 20px;">Nom: {0} → {1}</div>
'@ -f $previous.Name, $current.Name
                }
                
                # Ajouter les changements de membres
                if ($memberChanges.Count -gt 0) {
                    $changesHTML += @'
<div style="margin-left: 40px;">Membres:</div>
'@
                    
                    # Regrouper les changements par type
                    $groupedChanges = @{}
                    foreach ($change in $memberChanges) {
                        $match = $change -match '^(Ajout|Suppression):\s*(.+?)\s*\((.+?)\)$'
                        if ($match) {
                            $action = $matches[1]
                            $name = $matches[2]
                            $samAccount = $matches[3]
                            if (-not $groupedChanges.ContainsKey($action)) {
                                $groupedChanges[$action] = @()
                            }
                            if ($samAccount -like "CN=*") {
                                $groupedChanges[$action] += @'
<div style="margin-left: 20px;">{0}</div>
'@ -f $name
                            } else {
                                $groupedChanges[$action] += @'
<div style="margin-left: 20px;">{0} ({1})</div>
'@ -f $name, $samAccount
                            }
                        } else {
                            if (-not $groupedChanges.ContainsKey("Autres")) {
                                $groupedChanges["Autres"] = @()
                            }
                            $groupedChanges["Autres"] += @'
<div style="margin-left: 20px;">{0}</div>
'@ -f $change
                        }
                    }

                    # Afficher les changements regroupés
                    foreach ($action in $groupedChanges.Keys) {
                        $changesHTML += @'
<div style="margin-left: 60px;">{0}:</div>
'@ -f $action
                        $changesHTML += $groupedChanges[$action] -join ""
                    }
                }
                
                $changesString = $changesHTML -join ''
                @'
<tr class='modified'>
    <td>{0} ({1})</td>
    <td>Modification</td>
    <td>{2}</td>
</tr>
'@ -f $current.Name, $current.SamAccountName, $changesString
            })
            # Suppressions
            $(foreach ($group in $Changes.Groups.Deleted) {
                "<tr class='deleted'>" +
                "<td>$($group.Name) ($($group.SamAccountName))</td>" +
                "<td>Suppression</td>" +
                "<td></td>" +
                "</tr>"
            })
        </table>
        
        <h2>Changements des ordinateurs</h2>
        $(if ($Changes.Computers.Created.Count -eq 0 -and $Changes.Computers.Modified.Count -eq 0 -and $Changes.Computers.Deleted.Count -eq 0) {
            "<div class='no-changes'>Aucun changement observé</div>"
        } else {
            "<div class='table-container'>" +
            "<table>" +
            "<tr>" +
            "<th>Nom</th>" +
            "<th>Nature du changement</th>" +
            "<th>Informations complémentaires</th>" +
            "</tr>"
        })

            # Créations
            $(foreach ($computer in $Changes.Computers.Created) {
                "<tr class='created'>" +
                "<td>$($computer.Name) ($($computer.SamAccountName))</td>" +
                "<td>Création</td>" +
                "<td>OS: $($computer.OperatingSystem) $($computer.OperatingSystemVersion)</td>" +
                "</tr>"
            })

            # Modifications
            $(foreach ($computer in $Changes.Computers.Modified) {
                $current = $computer.InputObject
                $previous = $Changes.Computers.Previous | Where-Object SamAccountName -eq $current.SamAccountName
                $changesHTML = ""
                
                if ($current.OperatingSystem -ne $previous.OperatingSystem) {
                    $changesHTML += @'
<div style="margin-left: 20px;">Système d'exploitation: {0} → {1}</div>
'@ -f $previous.OperatingSystem, $current.OperatingSystem
                }
                if ($current.OperatingSystemVersion -ne $previous.OperatingSystemVersion) {
                    $changesHTML += @'
<div style="margin-left: 20px;">Version OS: {0} → {1}</div>
'@ -f $previous.OperatingSystemVersion, $current.OperatingSystemVersion
                }
                
                "<tr class='modified'>" +
                "<td>$($computer.Name) ($($computer.SamAccountName))</td>" +
                "<td>Modification</td>" +
                "<td>$changesHTML</td>" +
                "</tr>"
            })

            # Suppressions
            $(foreach ($computer in $Changes.Computers.Deleted) {
                "<tr class='deleted'>" +
                "<td>$($computer.Name) ($($computer.SamAccountName))</td>" +
                "<td>Suppression</td>" +
                "<td></td>" +
                "</tr>"
            })
        </table>

        <h2>Changements des unités d'organisation</h2>
        $(if ($Changes.OUs.Created.Count -eq 0 -and $Changes.OUs.Modified.Count -eq 0 -and $Changes.OUs.Deleted.Count -eq 0) {
            "<div class='no-changes'>Aucun changement observé</div>"
        } else {
            "<div class='table-container'>" +
            "<table>" +
            "<tr>" +
            "<th>Nom</th>" +
            "<th>Nature du changement</th>" +
            "<th>Informations complémentaires</th>" +
            "<th>Nom</th>"
            "<th>Nature du changement</th>"
            "</tr>"
        })
            # Créations
            $(foreach ($ou in $Changes.OUs.Created) {
                "<tr class='created'>" +
                "<td>$($ou.Name)</td>" +
                "<td>Création</td>" +
                "</tr>"
            })
            # Modifications
            $(foreach ($ou in $Changes.OUs.Modified) {
                "<tr class='modified'>" +
                "<td>$($ou.Name)</td>" +
                "<td>Modification</td>" +
                "</tr>"
            })
            # Suppressions
            $(foreach ($ou in $Changes.OUs.Deleted) {
                "<tr class='deleted'>" +
                "<td>$($ou.Name)</td>" +
                "<td>Suppression</td>" +
                "</tr>"
            })
        </table>
        </div>

        <h2>Groupes avec erreurs d'accès</h2>
        <table>
            <tr>
                <th>Nom du groupe</th>
                <th>Message d'erreur</th>
            </tr>
            $(foreach ($group in $currentData.Groups | Where-Object { $_.MembersAccessError }) {
                "<tr class='error'>
                    <td>$($group.Name) ($($group.SamAccountName))</td>
                    <td>$($group.MembersAccessError)</td>
                </tr>"
            })
        </table>
    </body>
    </html>
"@
    
    $html | Out-File "$OutputPath\ADReport-$ReportDate.html"
}

# Fonction pour envoyer le rapport par email
function Send-EmailReport {
    param($ReportPath)
    
    $subject = "AD Audit Report - $ReportDate"
    $body = Get-Content "$ReportPath" -Raw
    
    # Envoyer le message avec Send-MailMessage
    Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $From -To $To -Subject $subject -Body $body -BodyAsHtml -Credential $SMTPCredential
}

# Fonction pour récupérer les membres d'un groupe avec gestion d'erreur
function Get-GroupMembers {
    param(
        [string]$GroupSamAccountName,
        [string]$GroupName
    )
    
    try {
        $members = Get-ADGroupMember -Identity $GroupSamAccountName -Recursive -ErrorAction Stop | 
            Select-Object Name, SamAccountName
        return @{
            Success = $true
            Members = $members
            Error = $null
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Verbose "Impossible de récupérer les membres du groupe '$GroupName' ($GroupSamAccountName) : $errorMessage"
        return @{
            Success = $false
            Members = $null
            Error = $errorMessage
        }
    }
}

# Exécuter le script
try {
    Write-Verbose "Début de l'audit AD"
    
    # Exporter les données actuelles
    Write-Verbose "Export des données AD en cours..."
    Export-ADData -Path $OutputPath
    
    # Comparer avec la veille
    Write-Verbose "Comparaison des données en cours..."
    $changes = Compare-ADData -CurrentPath $OutputPath -PreviousPath $OutputPath
    
    # Générer le rapport HTML
    Write-Verbose "Génération du rapport HTML en cours..."
    Generate-HTMLReport -Changes $changes
    
    # Envoyer le rapport par email
    Write-Verbose "Envoi du rapport par email en cours..."
    Send-EmailReport -ReportPath "$OutputPath\ADReport-$ReportDate.html"
    
    Write-Verbose "Audit AD terminé avec succès"
} catch {
    Write-Verbose "Erreur lors de l'exécution du script: $_"
    Send-MailMessage -SmtpServer $SMTPServer -Port $Port -Credential $SMTPCredential -From $From -To $To -Subject "Erreur AD Audit Report" -Body "Une erreur s'est produite lors de l'exécution du script: $_"
}
