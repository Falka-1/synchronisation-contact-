##Declare Parameters##
$clientSecret =
$clientID = 
$tenantID = 
$CSVPath = "C:\Scripts\Contacts\main.txt"
$mailboxpath = "C:\Scripts\Contacts\cible.txt"
$mailbox1 = get-content -Path $mailboxpath
$mailbox = $mailbox1 | ForEach-Object { $_.ToString() }
$module = Get-Module -Name Microsoft.Graph
<#
"----Installation --ET-- importation du module Pwsh Microsoft Graph pour executer le script---" -ForegroundColor red
"----------Processus LONG, patientez 10/15min------"      
Install-module -Name Microsoft.Graph -verbose
Import-Module -Name Microsoft.Graph -verbose
#>
##Connexion###
$uri = "https://login.microsoftonline.com/*/oauth2/v2.0/token"
#
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }  
#
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -verbose  
#
write-host "-------------Obtention & Conversion du token en cours...-------------" -ForegroundColor green
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
#
    if ($? -eq $true){
        write-host "-------------Token genere >> securisation en cours-------------" -ForegroundColor green
                    }
else{
    write-host "-------------Erreur sur la recuperation du token-------------" -ForegroundColor red
    break
    }
$tokensecure = ConvertTo-SecureString -string $token -AsPlainText -Force -verbose
#
    if ($? -eq $true){
        write-host "-------------Token sécurisé obtenu, connexion...-------------" -ForegroundColor green
                try{
                    Connect-MgGraph -NoWelcome -AccessToken $tokensecure
                    "---------------------------------------"
                    "------------#CONNEXION OK!#------------"
                    "---------------------------------------"
                }catch{
                    write-host "-------------Connexion impossible avec le token spécifié-------------" -ForegroundColor red
                    break
                      }
}

foreach ($person in $mailbox ){

##Connexion###
$uri = "https://login.microsoftonline.com/*/oauth2/v2.0/token"
#
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }  
#
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -verbose  
#
write-host "-------------Obtention & Conversion du token en cours...-------------" -ForegroundColor green
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
#
    if ($? -eq $true){
        write-host "-------------Token genere >> securisation en cours-------------" -ForegroundColor green
                    }
else{
    write-host "-------------Erreur sur la recuperation du token-------------" -ForegroundColor red
    break
    }
$tokensecure = ConvertTo-SecureString -string $token -AsPlainText -Force -verbose
#
    if ($? -eq $true){
        write-host "-------------Token sécurisé obtenu, connexion...-------------" -ForegroundColor green
                try{
                    Connect-MgGraph -NoWelcome -AccessToken $tokensecure
                    "---------------------------------------"
                    "------------#CONNEXION OK!#------------"
                    "---------------------------------------"
                }catch{
                    write-host "-------------Connexion impossible avec le token spécifié-------------" -ForegroundColor red
                    break
                      }
}
#########################################"

$folder = Get-MgUserContactFolder -UserID $person | ? {$_.DisplayName -eq "Sync-Opac"}
 New-MgUserContactFolder -UserId $person -DisplayName Sync-Opac
                            start-sleep 5
$folder = Get-MgUserContactFolder -UserID $person | ? {$_.DisplayName -eq"Sync-Opac"}
                            $folderid = $folder.Id
                            ./pwsh_graph_contacts.ps1
                            clear-host
                            write-host "----------Contact importé && Dossier Contacts créé pour $person------------" -ForegroundColor Green
                       
                          
                            } 
