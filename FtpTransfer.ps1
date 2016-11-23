<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.128
	 Created on:   	13/10/2016 14:29
	 Created by:   	Antonio de Almeida
	 Organization: 	antoniodealmeida.net
	 Filename:     	FtpTransfer.ps1
	 Version:		1.0
	===========================================================================
	.DESCRIPTION
		Permet la gestion du transfert SFTP des fichiers contenus dans des répertoires vers
		un serveur SFTP.
		Ce script utilise l'assembly WinSCP .NET (https://winscp.net/eng/download.php)
#>

# DEBUT DU SCRIPT

# STOP A LA 1ERE ERREUR POWERSHELL
$ErrorActionPreference = "Stop"

cls
Get-Date -uformat "%Hh%M(%S) : DEBUT DU SCRIPT POWERSHELL`n"

if (-NOT (Test-Path "$PSScriptRoot\WinSCPnet.dll"))
{
	
	Get-Date -uformat "%Hh%M(%S) : ERR. : Assembly WinSCP .NET manquante"
	exit 1
	
}
# Load WinSCP .NET assembly
Add-Type -Path "$PSScriptRoot\WinSCPnet.dll"



# Infos
$Server = "mon_serveur_SFTP"
$RemotePath = "/content/"
$LocalPath = "\\Mon_serveur_local\partage$"
$ArchivesPath = "$PSScriptRoot\ARCHIVES_INFOS"
$NotifEmail = "adealmeida@netc.fr"
$port = 22

# Set the credentials
$login = "login"
$cryptedMDP = '76492d1116743f0423413b16050a534Gh8B8AHcAYgBxADgA8ABRAEUAMAArAGcAcQBMAFYAagB5AEcAQwAzAEQAaAA2AHcAPQA9AHwAYQAyADfANgAzADgAMgAzAGMAMQBlAGYAZgBkAGUAYgAxADcAYgA5AGIAOQBhADkANgA0AGEAHAAzAGIAYQA2ADcAYgA2AGQAYQBjAGEAMAA2ADgAZgA2ADMAOAFjADgAMwA3AGMAZAA3AGIAMwAwADcA8AA0AGUAZgAxADMANAA='
$SSHkey = 'ssh-ed25519 256 dc:57:5a:03:11:fa:7f:06:1c:02:54:5c:79:21:d6:02'

#-------------------------------------------------------------------------------------
#A savoir: les variables $cryptedMDP et $SSHkey au dessus contienent des valeurs d'exemples.
# Vous devez générez vos propres valeurs
#
#Définir un nouveau mdp chiffré
#En console taper $SecureString = Read-Host -AsSecureString et entrez votre mdp
#Toujours en console taper $StandardString = ConvertFrom-SecureString $SecureString -Key (1..16) >> text.txt
#Ensuite copier coller le contenu du fichier test.txt dans la variable $cryptedMDP au dessus.
# pour $SSHkey, connectez vous en SFTP avec Filezilla par exemple et recupérez la clé qu'il vous fournira
#-------------------------------------------------------------------------------------


$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $login, ($cryptedMDP | ConvertTo-SecureString -Key (1..16))
$decryptedMDP = $credential.GetNetworkCredential().password

If (!(test-path $ArchivesPath))
{
	New-Item -ItemType Directory -Force -Path $ArchivesPath
}

function FileEveryday([parameter(Mandatory)][string]$Path)
{
	# Check si pour chaque agence il y'a bien le fichier du jour
	
	$FoldersList = GCI $Path | Where-Object { $_.Name.Length -eq 3 -and $_.PSIsContainer -and $_.Name -notmatch "999" -and $_.Name -notmatch "_" }
	
	
	foreach ($folder in $FoldersList)
	{
		$files = GCI $folder.FullName #*.csv
		$i = 1
		foreach ($file in $files)
		{
			if ($i -lt 2)
			{
				$DayFile = $file | Where-Object{ $_.CreationTime -lt (Get-Date).AddDays(1) }
				if ($DayFile)
				{
					Write-Host "Au moins un fichier nomme $DayFile.Name datant de ce jour dans $folder"
				}
				if (!$DayFile)
				{
					Write-Host "Fichier dans $folder obsolete !"
					$EmptyAgency += $folder.Name + "<BR>"
				}
			}
			$i++
		}
		if (!$files)
		{ 
		Write-Host "Dossier $folder vide !"
			$EmptyAgency += $folder.Name + "<BR>"
		}
	}
	return $EmptyAgency
}

function FtpTransfer()
{


	# Setup session options
	$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
		Protocol = [WinSCP.Protocol]::Sftp
		HostName = $Server
		UserName = $login
		Password = $decryptedMDP
		SshHostKeyFingerprint = $SSHkey
	}


	'Establish SFTP connection'
	$session = New-Object WinSCP.Session


	try
	{
		'Connect'
		$session.Open($sessionOptions)
		
		'Upload'
		$session.SynchronizeDirectories('Remote', $LocalPath, $RemotePath, $false, $true).Check()
		
		$result = FileEveryday -Path $LocalPath
		'ENVOI DU RECEPITULATIF DES DOSSIERS VIDE OU CONTENANT DES FICHIERS OBSOLETES PAR EMAIL'
		SendMail -body $result
		Archive
		
	}
	catch
	{
		$ErrorMessage = $_.Exception.Message
		$FailedItem = $_.Exception.ItemName
		
		Get-Date -uformat "%Hh%M(%S) : ERROR : $ErrorMessage $FailedItem"
		return $Global:rc = 1
	}
	finally
	{
		'Disconnect, clean up'
		$session.Dispose()
	}
	
	
}

#Permet l'envoi d'un email récapitulatif des sites n'ayant pas de fichier du jour
	###### DEBUT DE LA FONCTIONS MAIL ######
	Function SendMail
	{
		
		param (
			
			[Parameter(Mandatory = $False)]
			[string]$from = "adealmeida@netc.fr",
			[Parameter(Mandatory = $false)]
			[string]$to = $NotifEmail, #"$($xml.flux.alerte.mail.to)"
			[Parameter(Mandatory = $false)]
			[string]$cc = "", #"$($xml.flux.alerte.mail.cc)"
			[Parameter(Mandatory = $false)]
			[string]$Subject = "MON TITRE",
			[Parameter(Mandatory = $false)]
			$body , #= $(get-content "$Path_Temp\body.html")
			[Parameter(Mandatory = $false)]
			[string]$pj,
			[Parameter(Mandatory = $false)]
			[string]$smtpserver = "mon_serveur_smtp"
			
		)
		
		# TEMPLATE DU MESSAGE MAIL
		$body_html = @"
<br>
Bonjour,
<br><br>
Voici les dossiers vides ou contenant des fichiers obsoletes sur le serveur MON_SERVEUR dans le dossier E:\CONTENT\FTP\PROJET:<BR>
<br><br>
$body
<br><br>
FtpTransfer Script
"@
		
		# CONSTRUCTION DU MAIL
		$message = new-object System.Net.Mail.MailMessage
		
		# FROM
		$message.From = $from
		# TO
		if ("$to" -eq "") { $message.To.Add("adealmeida@netc.fr") }
		else { $message.To.Add($to) }
		# CC
		if ("$cc" -ne "") { $message.To.Add($cc) }
		
		# ENVOI AU FORMAT HTML
		$message.IsBodyHtml = $True
		
		# PIECE JOINTES
		if ("$pj" -ne "")
		{
			$attach = new-object Net.Mail.Attachment($pj)
			$message.Attachments.Add($attach)
		}
		
		# SUJET
		$message.Subject = $Subject
		$message.body = $body_html
		
		
		# ENVOI DU MAIL
		$smtp = new-object Net.Mail.SmtpClient($smtpserver)
		$smtp.Send($message)
		if ($? -eq $true) { Get-Date -uformat "%Hh%M(%S) : MAIL ENVOYE" }
		
		# MENAGE
		#Remove-Item "$Path_Temp\body.html"
		
	}
	###### FIN DE LA FONCTIONS MAIL ######


function Archive
{#Permet d'archiver et compresser les fichiers avant l'envoi sur le SFTP
	$timestamp = Get-Date -f dd-MM-yyyy_HH_mm_ss
	If (!(test-path "$ArchivesPath\_ARCHIVES\$timestamp"))
	{
		New-Item -ItemType Directory -Force -Path "$ArchivesPath\_ARCHIVES\$timestamp"
	}
	
	$FlagIfFilesExists = $false
	
	#Archivage de l'ensemble des répertoires
    Get-ChildItem $LocalPath -Directory |
	Foreach-Object  {
		if (($_.enumeratefiles() | measure).count -gt 0)
		{ 
			#Invoke-Expression -Command "Copy-Item -path $_.fullname -Destination $ArchivesPath\_ARCHIVES\$timestamp -Recurse"
			Copy-Item -path "$LocalPath\$_" -Destination $ArchivesPath\_ARCHIVES\$timestamp -Recurse
			if ($? -eq $true)
			{
					Remove-Item "$LocalPath\$_\*" -Recurse
					if ($? -eq $true)
					{
					Get-Date -uformat "%Hh%M(%S) : INFO : Suppression des fichiers ayant ete archives..."
					$FlagIfFilesExists = $true
					}
				
			}
		}
	}
	
	if ($? -eq $true)
	{
		#On compresse l'archive
		if ($FlagIfFilesExists -eq $true)
		{
			Get-Date -uformat "%Hh%M(%S) : INFO : Essai de compression archive..."
			compress-archive -path "$ArchivesPath\_ARCHIVES\$timestamp" -update -CompressionLevel fastest -DestinationPath "$ArchivesPath\_ARCHIVES\$timestamp"
			Get-Date -uformat "%Hh%M(%S) : INFO : Compression archive OK..."
		}
		remove-item "$ArchivesPath\_ARCHIVES\$timestamp" -Force -Recurse
	}
}

FtpTransfer


if (-not ($Global:rc))
{	
	Get-Date -uformat "`n%Hh%M(%S) : FIN DU SCRIPT POWERSHELL"
	EXIT 0
}
