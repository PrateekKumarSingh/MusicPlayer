Param($Path)

# Function to calculate duration of song in Seconds
Function Get-SongDuration($FullName)
{
	$Shell = New-Object -COMObject Shell.Application
	$Folder = $shell.Namespace($(Split-Path $FullName))
	$File = $Folder.ParseName($(Split-Path $FullName -Leaf))
	
	[int]$h, [int]$m, [int]$s = ($Folder.GetDetailsOf($File, 27)).split(":")
	
	$h*60*60 + $m*60 +$s
}

# Function to Notify Information balloon message in system Tray
Function Show-NotifyIcon($Message)
{
	[system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null            
	$Global:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon            
	$NotifyIcon.Icon =  [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -id $pid | Select-Object -ExpandProperty Path))                    
	$NotifyIcon.BalloonTipIcon = 'Info'           
	$NotifyIcon.BalloonTipText = $Message            
	$NotifyIcon.BalloonTipTitle = 'Now Playing ..'            
	$NotifyIcon.Visible = $true            
	$NotifyIcon.ShowBalloonTip(1000)
}

# Function to Play\Stop songs
Function Start-PSRadio($path)
{
	# Calling required assembly
	Add-Type -AssemblyName PresentationCore

	# Instantiate Media Player Class
	$MediaPlayer = New-Object System.Windows.Media.Mediaplayer
	
	# Find the target Music Files and sort them Randomly
	$FileList = gci $Path -Recurse -Include *.mp3 | `
	
	select fullname, @{n='Duration';e={get-songduration $_.fullname}} | Sort-Object {Get-Random}
	
	$FileList |%{
				$MediaPlayer.Open($_.FullName)					# 1. Open Music file with media player
				$MediaPlayer.Play()								# 2. Play the Music File
				Show-NotifyIcon (Split-Path $_.fullname -Leaf)  # 3. Show a notification balloon in system tray
				Start-Sleep -Seconds $_.duration                # 4. Pause the script execution until song completes
				$MediaPlayer.Stop()                             # 5. Stop the Song
				$NotifyIcon.Dispose()                           
	}
}

# Function Call
Start-PSRadio $Path



