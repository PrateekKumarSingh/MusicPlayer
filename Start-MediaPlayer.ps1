<#
    Function to play songs in Windows media player from Powershell Console, change music tracks, shuffle, Loop, Stop.

    From : Prateek Singh - @SinghPrateik
    Blog : Geekeefy.wordpress.com
#> 
Function Start-MediaPlayer
{
    [cmdletbinding()]
    Param(
            [Alias('P')][String] $Path,
            [Alias('Sh')][switch] $Shuffle,
            [Alias('St')][Switch] $Stop,
            [Alias('L')][Switch] $Loop
    )

    If($Stop)
    {
        Get-Job MusicPlayer -ErrorAction SilentlyContinue | Remove-Job -Force
    }
    Else
    {       
        #Caches Path for next time in case you don't enter path to the music directory
        If($path)
        {
            $Path | out-file C:\Temp\Musicplayer.txt
        }
        else
        {
            If((cat C:\Temp\Musicplayer.txt -ErrorAction SilentlyContinue).Length -ne 0)
            {
                $path = cat C:\Temp\Musicplayer.txt

                If(-not (Test-Path $Path))
                {
                    "Please provide a path to music directory."       
                }
            }
            else
            {
                "Please provide a path to music directory."
            }
        }
 
        #initialization Script for back ground job
        $init = {
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
                    Function Show-NotifyBalloon($Message)
                    {
	                    [system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null            
	                    $Global:Balloon = New-Object System.Windows.Forms.NotifyIcon            
	                    $Balloon.Icon =  [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -id $pid | Select-Object -ExpandProperty Path))                    
	                    $Balloon.BalloonTipIcon = 'Info'           
	                    $Balloon.BalloonTipText = $Message            
	                    $Balloon.BalloonTipTitle = 'Now Playing'            
	                    $Balloon.Visible = $true            
	                    $Balloon.ShowBalloonTip(1000)
                    }
                    
                    Function PlayMusic($path, $Shuffle, $Loop)
                    {
    	                # Calling required assembly
    	                Add-Type -AssemblyName PresentationCore
    
    	                # Instantiate Media Player Class
    	                $MediaPlayer = New-Object System.Windows.Media.Mediaplayer
                        
                        # Crunching the numbers and Info
                        $FileList = gci $Path -Recurse -Include *.mp3 | select fullname, @{n='Duration';e={get-songduration $_.fullname}}
                        $FileCount = $FileList.count
                        $TotalPlayDuration =  [Math]::Round(($FileList.duration | measure -Sum).sum /60)
                        
                        if($Shuffle)
                        {
                            $Mode = "Shuffle"
                            $FileList = $FileList | Sort-Object {Get-Random}  # Find the target Music Files and sort them Randomly
                        }
                        Else
                        {
                            $Mode = "Sequential"
                        }
                        
                        If($Loop)
                        {
                            $Mode = $Mode + " in Loop"
                            $TotalPlayDuration = "Infinite"
                        }
                            
    	                ''| select @{n='TotalSongs';e={$FileCount};},@{n='PlayDuration';e={[String]$TotalPlayDuration + " Mins"}},@{n='Mode';e={$Mode}} 
                        
                        Do
                        {
    	                    $FileList |%{
                                            $CurrentSongDuration= New-TimeSpan -Seconds (Get-SongDuration $_.fullname)
                                            $Message = "Current Song : "+$(Split-Path $_.fullname -Leaf)+"`nPlay Duration : $($CurrentSongDuration.Minutes) Mins $($CurrentSongDuration.Seconds) Sec`nMode : $Mode"            
		                    		        $MediaPlayer.Open($_.FullName)					# 1. Open Music file with media player
		                    		        $MediaPlayer.Play()								# 2. Play the Music File
		                    		        Show-NotifyBalloon ($Message)  # 3. Show a notification balloon in system tray
		                    		        Start-Sleep -Seconds $_.duration                # 4. Pause the script execution until song completes
		                    		        $MediaPlayer.Stop()                             # 5. Stop the Song
		                    		        $NotifyIcon.Dispose()                           
	                        }
                        }While($Loop)
                    }
        }

        #Removes any already running Job, and start a new job, that looks like changing the track

        If($(Get-Job Musicplayer -ErrorAction SilentlyContinue))
        {
            Get-Job MusicPlayer -ErrorAction SilentlyContinue |Remove-Job -Force
        }

        Start-Job -Name MusicPlayer -InitializationScript $init -ScriptBlock {playmusic $args[0] $args[1] $args[2]} -ArgumentList $path, $Shuffle, $Loop | Out-Null
        Start-Sleep -s 3 # Sleep to allow media player some breathing time to load files
        Receive-Job -Name MusicPlayer | ft @{n='TotalSongs';e={$_.TotalSongs};alignment='left'},@{n='TotalPlayDuration';e={$_.PlayDuration};alignment='left'},@{n='Mode';e={$_.Mode};alignment='left'} -AutoSize
 }
    
}
