<#
.Synopsis
   Plays music on Windows Media PLayer in backgroud.
.DESCRIPTION
   Function automates Windows Media Player to play songs in Background in order (Random/Sequential) chosen by the user.
   Moreover, It generates a balloon notification in bottom Right corner of the screen, whenever a new song starts playing and continues to do that until manually stopped or it completes playing all songs.
.PARAMETER Path
    Music Directory [ Example - Music -Path C:\music\ ]
.PARAMETER Shuffle
    Switch to play music in shuffle mode, default value is 'sequential'
.PARAMETER Loop
    Switch to continuously play songs in a infinite Loop.
.PARAMETER Stop
    Switch to kill any instance of Music playing in backgroud.
.EXAMPLE
    PS Root\> Music F:\Data\Music\MyPlaylist
    TotalSongs TotalPlayDuration Mode      
    ---------- ----------------- ----      
    26         98 Mins           Sequential 
   
    Example shows how to run music from by passing a music directory to the Function
.EXAMPLE
    PS Root\> Music -Verbose
    [VERBOSE] You've not provided a music directory, looking for cached information from Previous use.
    [VERBOSE] Starting a background Job to play Music files
    
    TotalSongs TotalPlayDuration Mode      
    ---------- ----------------- ----      
    26         98 Mins           Sequential

    Example shows that in case you don't provide a music directory, the function Looks for the cached information of the diretory from previous us of the function.
    Moreover, It displays the information like TotalSongs, Total play duration, and Mode chosen by the user
.EXAMPLE
    PS Root\> Music -Shuffle
    TotalSongs TotalPlayDuration Mode   
    ---------- ----------------- ----   
    26         98 Mins           Shuffle

    Choose 'Shuffle' switch to play music in shuffle mode, default value is 'sequential'.
.EXAMPLE
    PS Root\> Music -Shuffle -Loop
    TotalSongs TotalPlayDuration Mode           
    ---------- ----------------- ----           
    26         Infinite Mins     Shuffle in Loop 

    Choose 'Loop' switch inorder to continuously play songs in a infinite Loop.
.EXAMPLE
    PS Root\> Music -Stop -Verbose
    [VERBOSE] Stoping any Already running instance of Media in background.

    When 'Stop' switch is used any instance Music playing in backgroud stops.
.NOTES
   Author Name    : Prateek Singh
   Twitter Handle : @SinghPrateik
   Blog           : http://geekeefy.wordpress.com
#>
Function Start-MediaPlayer
{
    [cmdletbinding()]
    Param(
            [Alias('P')]  [String] $Path,
            [Alias('Sh')] [switch] $Shuffle,
            [Alias('St')] [Switch] $Stop,
            [Alias('L')]  [Switch] $Loop
    )

    If($Stop)
    {
        Write-Verbose "Stoping any Already running instance of Media in background."
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
                Write-Verbose "You've not provided a music directory, looking for cached information from Previous use."
                $path = cat C:\Temp\Musicplayer.txt

                If(-not (Test-Path $Path))
                {
                    "Please provide a path to a music directory.`nFound a cached directory `"$Path`" from previous use, but that too isn't accessible!"
                    # Mark Path as Empty string, If Cached path doesn't exist
                    $Path = ''     
                }
            }
            else
            {
                "Please provide a path to a music directory."
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
                        
                        # Crunching the numbers and Information
                        $FileList = gci $Path -Recurse -Include *.mp* | select fullname, @{n='Duration';e={get-songduration $_.fullname}}
                        $FileCount = $FileList.count
                        $TotalPlayDuration =  [Math]::Round(($FileList.duration | measure -Sum).sum /60)
                        
                        # Condition to identifed the Mode chosed by the user
                        if($Shuffle)
                        {
                            $Mode = "Shuffle"
                            $FileList = $FileList | Sort-Object {Get-Random}  # Find the target Music Files and sort them Randomly
                        }
                        Else
                        {
                            $Mode = "Sequential"
                        }
                        
                        # Check If user chose to play songs in Loop
                        If($Loop)
                        {
                            $Mode = $Mode + " in Loop"
                            $TotalPlayDuration = "Infinite"
                        }
                        
                        If($FileList)
                        {    
    	                    ''| select @{n='TotalSongs';e={$FileCount};},@{n='PlayDuration';e={[String]$TotalPlayDuration + " Mins"}},@{n='Mode';e={$Mode}} 
                        }
                        else
                        {
                            "No music files found in directory `"$path`" ." 
                        }
                        
                        Do
                        {
    	                    $FileList |%{
                                            $CurrentSongDuration= New-TimeSpan -Seconds (Get-SongDuration $_.fullname)
                                            $Message = "Song : "+$(Split-Path $_.fullname -Leaf)+"`nPlay Duration : $($CurrentSongDuration.Minutes) Mins $($CurrentSongDuration.Seconds) Sec`nMode : $Mode"            
		                    		        $MediaPlayer.Open($_.FullName)					# 1. Open Music file with media player
		                    		        $MediaPlayer.Play()								# 2. Play the Music File
		                    		        Show-NotifyBalloon ($Message)                   # 3. Show a notification balloon in system tray
		                    		        Start-Sleep -Seconds $_.duration                # 4. Pause the script execution until song completes
		                    		        $MediaPlayer.Stop()                             # 5. Stop the Song
		                    		        $Balloon.Dispose();$Balloon.visible =$false                           
	                        }
                        }While($Loop) # Play Infinitely If 'Loop' is chosen by user
                    }
        }

        # Removes any already running Job, and start a new job, that looks like changing the track
        If($(Get-Job Musicplayer -ErrorAction SilentlyContinue))
        {
            Get-Job MusicPlayer -ErrorAction SilentlyContinue |Remove-Job -Force
        }

        # Run only if path was Defined or retrieved from cached information
        If($Path)
        {
            Write-Verbose "Starting a background Job to play Music files"
            Start-Job -Name MusicPlayer -InitializationScript $init -ScriptBlock {playmusic $args[0] $args[1] $args[2]} -ArgumentList $path, $Shuffle, $Loop | Out-Null
            Start-Sleep -Seconds 3       # Sleep to allow media player some breathing time to load files
            Receive-Job -Name MusicPlayer | ft @{n='TotalSongs';e={$_.TotalSongs};alignment='left'},@{n='TotalPlayDuration';e={$_.PlayDuration};alignment='left'},@{n='Mode';e={$_.Mode};alignment='left'} -AutoSize
        }
 }
    
}

Set-Alias -Name Music -Value Start-MediaPlayer
