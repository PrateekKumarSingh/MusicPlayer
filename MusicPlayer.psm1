<#
.Synopsis
   Plays music on Windows Media Player in backgroud.
.DESCRIPTION
   Invoke-MusicPlayer Cmdlet automates Windows Media Player to play songs in Background in order (Random/Sequential) chosen by the user.
   Moreover, It generates a balloon notification in bottom Right corner of the screen, whenever a new song starts playing and continues to do that until manually stopped or it completes playing all songs.
.PARAMETER Filter
    String that can be used to filter out songs in a directory. Wildcards are allowed.
.PARAMETER Path
    Path of the Music Directory. 
    like,  Music -Path C:\music\ 
.PARAMETER Shuffle
    Switch to play music in shuffle mode, default value is 'sequential'
.PARAMETER Loop
    Switch to continuously play songs in a infinite Loop.
.PARAMETER Stop
Switch to kill any instance of Music playing in backgroud.
.PARAMETER ShowPlaylist

.EXAMPLE
    PS Root\> Music F:\Data\Music\MyPlaylist
    Count TotalPlayDuration Mode      
    ---------- ----------------- ----      
    26         98 Mins           Sequential 
   
    Example shows how to run music from by passing a music directory to the Function
.EXAMPLE
    PS Root\> Music -Verbose
    [VERBOSE] You've not provided a music directory, looking for cached information from Previous use.
    [VERBOSE] Starting a background Job to play Music files
    
    Count TotalPlayDuration Mode      
    ---------- ----------------- ----      
    26         98 Mins           Sequential

    Example shows that in case you don't provide a music directory, the function Looks for the cached information of the diretory from previous us of the function.
    Moreover, It displays the information like Count, Total play duration, and Mode chosen by the user
.EXAMPLE
    PS Root\> Music -Shuffle
    Count TotalPlayDuration Mode   
    ---------- ----------------- ----   
    26         98 Mins           Shuffle

    Choose 'Shuffle' switch to play music in shuffle mode, default value is 'sequential'.
.EXAMPLE
    PS Root\> Music -Shuffle -Loop
    Count TotalPlayDuration Mode           
    ---------- ----------------- ----           
    26         Infinite     Shuffle in Loop 

    Choose 'Loop' switch inorder to continuously play songs in a infinite Loop.
.EXAMPLE
    PS Root\> Music -Stop -Verbose
    [VERBOSE] Stoping any Already running instance of Media in background.

    When 'Stop' switch is used any instance Music playing in backgroud stops.
.NOTES
   Author  : Prateek Singh
   Twitter : @SinghPrateik
   Blog    : RidiCurious.com
#>
Function Invoke-MusicPlayer {
    [cmdletbinding()]
    Param(
        [Alias('P')]  [String] $Path,
        [Alias('F')]  [String] $Filter,
        [Alias('Sh')] [switch] $Shuffle,
        [Alias('St')] [Switch] $Stop,
        [Alias('L')]  [Switch] $Loop,
        [Alias('Pl')] [switch] $ShowPlaylist
    )

    $DefaultPath = "$env:TEMP\MusicPlayer.txt"
    If ($Stop) {
        Write-Verbose "Stoping any Already running instance of Media in background."
        Get-Job MusicPlayer -ErrorAction SilentlyContinue | Remove-Job -Force
    }
    Else {       
        #Caches Path for next time in case you don't enter path to the music directory
        If ($path) {
            $Path | out-file $DefaultPath
        }
        else {
            If ((Get-Content $DefaultPath -ErrorAction SilentlyContinue).Length -ne 0) {
                Write-Verbose "You've not provided a music directory. Looking for cached information from previous execution of this cmdlet"
                $path = Get-Content $DefaultPath

                If (-not (Test-Path $Path)) {
                    Write-Warning "Please provide a path to a music directory.`nFound a cached directory `"$Path`" from previous use, but that too isn't accessible!"
                    # Mark Path as Empty string, If Cached path doesn't exist
                    $Path = ''     
                }
            }
            else {
                Write-Warning "Please provide a path to a music directory."
            }
        }
 
        #initialization Script for back ground job
        $init = {
            # Function to calculate duration of song in Seconds
            Function Get-SongDuration($FullName) {
                $Shell = New-Object -COMObject Shell.Application
                $Folder = $shell.Namespace($(Split-Path $FullName))
                $File = $Folder.ParseName($(Split-Path $FullName -Leaf))
	                    
                [int]$h, [int]$m, [int]$s = ($Folder.GetDetailsOf($File, 27)).split(":")
	                    
                $h * 60 * 60 + $m * 60 + $s
            }

            # Converts seconds to HH:mm:ss string format
            Function ConvertTo-HHmmss($Seconds) {
                $Time = New-TimeSpan -Seconds $Seconds
                "{0:D2}:{1:D2}:{2:D2}" -f $Time.Hours, $Time.Minutes, $Time.Seconds
            }
                    
            # Function to Notify Information balloon message in system Tray
            Function Show-NotifyBalloon($Message) {
                [system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null            
                $Global:Balloon = New-Object System.Windows.Forms.NotifyIcon            
                $Balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -id $pid | Select-Object -ExpandProperty Path))                    
                $Balloon.BalloonTipIcon = 'Info'           
                $Balloon.BalloonTipText = $Message            
                $Balloon.BalloonTipTitle = 'Now Playing'            
                $Balloon.Visible = $true            
                $Balloon.ShowBalloonTip(1000)
            }
                    
            Function PlayMusic($Path, $Shuffle, $Loop, $Filter, $ShowPlaylist) {
                # Calling required assembly
                Add-Type -AssemblyName PresentationCore
    
                # Instantiate Media Player Class
                $MediaPlayer = New-Object System.Windows.Media.Mediaplayer
                        
                # Crunching the numbers and Information
                $FileList = Get-ChildItem $Path -Recurse -Filter $Filter -Include *.mp* | Select-Object fullname, @{n = 'Duration'; e = {get-songduration $_.fullname}}
                $FileCount = ($FileList | Measure-Object).count
                $TotalPlayDuration = [Math]::Round(($FileList.duration | Measure-Object -Sum).sum / 60)
                        
                # Condition to identifed the Mode chosed by the user
                if ($Shuffle) {
                    $Mode = "Shuffle"
                    $FileList = $FileList | Sort-Object {Get-Random}  # Find the target Music Files and sort them Randomly
                }
                Else {
                    $Mode = "Sequential"
                }
                        
                # Check If user chose to play songs in Loop
                If ($Loop) {
                    $Mode = $Mode + " in Loop"
                    $TotalPlayDuration = "Infinite"
                }
                        
                If ($FileList) {
                    If ($FileCount -gt 1) {
                        $Current = Split-Path $FileList[0].fullname -Leaf
                        $Next = Split-Path $FileList[1].fullname -Leaf
                    }
                    ElseIf ($FileCount -eq 1) {
                        $Current = Split-Path $FileList.fullname -Leaf
                        $Next = $null
                        $FileCount = '1'
                    }

                    [PSCustomObject] @{
                        Directory    = $Path
                        Count        = $FileCount
                        'PlayDuration(in mins)' = [String]$TotalPlayDuration
                        Mode         = $Mode
                        Current      = $Current
                        Next         = $Next
                        Playlist     = $FileList | Foreach-Object { [PSCustomObject] @{ FullName = $_.FullName; Duration = $(ConvertTo-HHmmss $_.duration) } }
                    }
                }
                else {
                    Throw "No music files found in directory:`"$path`" that matches Filter: $Filter ." 
                }
                        
                Do {
                    $FileList |ForEach-Object {
                        $CurrentFile = $(Split-Path $_.fullname -Leaf)
                        $Message = "File: {0} `nPlayDuration: {1}`nMode: {2}" -f $CurrentFile, $(ConvertTo-HHmmss (Get-SongDuration $_.fullname)), $Mode            
                        $MediaPlayer.Open($_.FullName)					# 1. Open Music file with media player
                        $MediaPlayer.Play()								# 2. Play the Music File
                        Show-NotifyBalloon ($Message)                   # 3. Show a notification balloon in system tray
                        Start-Sleep -Seconds $_.duration                # 4. Pause the script execution until song completes
                        $MediaPlayer.Stop()                             # 5. Stop the Song
                        $Balloon.Dispose(); $Balloon.visible = $false                           
                    }
                }While ($Loop) # Play Infinitely If 'Loop' is chosen by user
            }
        }

        # Removes any already running Job, and start a new job, that looks like changing the track
        If ($(Get-Job Musicplayer -ErrorAction SilentlyContinue)) {
            Get-Job MusicPlayer -ErrorAction SilentlyContinue |Remove-Job -Force
        }

        # Run only if path was Defined or retrieved from cached information
        If ($Path) {
            Write-Verbose "Starting a background Job to play audio files"
            Start-Job -Name MusicPlayer -InitializationScript $init -ScriptBlock {playmusic $args[0] $args[1] $args[2] $args[3] $args[4]} -ArgumentList $path, $Shuffle, $Loop, $Filter, $ShowPlaylist | Out-Null
            Start-Sleep -Seconds 3       # Sleep to allow media player some breathing time to load files
            $Results = Receive-Job -Name MusicPlayer 
            $Results | Select-Object Directory, Count, 'PlayDuration(in mins)', Mode, Current, Next
            If ($ShowPlaylist){
                $Results.Playlist | Format-Table -AutoSize
            }
        }
    }
    
}

Set-Alias -Name Music -Value Invoke-MusicPlayer
Set-Alias -Name Play -Value Invoke-MusicPlayer


# Exporting the members and their aliases
Export-ModuleMember -Function "Invoke-MusicPlayer" -Alias *
