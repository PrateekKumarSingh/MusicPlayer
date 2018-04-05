# Powershell Music Player
Invoke-MusicPlayer automates Windows Media Player, to play songs in Background with a set of user preferences available as a 'switch' in the cmdlet.
[Full Article](http://ridicurious.com/2018/04/03/powershell-module-to-play-music/)

## Features and Benefits
*  **Plays all audio in background on a hidden Media player instance**
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Default.jpg">

*  **Filter Audio Files**
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Filter.jpg">

*  **Randomly shuffles your play list**
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Shuffle.jpg">

*  **Runs your playlist in a never ending loop**
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/ShuffleAndLoop.jpg">

*  **Stop the playing audio on demand**
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Stop.jpg">

*  **Stores\Caches last accessed directory**
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Cached.jpg">

*  **Displays information and user preferences**
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Output.jpg">

*  **Popup a balloon notification** in bottom Right corner of the screen, whenever a new song starts playing and continues to do that until manually stopped or it completes playing all songs.
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Balloon.jpg">

* ** Support Switch aliases**
    <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/SwitchAlias.jpg">


 Installation
 -
 #### [PowerShell V5](https://www.microsoft.com/en-us/download/details.aspx?id=50395) and Later
 You can install the `MusicPlayer` module directly from the PowerShell Gallery

 * [Recommended] Install to your personal PowerShell Modules folder
 ```PowerShell
 Install-Module MusicPlayer -scope CurrentUser
 ```

 <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Installation_v5.jpg">

 * [Requires Elevation] Install for Everyone (computer PowerShell Modules folder)
 ```PowerShell
 Install-Module MusicPlayer
 ```

 #### PowerShell V4 and Earlier
 To install to your personal modules folder run:

 ```PowerShell
 iex (new-object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/PrateekKumarSingh/MusicPlayer/master/Install.ps1')
 ```
 <br/><br/><img src="https://github.com/PrateekKumarSingh/MusicPlayer/blob/master/media/Installation_v4.jpg">

Help Information
-
Run below commands to see some examples
```PowerShell
Get-Help Invoke-MusicPlayer -Examples
```
