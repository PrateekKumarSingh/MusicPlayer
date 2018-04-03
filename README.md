# Powershell Music Player
Invoke-MusicPlayer automates Windows Media Player, to play songs in Background with a set of user preferences available as a 'switch' in the cmdlet.

## Features and Benefits
*  Plays all audio in background on a hidden Media player instance
*  Randomly shuffles your play list
*  Runs your playlist in a never ending loop
*  Stop the playing audio on demand
*  Stores\Caches last accessed directory
*  Displays information and user preferences
*  Popup a balloon no tification in bottom Right corner of the screen, whenever a new song starts playing and continues to do that until manually stopped or it completes playing all songs.

 Installation
 -
 #### [PowerShell V5](https://www.microsoft.com/en-us/download/details.aspx?id=50395) and Later
 You can install the `MusciPlayer` module directly from the PowerShell Gallery

 * [Recommended] Install to your personal PowerShell Modules folder
 ```PowerShell
 Install-Module MusicPlayer -scope CurrentUser
 ```

 ![](https://raw.githubusercontent.com/PrateekKumarSingh/Gridify/master/Images/Installation_v5.jpg)

 * [Requires Elevation] Install for Everyone (computer PowerShell Modules folder)
 ```PowerShell
 Install-Module MusicPlayer
 ```

 #### PowerShell V4 and Earlier
 To install to your personal modules folder run:

 ```PowerShell
 iex (new-object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/PrateekKumarSingh/MusicPlayer/master/Install.ps1')
 ```

Help Information
-
Run below commands to see some examples
```PowerShell
Get-Help Set-GridLayout -Examples
```



For example, in following image function `Show-Graph` takes data points as input and plot them on a 2D graph

![](https://github.com/PrateekKumarSingh/PSConsoleGraph/blob/master/Images/Example1.png)


You can also customize the labels on X and Y-Axis

![](https://github.com/PrateekKumarSingh/PSConsoleGraph/blob/master/Images/Example2.png)

The function consumes data points, generated during script execution or Pre stored data like in a file or database.

![](https://github.com/PrateekKumarSingh/PSConsoleGraph/blob/master/Images/Example3.png)
