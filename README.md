# Initialize-WinTerminal
Sets up the Windows Terminal how I like it.

```powershell
 iwr https://raw.githubusercontent.com/JamesKehr/Initialize-WinTerminal/main/launcher.ps1 -UseBasicParsing | iex
```

launcher.ps1 downloads the required files to the present working directory ($PWD), then executes Initialize-WindowsTerminal.ps1.

The script requires internet connectivity to download all the applications and their pre-requisites.
