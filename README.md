Install Fonts using Powershell


How it works:

This script installs fonts in the "FontFiles" Folder. Checks if fonts already exist by using md5 hash keys,
checking font families and checking if the font files exist in the c:\Windows\Fonts folder.

Usage:
- Delete the dummyFile.txt in the FontFiles folder
- Add your fonts to the FontFiles folder in the script directory. This script works for all types of fonts.
- Run the script.
- You can check log file for results in "C:\Windows\Logs\fontInstall.log"


