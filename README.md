# Install Fonts using Powershell


## How it works:

This script installs fonts in the "FontFiles" Folder. Checks if fonts already exist by using md5 hash keys,
checking font families and checking if the font files exist in the C:\Windows\Fonts folder.

## How to use it:
- Delete the dummyFile.txt in the FontFiles folder
- Copy & Paste your fonts to the FontFiles folder in the script directory. This script works for all types of fonts.
- Run the script.
- You can check log file for results in C:\Windows\Logs\fontInstallPVH.log

### Log Legend:
The log file is divided into 4 sections.

- Meta Data from fonts in the Font files folder
- Meta Data from fonts in the C:\Windows\Fonts folder
	-	Font Meta Data:
		- Font name
		- File Name
		- MD5 Hash key
		- Manipulated font name string with all special characters removed
		- File Path
- Installing process: States if a new font was installed or if the font already exists.
- Test Fonts: States if the script ran successfully, if it didn't. States the names of the fonts that did not install.



