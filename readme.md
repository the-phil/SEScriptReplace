# SE Script Replace
Replace scripts in saved games and SE style XML files.


### How to Use
Run from the command line
```$ cscript SEScriptReplace.vbs
You will need to know how to change directories or folders and be familiar with the command prompt. Most SE files you will modify will be in
```C:\Users\<YourUserName>\AppData\Roaming\SpaceEngineers
Example Change Directory to the above location ^
```$ cd\Users\<YourUserName\AppData\Roaming\SpaceEngineers
Use tab to auto complete, spaces and special characters don't play will with most CLI's.


### General Information (algorithm)
Opens all files recursively from where the VBS script is and searches for the "program element" in the file if it is XML based.
It will then check to see if the contents in that element have a match to the strReplaceProgramStartingWith variable below.
If there is a match for the string variable the data within element will be replaced with the contents of the file strNewScriptName defaulted to Script.cs, the same file that is extracted from the workshop sbc files.
Backup files are NOT made due to VBS default security perms so backup your files before running this!
A list of changes are added to the log file named in the strLogFile variable. The default log file is script-updates.log
This script can also be ran in the saved games folder to update the ships that currently exist, not sure if they auto update based on the mod.
This script will search for existing platformID variables for the IOS ATI SE in game script. As long as the default is "Exmple" it will be replaced.

### Additional Info
This script was written to update the exploration enhancement mod
[Exploration Enhancement Mod](http://steamcommunity.com/sharedfiles/filedetails/?id=531659576)
with the IOS ATI script
[$IOS <ATI> (Automated Trading Interface)](https://steamcommunity.com/sharedfiles/filedetails/?id=539692861)

I'm sure it can work and be modified to work for other scripts.


### Copyright (C) 2016  Philip Allen
This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public License for more details.
You should have received a copy of the GNU Lesser General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.