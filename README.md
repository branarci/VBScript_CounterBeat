# VBScript_CounterBeat
A short VBScript program that reads/writes a value from/to a file and prints it to the console. Constantly cycles from 0 to 1000 and back again in a never-ending loop.

'-------------------------------------------------------------------------
'FEATURES:
'Simulates a kind of breathing
'	slowly going from 0(full exhale) - 1000(full inhale) and 
'	back again in an unending loop
'Uses a property file to store the variables used so that interrupted
'	actions can be restarted
'-------------------------------------------------------------------------
'EDITS:
' - originally a simple counting program that increments the value
' -	originally used a boolean to swap counting directions but this seemed
'	to be prone to error
'		- switched to a simple binary int
' - if the properties.txt was interrupted during writing or the file was
'	modified while code is running the resulting file would be blank and
'	this caused the script to not work
'		- added an if clause to delete the file if not properly formatted
'			- this required the script to have to be re-run
'			- updated it to delete the file and continue to run
'-------------------------------------------------------------------------
