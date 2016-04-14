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

'Global values for the properties file and the variable's Array position
cnFileName = "properties.txt"
cnValue = 1
cnDirection = 2

'create a file system object to enable IO operations
Dim FSo : Set FSo = CreateObject("Scripting.FileSystemObject")

'constantly loop
Do While True
	'check if properties file exists
	If FSo.FileExists(cnFileName) Then
		'check if properties file is properly formatted
		'if not then delete
		If(FSo.OpenTextFile(cnFileName).AtEndOfStream = True) Then
			FSo.DeleteFile(cnFileName)
		'if useable then get variables and modify them
		Else
			'load variables into array from file
			aData = Split(FSo.OpenTextFile(cnFileName).ReadAll(), vbCrLf)

			'find the direction value
			'increment or decrement value based on direction
			If(aData(cnDirection) = 1) Then
				aData(cnValue) = aData(cnValue) + 1
			ElseIf(aData(cnDirection) = 0) Then
				aData(cnValue) = aData(cnValue) - 1
			End If

			'if the value reaches limits
			'reverse the direction of counting
			If (aData(cnValue) > 1000) Then
				aData(cnDirection) = 0
			ElseIf (aData(cnValue) < 1) Then
				aData(cnDirection) = 1
			End If

			'placeholder for the result of all digits with a tab at the beginning
			bar = vbTab
			
			'if 3 or more digits
			If(Len(CStr(aData(cnValue))) > 2) Then
				'aData - get the value from the aData array
				bar = updateBarDetails(aData(cnValue), 1, "|", bar)
				bar = updateBarDetails(aData(cnValue), 2, "/", bar)
				bar = updateBarDetails(aData(cnValue), 3, "-", bar)
			'if 2 or more digits
			ElseIf(Len(CStr(aData(cnValue))) > 1) Then
				bar = updateBarDetails(aData(cnValue), 1, "/", bar)
				bar = updateBarDetails(aData(cnValue), 2, "-", bar)
			'if 1 or more digits
			ElseIf(Len(CStr(aData(cnValue))) > 0) Then
				bar = updateBarDetails(aData(cnValue), 1, "-", bar)
			'if anything else or something wrong
			Else
				bar = "***"
			End If
									
			'print the value to the console
			WScript.Echo aData(cnValue) & bar

			'save changes to file
			FSo.CreateTextFile(cnFileName, True).Write Join(aData, vbCrLf)
		End If
	'if properties file doesn't exist
	Else
		'create an array with default variables
		Dim aData : aData = Array("properties.txt", 1, 1)

		'save defaults to file
		FSo.CreateTextFile(cnFileName, True).Write Join(aData, vbCrLf)
	End If
Loop


Function updateBarDetails(value, position, barValue, barRemaining)
	'multiple functions used
	'CStr - convert the result into a string
	'Trim - trim the string
	'	- probably not necessary but had error at one point that this solved
	'Mid - get a character in the string at a specific position
	digitString = Mid(Trim(CStr(value)), position, 1)
	'convert back into integer
	digit = CInt(digitString)
	'constructs a string of the relevant amount of bars
	For i = 0 To digit
		barTotal = barTotal + barValue
	Next
	'removes the first character in the string
	'	- had an error and way of fixing it was to 
	'	- make i start at 0 instead of 1
	'	- so this removes the extra bar from 0
	barTotal = Right(barTotal, Len(barTotal)-1)
	'adds the tab to the beginning of the bars
	barTotal = barRemaining + barTotal
	'return result
	updateBarDetails = barTotal
End Function
