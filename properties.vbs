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

			'used to add a bar for each digit in the value
			'after the value printed to console
			'to display a more graphical representation of the value
			digit1 = 0
			digit2 = 0
			digit3 = 0
			bar1 = ""
			bar2 = ""
			bar3 = ""
			bar = vbTab
			
			If(Len(CStr(aData(cnValue))) > 2) Then
				digit1string = Mid(Trim(CStr(aData(cnValue))), 1, 1)
				digit1 = CInt(digit1string)
				For i = 0 To digit1
					bar1 = bar1 + "|"
				Next
				bar1 = Right(bar1, Len(bar1)-1)
				bar = bar + bar1
				
				digit2string = Mid(Trim(CStr(aData(cnValue))), 2, 1)
				digit2 = CInt(digit2string)
				For i = 0 To digit2
					bar2 = bar2 + "/"
				Next
				bar2 = Right(bar2, Len(bar2)-1)
				bar = bar + bar2

				digit3string = Mid(Trim(CStr(aData(cnValue))), 3, 1)
				digit3 = CInt(digit3string)
				For i = 0 To digit3
					bar3 = bar3 + "-"
				Next
				bar3 = Right(bar3, Len(bar3)-1)
				bar = bar + bar3
				
				If(aData(cnValue) > 999) Then
					bar = bar + "*"
				End If
			ElseIf(Len(CStr(aData(cnValue))) > 1) Then
				digit1string = Mid(Trim(CStr(aData(cnValue))), 1, 1)
				digit1 = CInt(digit1string)
				For i = 0 To digit1
					bar1 = bar1 + "/"
				Next
				bar1 = Right(bar1, Len(bar1)-1)
				bar = bar + bar1
				
				digit2string = Mid(Trim(CStr(aData(cnValue))), 2, 1)
				digit2 = CInt(digit2string)
				For i = 0 To digit2
					bar2 = bar2 + "-"
				Next
				bar2 = Right(bar2, Len(bar2)-1)
				bar = bar + bar2
			ElseIf(Len(CStr(aData(cnValue))) > 0) Then
				digit1string = Mid(Trim(CStr(aData(cnValue))), 1, 1)
				digit1 = CInt(digit1string)
				For i = 0 To digit1
					bar1 = bar1 + "-"
				Next
				bar1 = Right(bar1, Len(bar1)-1)
				bar = bar + bar1
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