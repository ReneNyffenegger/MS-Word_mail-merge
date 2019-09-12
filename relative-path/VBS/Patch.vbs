' Patch DOCX to find its database in its folder instead of absolute path
' Written by Raymai97 (MaiSoft) on 14 March 2015

Set Shell=CreateObject("WScript.Shell")
Set FS = CreateObject("Scripting.FileSystemObject")
Set Stream = CreateObject("ADODB.Stream")

Dim FilePath, MyDir, TempDir, Text
MyDir = DirPathOf(WScript.ScriptFullName)

' If no argument, let user select file
If WScript.Arguments.Count < 1 Then
	Set oExec=Shell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
	FilePath = oExec.StdOut.ReadLine
	If FilePath = "" Then GTFO
Else
	FilePath = WScript.Arguments(0)
End If

' Check if it's DOCX and look for 7za.exe
If FS.FileExists(FilePath)=False or FileExtOf(FilePath)<>"docx" Then Hey "Only DOCX file is supported!" : GTFO
If FS.FileExists(MyDir&"\7za.exe")=False Then Hey "7za.exe not found! Please download it from 7-zip.org!" : GTFO
' Ask for confirmation
If Shell.Popup("Patch this DOCX file so it finds database in its folder?" & VbCrLf & VbCrLf & FilePath, , "Are you sure?", &H4 + &H20 + &H40000) = 7 Then GTFO
' Does 7z agree this DOCX is vaild?
Shell.CurrentDirectory = MyDir
If Shell.Run("7za.exe t """ & FilePath & """",0,True) <> 0 Then Hey "This file doesn't look alright! Are you sure it is a vaild DOCX file?" : GTFO
' Extract to a temp folder
TempDir = FS.GetSpecialFolder(2) & "\" & FS.GetTempName
Shell.Run "7za.exe x """ & FilePath & """ -o""" & TempDir & """",0,True
' Does it use Data Source?
If FS.FileExists(TempDir & "\word\_rels\settings.xml.rels")=False Then
	Hmm "This file doesn't use Data Source, so no action will be taken."
Else
	' Open \word\settings.xml
	Stream.Charset = "utf-8"
	Stream.Open
	Stream.LoadFromFile TempDir & "\word\settings.xml"
	Text = Stream.ReadText
	Stream.Close
	' find 'Data Source=' till ';'
	i = 1
	Do
		i = InStr(i, Text, "Data Source=")
		If Not i > 0 Then Exit Do
		i = i + 12
		j = InStr(i, Text, ";")
		If Not j > 0 Then Exit Do
		DBName = Mid(Text,i,j-i)
		Text = Replace(Text,DBName,FileNameOf(DBName))
		i=j
	Loop
	' save
	Stream.Open
	Stream.WriteText Text
	Stream.SaveToFile TempDir & "\word\settings.xml", 2
	Stream.Close
	' Open \word\_rels\settings.xml.rels
	Stream.Open
	Stream.LoadFromFile TempDir & "\word\_rels\settings.xml.rels"
	Text = Stream.ReadText
	Stream.Close
	' find 'Target="' till '"'
	i = 1
	Do
		i = InStr(i, Text, "Target=""")
		If Not i > 0 Then Exit Do
		i = i + 8
		j = InStr(i, Text, """")
		If Not j > 0 Then Exit Do
		DBName = Mid(Text,i,j-i)
		Text = Replace(Text,DBName,FileNameOf(DBName))
		i=j
	Loop
	' save
	Stream.Open
	Stream.WriteText Text
	Stream.SaveToFile TempDir & "\word\_rels\settings.xml.rels", 2
	Stream.Close
	' Update the docx
	Ret = Shell.Run("7za.exe u """ & FilePath & """ """ & TempDir & "\*""",0,True)
	If Ret = 0 Then
		Hmm "Done! Have a nice day! :D"
	Else
		Hey "No good! 7za.exe returned an error code: " & Ret
	End If
End If
Shell.Run "cmd /c rd /s /q """ & TempDir & """",0,True

Function DirPathOf(ByVal Str)
i = InStrRev(Str, "\")
If i > 0 Then Str = Mid(Str, 1, i - 1)
DirPathOf = Replace(Str, """", "")
End Function

Function FileExtOf(ByVal Str)
i = InStrRev(Str, ".")
If i > 0 Then Str = Mid(Str, i + 1)
FileExtOf = Replace(Str, """", "")
End Function

Function FileNameOf(ByVal Str)
i = InStrRev(Str, "\")
If i > 0 Then
	FileNameOf = Mid(Str, i + 1)
Else
	FileNameOf = Str
End If
End Function

Sub GTFO()
WScript.Quit
End Sub

Sub Hey(Msg)
Shell.Popup Msg, , "Hey!", &H10 + &H40000
End Sub

Sub Hmm(Msg)
Shell.Popup Msg, , "Patch Info", &H40 + &H40000
End Sub