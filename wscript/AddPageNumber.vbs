' AddPageNumber.vbs script
' PDFCreator Custom Script.
' Version: 1.0.0.0
' Date: March, 12. 2020
' Author: Miguel A. Camacho
' Comment: This script adds page number to a given inf spool file.

Option Explicit

'On Error Resume Next
Const AppTitle ="PDFCreator - AddPageNumbers"

Dim pdf, pdfText, WshShell
Dim objArgs
Dim fname_in,fname_out
Dim FSO, pathGetCurrentFolder
Dim fromPage, toPage, startPageNumber, numberOfPages, pageNumberPosition
Dim borderXMillimeter, borderYMillimeter

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")


'Main enty point
AddPageNumber()



'Add Page Number Subroutine using WScript
Sub AddPageNumber()

	'Running x32 WScript in x64 system (SysWOW64)
	if getSystemArchitecture()="AMD64" Then
		If InStr(LCase(WScript.FullName),"syswow64")=0 Then
	  		'CreateObject("WScript.Shell").Run """C:\Windows\SysWOW64\wscript.exe"" """ & WScript.ScriptFullName & """" <-- sin parámetros
			CreateObject("Wscript.Shell").Run "C:\Windows\SysWOW64\wscript.exe " & _
				chr(34) & WScript.ScriptFullName & chr(34) & " " & chr(34) & WScript.Arguments(0) & chr(34)
			WScript.Quit
		End If
	End If


	Set pdf = WScript.CreateObject("pdfforge.pdf.pdf")
	Set pdfText  = Wscript.CreateObject("pdfforge.pdf.pdfText")
	Set objArgs=WScript.Arguments

	fname_in=objArgs(0)
	fname_out=fso.GetParentFolderName(fname_in) & "\" & fso.GetBasename(fname_in) & "_pag.pdf"

	pdfText.Text = "[PAGE]/[PAGES]"
	'pdfText.FontColorRed = 200
	pdfText.FontName = "Arial.ttf"
	pdfText.FontPath = WshShell.SpecialFolders("Fonts")
	pdfText.FontSize = 18

	fromPage			= 1
	toPage				= 0
	startPageNumber		= 1
	numberOfPages		= 0
	pageNumberPosition 	= 3
	borderXMillimeter 	= 10.0
	borderYMillimeter 	= 10.0

	pdf.AddPageNumberToPDFFile fname_in, fname_out, fromPage, toPage, startPageNumber, numberOfPages, pageNumberPosition, borderXMillimeter, borderYMillimeter, (pdfText)

	If Err.Number<>0 Then
		msgbox Err.Description,vbCritical,"Error " & Err.Number
		Err.clear
	End If

	Set pdfText = Nothing
	Set pdf = Nothing
	Set WshShell = Nothing
	Set fso = Nothing
End Sub

'Get system architecture
Function getSystemArchitecture()
Dim WshProcEnv,process_architecture,system_architecture

	if IsEmpty(WshProcEnv) Then
		Set WshProcEnv=WshShell.Environment("Process")

		process_architecture= WshProcEnv("PROCESSOR_ARCHITECTURE") 
		If process_architecture = "x86" Then    
		    system_architecture= WshProcEnv("PROCESSOR_ARCHITEW6432")
		    If system_architecture = ""  Then    
		        system_architecture = "x86"
		    End if    
		Else    
		    system_architecture = process_architecture    
		End If
	End If
	getSystemArchitecture=system_architecture
End Function