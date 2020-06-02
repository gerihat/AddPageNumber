' AddPageNumber.vbs script
' PDFCreator Custom Script.
' Version: 1.0.0.0
' Date: March, 12. 2020
' Author: Miguel A. Camacho
' Comment: This script adds page number to a given inf spool file.

Option Explicit

Const AppTitle ="PDFCreator - AddPageNumbers"

Dim pdf, pdfText, WshShell
Dim objArgs
Dim fname_in,fname_out
Dim FSO
Dim WshProcEnv, system_architecture

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshProcEnv=WshShell.Environment("Process")
system_architecture=WshProcEnv("PROCESSOR_ARCHITEW6432")

'Running x64 WScript in x64 system (system32)
If system_architecture="AMD64" Then
	MsgBox "Script for x32 system version only.",48
	Wscript.Quit
End If

Set pdf = WScript.CreateObject("pdfforge.pdf.pdf")
Set pdfText  = Wscript.CreateObject("pdfforge.pdf.pdfText")
Set objArgs=WScript.Arguments

fname_in=objArgs(0)
fname_out=fso.GetParentFolderName(fname_in) & "\output_pag.pdf"

pdfText.Text = "[PAGE]/[PAGES]"
'pdfText.FontColorRed = 200
pdfText.FontName = "Arial.ttf"
pdfText.FontPath = WshShell.SpecialFolders("Fonts")
pdfText.FontSize = 18
pdf.AddPageNumberToPDFFile fname_in, fname_out, 1, 0, 1, 0, 3, 10.0, 10.0, (pdfText)


Set pdfText = Nothing
Set pdf = Nothing
Set WshShell = Nothing
Set fso = Nothing