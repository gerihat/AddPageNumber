' AddPageNumber.vbs script
' PDFCreator Custom Script.
' Version: 1.1.0.0
' Date: June, 1. 2020
' Author: Miguel A. Camacho
' Comment: This script adds page number to a given inf spool file using HTA application
' This file must be linked or incrustated in HTA application

Option Explicit

Const AppTitle ="PDFCreator - AddPageNumbers"

Dim pdf, pdfText, WshShell
Dim objArgs
Dim mFname_in, mFname_out
Dim FSO
Dim WshProcEnv, system_architecture
Dim mFromPage, mToPage, mStartPageNumber, mNumberOfPages, mPageNumberPosition
Dim mBorderXMillimeter, mBorderYMillimeter

Sub CenterWindow( widthX, heightY )
    self.ResizeTo widthX, heightY 
    self.MoveTo (screen.Width - widthX)/2, (screen.Height - heightY)/2
End Sub

Sub AddPageNumber()
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")

	Set pdf = CreateObject("pdfforge.pdf.pdf")
	Set pdfText  = CreateObject("pdfforge.pdf.pdfText")

	mFname_in=document.getElementByID("fname_in").value
	if (mFname_in="") Then
		msgbox "Debe seleccionar un fichero",vbExclamation,"Fichero no encontrado"
		Exit Sub
	End If
	mFname_out=FSO.GetParentFolderName(mFname_in) & "\" & FSO.GetBasename(mFname_in) & "_pag.pdf"
	pdfText.Text = "[PAGE]/[PAGES]"
	pdfText.FontName = "Arial.ttf"
	pdfText.FontPath = WshShell.SpecialFolders("Fonts")
	pdfText.FontSize = 18

	mFromPage			= document.getElementByID("fromPage").value
	mToPage				= document.getElementByID("toPage").value
	mStartPageNumber	= document.getElementByID("startPageNumber").value
	mNumberOfPages		= document.getElementByID("numberOfPages").value
	mPageNumberPosition = document.getElementByID("pageNumberPosition").value
	mBorderXMillimeter 	= 10.0
	mBorderYMillimeter 	= 10.0

	pdf.AddPageNumberToPDFFile mFname_in, mFname_out, mFromPage, mToPage, mStartPageNumber, mNumberOfPages, mPageNumberPosition, mBorderXMillimeter, mBorderYMillimeter, (pdfText)
	WshShell.Run(chr(34) & mFname_out & chr(34))
	
	pgBar "end"

	Set pdfText = Nothing
	Set pdf = Nothing
	Set WshShell=Nothing
	Set WshProcEnv=Nothing
	Set FSO=Nothing

End Sub


Sub start()
	pgBar "begin"
	setTimeout "AddPageNumber()", 0
End Sub

Sub pgBar(strCommand)
	select case strCommand
		case "begin"
			progressbar.innerHTML="Procesando..."
		case "end"
			progressbar.style.color="MediumSeaGreen"
			progressbar.innerHTML="Finalizado"		
	end select
	
End Sub







