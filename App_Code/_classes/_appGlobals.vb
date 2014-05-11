Imports Microsoft.VisualBasic
Imports System
Imports System.Configuration

Public Module _appGloBals
		
	Public currentURL as string
	Public setExpirerDate as date
	Public appPageID, appModuleID, appContentID as Integer
	
	'***********************************************************************************************************
	'* description
	'***********************************************************************************************************
	Public Class _NPMarkupPage
		Inherits Page
		Public Overloads Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
		End Sub
	End Class
	
End Module