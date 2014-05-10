Imports Microsoft.VisualBasic
Imports System
Imports System.Configuration

Public Module _appGloBals
		
	Public currentURL as string
	
	'***********************************************************************************************************
	'* description
	'***********************************************************************************************************
	Public Class _NPMarkupPage
		Inherits Page
		Public Overloads Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
		End Sub
	End Class
	
End Module