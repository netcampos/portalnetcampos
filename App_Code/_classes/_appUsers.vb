Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports MySql.Data
Imports MySql.Data.MySqlClient

Imports System.IO
Imports System.Xml

Public Class _appUsers
	
	#Region "INIT"
		
		Private _lib as new _appLibrary()
		Private _db as new _appDB()
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:
		'******************************************************************************************************************
		Public Sub New()
			MyBase.New()
		End Sub
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub
		
	#End Region
	
	#Region "Commom Functions"
	
		'******************************************************************************************************************
		'' @SDESCRIPTION:	writes a string to the output in the same line
		'' @PARAM:			value [string]: output string
		'******************************************************************************************************************
		Public Function userON() as Boolean
			dim retorno as boolean = false
			retorno = HttpContext.Current.User.Identity.IsAuthenticated
			return retorno
		end Function
		
	#End Region	
			
End Class