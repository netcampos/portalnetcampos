Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.IO
Imports System.Configuration
Imports System.Text
Imports System.Security.Cryptography
Imports System.Web
Imports System.Web.Configuration
Imports System.Web.Management
Imports System.Web.Security
Imports System.Web.Caching
Imports System.Globalization
Imports System.Collections.Generic
Imports TwitterVB2
Imports System.Dynamic

Public Class npOAuth
	
	Private str as new StringOperations()
	Private db as new database()
	Private user as new npUsers()
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Sub New()
		MyBase.New()	
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Sub connectTwitter()
		Dim urlRetorno As String = db.getOAuth(2,"urlRetorno")
		Dim ConsumerKey As String = db.getOAuth(2,"appID")
		Dim ConsumerKeySecret As String = db.getOAuth(2,"secretKey")
		Dim tw As New TwitterAPI
		_lib.responseRedirect(tw.GetAuthorizationLink(ConsumerKey, ConsumerKeySecret, urlRetorno))
	end sub
	
End Class