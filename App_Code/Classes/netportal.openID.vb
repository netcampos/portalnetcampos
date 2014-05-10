Imports Microsoft.VisualBasic
Imports System
Imports System.Net
Imports System.XML
Imports System.Type
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
Imports TwitterVB2
Public Class _openID
	
	'Public Members
	Public idCidade, idServico, idUsuario, ErrorReturn, SuccessReturn as String
	
	'Private Members
	Private appId, apiKey, secretKey, urlRetorno as String
	
	'******************************************************************************************************************
	'' Constructor
	'******************************************************************************************************************	
	Public Sub New()
		MyBase.New()
	End Sub
	
	'******************************************************************************************************************
	'' Destructor
	'******************************************************************************************************************		
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub	
	
	Public Sub _init()
		appId = city._OAuth(idCidade, idServico, "appId")
		secretKey = city._OAuth(idCidade, idServico, "secretKey")
		apiKey = city._OAuth(idCidade, idServico, "apiKey")	
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Sub _redirectAPP()
		'urlRetorno = city.dados(idcidade,"urlCidade") & "/OAuth2.0/" & city._OAuth(idCidade, idServico, "urlRetorno")
		'urlRetorno = "http://marina.gruponetcampos.corp/OAuth2.0/" & city._OAuth(idCidade, idServico, "urlRetorno")
		'redirectServico()
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Private Sub redirectServico()
		Dim urlServico as String = String.Empty
		
		if idServico = 1 then  'facebook
			'urlRetorno = urlRetorno & "/?ukey=" & idUsuario & "|" & idServico & "|" & idCidade
			'urlServico = "https://graph.facebook.com/oauth/authorize?client_id=" & appId & "&scope=offline_access,read_stream,publish_stream,email&display=popup&redirect_uri=" & urlRetorno
			'HttpContext.Current.response.redirect(urlServico, true)
			Exit Sub			
		End If
		
		If idServico = 2 then 'twitter
			'urlRetorno = urlRetorno & "/?ukey=" & idUsuario & "&idServico=" & idServico & "&idCidade=" & idCidade
			'Dim ConsumerKey As String = appId
			'Dim ConsumerKeySecret As String = secretKey
			'Dim tw As New TwitterAPI
			'HttpContext.Current.response.Redirect(tw.GetAuthorizationLink(ConsumerKey, ConsumerKeySecret, urlRetorno))
			Exit Sub
		End If
		
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Sub _loginAPP()
		urlRetorno = city.dados(idcidade,"urlCidade") & "/openID/" & city._OAuth(idCidade, idServico, "urlRetorno")
		'urlRetorno = "http://marina.gruponetcampos.corp/OAuth2.0/" & city._OAuth(idCidade, idServico, "urlRetorno")
		redirectLogin()
	End Sub	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Private Sub redirectLogin()
		Dim urlServico as String = String.Empty
		
		if idServico = 1 then  'facebook
			urlRetorno = urlRetorno & "/?citykey=" & idServico & "|" & idCidade
			urlServico = "https://graph.facebook.com/oauth/authorize?client_id=" & appId & "&scope=offline_access,read_stream,publish_stream,email&display=popup&redirect_uri=" & urlRetorno
			'HttpContext.Current.response.redirect(urlServico, true)
			Exit Sub
		End If
		
		If idServico = 2 then 'twitter
			urlRetorno = urlRetorno & "/?citykey=" & "&idServico=" & idServico & "&idCidade=" & idCidade
			Dim ConsumerKey As String = appId
			Dim ConsumerKeySecret As String = secretKey
			Dim tw As New TwitterAPI
			'HttpContext.Current.response.Redirect(tw.GetAuthorizationLink(ConsumerKey, ConsumerKeySecret, urlRetorno))
			Exit Sub
		End If
		
	End Sub	
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************			
	Private Function _loginFacebook() as Boolean

		
	End Function
		
End class