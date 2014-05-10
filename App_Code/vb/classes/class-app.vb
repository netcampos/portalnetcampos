Imports Microsoft.VisualBasic
Imports System
			
Public Class _npApp
	
	Private _lib as new _npLibrary()
	Private _db as new _npDbViews()
			
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Public Sub New()			
		MyBase.New()
	End Sub
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property domainName() As String
		Get
			Return HttpContext.Current.Request.Url.Scheme & "://" & _db.getCidade("domain")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property currentURL() As String
		Get
			Return domainName & _lib.QS("Slug")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property currentSlug() As String
		Get
			Return _lib.QS("Slug")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property idCidade() As Integer
		Get
			Return _lib.QS("idCidade")
		End Get
		Set(ByVal value As Integer)
		
		End Set
	End Property
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Public sub setPageHttpHeader()
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
		Dim ts As New TimeSpan(0,60,0)
		HttpContext.Current.Response.Cache.SetMaxAge(ts)
		HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddSeconds(3600))
		HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
		HttpContext.Current.Response.Cache.SetValidUntilExpires(true)
		HttpContext.Current.Response.Cache.VaryByHeaders("Accept-Language") = true
		HttpContext.Current.Response.Cache.VaryByHeaders("User-Agent") = true
		HttpContext.Current.Response.Cache.SetLastModified(DateTime.Now.AddHours(-2))
		HttpContext.Current.Response.AppendHeader("X-UA-Compatible", "IE=edge,chrome=1")
		HttpContext.Current.Response.Headers.Remove("Server")
		HttpContext.Current.Response.AppendHeader("Server", "NetCampos/1.0")
		HttpContext.Current.Response.AppendHeader("X-XSS-Protection", "1; mode=block")
	End sub	
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Public sub setHttpHeader()
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
		Dim ts As New TimeSpan(0,60,0)
		HttpContext.Current.Response.Cache.SetMaxAge(ts)
		'HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddSeconds(3600))
		'HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
		'HttpContext.Current.Response.Cache.SetValidUntilExpires(true)
		'HttpContext.Current.Response.Cache.VaryByHeaders("Accept-Language") = true
		'HttpContext.Current.Response.Cache.VaryByHeaders("User-Agent") = true
		'HttpContext.Current.Response.Cache.SetLastModified(DateTime.Now.AddHours(-3))
		HttpContext.Current.Response.AppendHeader("X-UA-Compatible", "IE=8")
		HttpContext.Current.Response.AppendHeader("Server", "NetCampos/1.0")
	End sub		
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Public sub render()

	End sub	

End Class