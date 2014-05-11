Imports Microsoft.VisualBasic
Imports System.Configuration
Imports System.Configuration.ConfigurationManager
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.IO

Public Class _appLibrary
		
	#Region "INIT"
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************		
		Public Sub New()
			MyBase.New()
		End Sub
		
		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Property domainName() As String
			Get
				if ( QS("dev") = "1") then
					Return HttpContext.Current.Request.Url.Scheme & "://localhost"
				else
					Return HttpContext.Current.Request.Url.Scheme & "://www." & QS("domain") 
				end if
				
			End Get
			Set(ByVal value As String)
			
			End Set
		End Property
				
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Property domainAssets() As String
			Get
				Return System.Configuration.ConfigurationManager.AppSettings.Get("appAssets")
			End Get
			Set(ByVal value As String)
			
			End Set
		End Property
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Property templatePath() As String
			Get
				Return System.Configuration.ConfigurationManager.AppSettings.Get("templatePath")
			End Get
			Set(ByVal value As String)
			
			End Set
		End Property		
		
	#End Region
	
	#Region "Comom Functions"
			
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function LenB(ByVal str As String) As Integer
			if isNothing(str) or str = "" then
				Return 0
			else
				Dim uenc As New System.Text.UnicodeEncoding
				Return uenc.GetByteCount(str)
			end if
		End Function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function QS(ByVal str as string) as String
			if isNothing(str) or str = "" then
				return HttpContext.Current.request.querystring().toString()
			else
				return HttpContext.Current.request.querystring(str)
			end if	
		End Function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function QS2(ByVal str as string) as String
			dim retorno, sQuery as string
			if lenb( qs("uri") ) > 0 then
				dim uri as string = replace(HttpContext.Current.Request.RawUrl,"/__webapp/default.aspx?idCidade=4787&domain=netcampos.com&uri=","")
				dim tempURI as New System.Uri(domainName() & uri)
				sQuery = tempURI.Query
				retorno = HttpUtility.ParseQueryString(sQuery).Get(str)
			end if	
			return retorno
		End Function			
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function currentURL() as String
			dim retorno as string = domainName() & qs("uri")
			return retorno
		End Function
			
			
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function requestmethod() as String
			return lcase(HttpContext.Current.request.serverVariables("REQUEST_METHOD"))
		End Function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function getSessionID() as String
			dim retorno as string = HttpContext.Current.Session.SessionID		
			return retorno
		End Function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function RF(ByVal str as string) as String
			dim retorno as string = ""
			if isNothing(str) or str = "" then
				retorno = HttpContext.Current.request.form().toString()
			elseif requestmethod = "post"
				retorno =  HttpContext.Current.request.form(str).toString()
			end if
			return retorno
		End Function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public sub write(ByVal str as string)
			HttpContext.Current.response.write(str)
		End Sub
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:	writes a line to the output
		'' @PARAM:			value [string]: output string
		'******************************************************************************************************************
		Public sub writeln(ByVal str as string)
			HttpContext.Current.response.write(str & vbcrlf)
		end sub
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************	
		Public Function getGUID() as String  
		  Dim guidResult as String = System.Guid.NewGuid().ToString()
		  guidResult = guidResult.Replace("-", String.Empty)
		  Return guidResult
		End Function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public function mapPath(ByVal path as string) as string
			return httpContext.Current.server.mappath(path)
		End function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public function ipInternauta() as string
			return HttpContext.Current.request.ServerVariables("REMOTE_ADDR")
		End function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function getMD5(ByVal SourceText As String) As String
			Dim Ue As New UnicodeEncoding()
			Dim ByteSourceText() As Byte = Ue.GetBytes(SourceText)
			Dim hash() As Byte
			Dim Md5 As New MD5CryptoServiceProvider
			hash = Md5.ComputeHash(ByteSourceText)
			
			dim sb as new System.Text.StringBuilder
			dim outputByte as byte
			for each outputByte in hash
				sb.Append(outputByte.ToString("x2").ToUpper)
			next outputByte
			
			Return lcase(sb.ToString)
		End Function
		
		'***********************************************************************************************************
		'* description:
		'***********************************************************************************************************		
		Public Function getTemplatePart(ByVal virtualPath As String) As [String]
			Dim page As New _NPMarkupPage()
			try	
				Dim ctl As UserControl = DirectCast(page.LoadControl(virtualPath), UserControl)
				page.Controls.Add(ctl)
				Dim sb As New StringBuilder()
				Dim writer As New StringWriter(sb)
			
				page.Server.Execute(page, writer, True)
				
				Return sb.ToString()
			Catch ex As Exception
				NotFound(virtualPath)
			End Try			
			
		End Function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public sub responseEnd()
			HttpContext.Current.response.end()
		End Sub	
			
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function fileExists(byVal arquivo as string) as boolean
			dim retorno as boolean 
			if lenb(arquivo) > 0 then
				If System.IO.File.Exists(mapPath(arquivo)) Then retorno = true
			end if
			return retorno
		End Function
	
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Sub NotFound(Optional ByVal arquivo as string = "")
			HttpContext.Current.Response.Clear()
			
			try
				HttpContext.Current.Server.Execute("/__webapp/errors/404.aspx",false)
				HttpContext.Current.Response.Status = "404 Not Found"
				HttpContext.Current.Response.StatusCode = 404
				responseEnd()				
			catch ex as exception
				write("<div style=""width:auto;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Não Encontrado:</h3><hr/>" & arquivo & "</div>" )
				HttpContext.Current.Response.Status = "404 Not Found"
				HttpContext.Current.Response.StatusCode = 404
				responseEnd()
			end try
			
		End Sub
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:	writes a string to the output in the same line
		'' @PARAM:			value [string]: output string
		'******************************************************************************************************************
		Public sub setHttpHeader()
			HttpContext.Current.Response.Clear()
			HttpContext.Current.Response.ClearContent()
			HttpContext.Current.Response.ClearHeaders()			
			HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
			HttpContext.Current.Response.ContentType = "text/html"
			HttpContext.Current.Response.AppendHeader("Server", "NetCampos/1.0")
			HttpContext.Current.Response.AppendHeader("X-Powered-By","Grupo NetCampos Tecnologia")
			
			HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.Public)
			HttpContext.Current.Response.Cache.SetValidUntilExpires(true)			
			HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddMonths(1))
			HttpContext.Current.Response.Cache.SetLastModified(DateTime.Now.AddMonths(-1))
			HttpContext.Current.Response.Cache.SetRevalidation(HttpCacheRevalidation.AllCaches)
			HttpContext.Current.Response.Cache.SetOmitVaryStar(true)
			
			HttpContext.Current.Response.Cache.SetETag( """" &  getGUID.ToString().Replace( "-", "" ) & """" )
			
			'Dim ts As New TimeSpan(0,60,0)
			'HttpContext.Current.Response.Cache.SetMaxAge(ts)
			
			'Dim sIfModifiedSince As String = HttpContext.Current.Request.Headers("If-Modified-Since")
			

			'HttpContext.Current.Response.Cache.VaryByHeaders("Accept-Language") = true
			'HttpContext.Current.Response.Cache.VaryByHeaders("User-Agent") = true
			'HttpContext.Current.Response.Cache.SetLastModified(DateTime.Now.AddHours(-2))
			
			'If Not String.IsNullOrEmpty(sIfModifiedSince) then
				'HttpContext.Current.Response.Status = "304 Not Modified"
			'end if
			
						
			'HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
			'Dim ts As New TimeSpan(0,60,0)
			'HttpContext.Current.Response.Cache.SetMaxAge(ts)
			'HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddSeconds(3600))
			'HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
			'HttpContext.Current.Response.Cache.SetValidUntilExpires(true)
			'HttpContext.Current.Response.Cache.VaryByHeaders("Accept-Language") = true
			'HttpContext.Current.Response.Cache.VaryByHeaders("User-Agent") = true
			'HttpContext.Current.Response.AppendHeader("X-Powered-By","Grupo NetCampos Tecnologia")						
			
		End sub	
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************		
		Public sub defineHttpHeader(Optional ByVal Override304 As Boolean = False)
			Dim _user as new _appUsers()
					
			Try
				HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
				HttpContext.Current.Response.ContentType = "text/html"
				HttpContext.Current.Response.AppendHeader("Server", "NetCampos/1.0")
				HttpContext.Current.Response.AppendHeader("X-Powered-By","Grupo NetCampos Tecnologia")
				'HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddMonths(1))
				
				if not _user.UserON() then
				
					if Override304 then
						HttpContext.Current.Response.StatusCode = System.Net.HttpStatusCode.OK
					else		
						Dim sIfModifiedSince As String = HttpContext.Current.Request.Headers("If-Modified-Since")
						
						If Not String.IsNullOrEmpty(sIfModifiedSince) then
							HttpContext.Current.Response.StatusCode = 304
							HttpContext.Current.Response.StatusDescription = "Not Modified"
							HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.Public)
							HttpContext.Current.Response.Cache.SetExpires(setExpirerDate)
						else
							HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.Public)
							HttpContext.Current.Response.Cache.SetLastModified(DateTime.UtcNow)
							HttpContext.Current.Response.AddHeader("If-Modified-Since", DateTime.UtcNow.ToString("r"))
							
							Dim maxAge as integer = 86400 * 14
							HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddSeconds(maxAge))
							HttpContext.Current.Response.Cache.SetMaxAge(new TimeSpan(0, 0, maxAge))
							HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.Public)
							HttpContext.Current.Response.Cache.SetValidUntilExpires(true)
							
							setExpirerDate = DateTime.Now.AddSeconds(7200)						
						end if	
					end if
					
				end if
			
			Catch taex As System.Threading.ThreadAbortException
				Throw taex
			Catch ex As Exception
				System.Diagnostics.Debug.Print(ex.Message)
			End Try
			
		End Sub
				
	#End Region					
			

	#Region "Function reWrite"
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function getReWrite() as string
			dim retorno as string
			retorno = HttpContext.Current.Request.ServerVariables("HTTP_X_REWRITE_URL")
			if lenb(retorno) = 0 then
				retorno = HttpContext.Current.Request.ServerVariables("HTTP_X_ORIGINAL_URL")
			end if		
			return retorno
		End Function
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function pageController() as string
			dim controller as string = getURL(0)
			dim retorno as string
			if lenb(controller) > 0 then 
				retorno = controller
			else
				retorno = "home"
			end if
			return retorno
		end function		
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function getURL(byVal group as integer) as string
			dim retorno as string = ""
			Dim urlRoutes As New ArrayList()		
			Dim uri As String = QS("uri")
			Dim parts As String() = uri.Split(New Char() {"/"c})
			Dim part As String
				
			For Each part In parts
				if lenb(part) > 0 then
					urlRoutes.Add(part)
				end if
			Next		
			if ( urlRoutes.count() - 1 ) >= group then
				retorno = urlRoutes.Item(group)	
			end if
	
			return retorno
		End Function		
		
		
	#End Region 			
			
End class