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
		''' <summary>
		''' Initialize the program.
		''' </summary>	
		Public Sub New()
			MyBase.New()
		End Sub
		
		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub
	#End Region
	
	#Region "Comom Functions"
		
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
				'NotFound(ex.message())
			End Try			
			
		End Function		
				
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
			if lenb(controller) > 0 then retorno = controller
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