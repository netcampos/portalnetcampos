Imports Microsoft.VisualBasic
Imports System.Configuration
Imports System.Text
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Xml
Imports System.Threading
Imports System.Globalization
Imports System.Net

Public Class _lib
		
	Private shared urlRewriteBase as string = "/App/web/controllers/institucionais/?uri="
		
	Public Sub New()
		MyBase.New()
	End Sub
	
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	just an equivalent for <em>request.querystring</em>. if empty then returns whole querystring.
	'' @PARAM:			name [string]: name of the value you want to get. leave it EMPTY to get the whole querystring
	'' @RETURN:			[string] value from the request querystring collection - implementation in page class
	'******************************************************************************************************************	
	Public Shared Function QS(ByVal str as string) as String
		if isNothing(str) or str = "" then
			return HttpContext.Current.request.querystring().toString()
		else
			return HttpContext.Current.request.querystring(str)
		end if	
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	just an equivalent for <em>request.querystring</em>. if empty then returns whole querystring.
	'' @PARAM:			name [string]: name of the value you want to get. leave it EMPTY to get the whole querystring
	'' @RETURN:			[string] value from the request querystring collection - implementation in page class
	'******************************************************************************************************************	
	Public Shared Function baseURL(ByVal str as string) as String
		dim urlSite as string = HttpContext.Current.Request.Url.Scheme & "://" & HttpContext.Current.Request.Url.Host & "/"
		dim url as string
		select case str
			case "institucional"
			url = urlSite & "sobre-campos-do-jordao/"
		end select
		return url
	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	just an equivalent for <em>request.querystring</em>. if empty then returns whole querystring.
	'' @PARAM:			name [string]: name of the value you want to get. leave it EMPTY to get the whole querystring
	'' @RETURN:			[string] value from the request querystring collection - implementation in page class
	'******************************************************************************************************************	
	Public Shared Function baseImages(byval conteudo as string) as String
		dim html as string = conteudo
		dim base as string = "http://images.netcampos.com/"
		if lenb(conteudo) > 0 then
			html = replace(html,"http://www.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/",base)
			html = replace(html,"http://www.netcampos.com/cidades/cms/netgallery/media/saopaulo/camposdojordao/",base)
			html = replace(html,"http://cms.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/",base)			
			html = replace(html,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/",base)
		end if
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	just an equivalent for <em>request.querystring</em>. if empty then returns whole querystring.
	'' @PARAM:			name [string]: name of the value you want to get. leave it EMPTY to get the whole querystring
	'' @RETURN:			[string] value from the request querystring collection - implementation in page class
	'******************************************************************************************************************	
	Public Shared Function baseThumb(byval image as string, byval largura as integer, byval altura as integer) as String
		dim html as string = image
		dim base as string = "http://images.netcampos.com/thumb/"
		if lenb(image) > 0 then
			html = replace(html,"http://www.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"http://www.netcampos.com/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"http://cms.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
		end if
		html = base & largura & "/" & altura & "/" & html
		return html
	End Function	
	

	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function RF(ByVal str as string) as String
		if isNothing(str) or str = "" then
			return HttpContext.Current.request.form().toString()
		else
			return HttpContext.Current.request.form(str).toString()
		end if
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function getMD5(ByVal SourceText As String) As String
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
	Public shared function ipInternauta() as string
		return HttpContext.Current.request.ServerVariables("REMOTE_ADDR")
	End function		
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared sub write(ByVal str as string)
		HttpContext.Current.response.write(str)
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared Function REWRITE_URL()
		return HttpContext.Current.Request.ServerVariables("HTTP_X_REWRITE_URL")
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared Function urlRewrite()
		dim retorno as string
		retorno = HttpContext.Current.Request.ServerVariables("HTTP_X_REWRITE_URL")
		if lenb(retorno) = 0 then
			retorno = HttpContext.Current.Request.ServerVariables("HTTP_X_ORIGINAL_URL")
		end if		
		return retorno
	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared Function getUrlRewrite()
		dim retorno as string
		retorno = HttpContext.Current.Request.ServerVariables("HTTP_X_REWRITE_URL")
		if lenb(retorno) = 0 then
			retorno = HttpContext.Current.Request.ServerVariables("HTTP_X_ORIGINAL_URL")
		end if
		
		if lenb(urlRewriteBase) > 0 then retorno = replace(retorno, urlRewriteBase, "")
		
		return retorno
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Sub pageExecute(byval arquivo as string)
		If not System.IO.File.Exists( mapPath(arquivo) ) Then
			pageNotFound()
		else
			HttpContext.Current.response.clear()
			HttpContext.Current.server.Execute(arquivo)
			HttpContext.Current.response.end()			
		end if
	End Sub
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Sub pageExecute1(byval arquivo as string)
		If not System.IO.File.Exists( mapPath(arquivo) ) Then
			pageNotFound()
		else
			'HttpContext.Current.response.clear()
			HttpContext.Current.server.Execute(arquivo)
			'HttpContext.Current.response.end()			
		end if
	End Sub		
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared Function Capitalize(byVal str as String) as String
		dim retorno as string = str
		if _lib.lenb(retorno) > 0 then
			retorno = StrConv(retorno, VbStrConv.ProperCase)
		end if
		return retorno
	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared Function MakeLinksComments(StringValue As String)
		Dim uriTube as String = String.Empty
		Dim strContent As String = StringValue
		Dim urlregex As Regex = New Regex("(?<!http://)(?<!https://)(www\.([\w.]+\/?)\S*)", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))
		strContent = urlregex.Replace(strContent, "http://$1")	
		Dim urlregex2 As Regex = New Regex("(http:\/\/([\w.]+\/?)\S*|https:\/\/([\w.]+\/?)\S*)", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))
		strContent = urlregex2.Replace(strContent, "<a href=""$1"" target=""_blank"">$1</a>")
		Return strContent
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Shared Function clearCommentsPost(byVal comments as String) as string
		dim html as string
		if _lib.lenb(comments) > 0 then
			Dim objRegExp as New RegEx("<(.|\n)+?>")
			comments = objRegExp.Replace(comments, "")
			comments = HttpUtility.HtmlAttributeEncode(comments) 
			comments = _lib.MakeLinksComments(comments)
			
			html = comments
		end if
		
		return html
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function getRoute(byVal group as integer) as string
		Dim route as String = clearRoute()
		Dim pattern As String
		Dim rewrite_Match As Match
		Dim retorno as String
		select case group
			case 1
				pattern = "(^\/(.*?)(.+?)\/)"
				rewrite_Match = Regex.Match(route, pattern)
				if rewrite_Match.Success then
					retorno = rewrite_Match.Groups(3).Captures(0).ToString()
				end if
		end select
		return retorno
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function clearRoute() as string
		Dim route as String = REWRITE_URL()
		route = replace(route,"/App/web/controllers/institucional/?uri=","")
		Dim pattern As String = "[\?&](?<name>[^&=]+)=(?<value>[^&=]+)"
		Dim myRegexOptions As RegexOptions = RegexOptions.Multiline Or RegexOptions.ECMAScript
		Dim myRegex As New Regex(pattern, myRegexOptions)
		Dim rewrite_Match As Match
		Dim retorno as String
			rewrite_Match = Regex.Match(route, pattern)
			
			if rewrite_Match.Success then
				retorno = myRegex.Replace(route, "")
			else
				retorno = route
			end if
			
		return retorno
	End Function			
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function getURL(byVal group as integer) as string
		dim retorno as string = ""
		Dim urlRoutes As New ArrayList()		
		Dim uri As String = urlRewrite()
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
		
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function slugContent() as string
		Dim retorno as string
		dim slug as string
		Dim uri As String = _lib.QS("slug")
		 if _lib.lenB(uri) > 0 then retorno = uri
		return retorno
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function slugHTML(slug) as String
		Dim retorno as String = ""
		try	
			Dim slashPosition As Integer = slug.IndexOf(".")
			Dim id As String = slug.Substring(slashPosition)
			id = replace(slug,id,"")	
			retorno = id
		Catch ex As Exception
			retorno = ""
		End Try
		return retorno
	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************	
	Public shared function dataBR(data As Date) as string
		dim out as String = ""
		dim dia as string = ""
		dim mes as string = ""
		dim ano as string = ""
		dia = day(data)
		mes = month(data)
		ano = year(data)
		
		if dia < 10 then dia = "0" & dia
		if mes < 10 then mes = "0" & mes
		out = dia & "/" & mes & "/" & ano
		
		return out
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************	
	Public shared function stringToDate(data As String) as date
		Dim asData as Date
		Dim dia, mes, ano as string
		If ( InStr(1, data, "/") ) Then
			dim np() as String
			np = data.split("/")
			dia = np(1)
			mes = np(0)
			ano = np(2)
			asData = New Date(ano, mes, dia)
		End If	
		return asData
	end function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************	
	Public shared function dataUSA(data As Date) as string
		dim out as String = ""
		dim dia as string = ""
		dim mes as string = ""
		dim ano as string = ""
		dia = day(data)
		mes = month(data)
		ano = year(data)
		
		if dia < 10 then dia = "0" & dia
		if mes < 10 then mes = "0" & mes
		out = ano & "-" & mes & "-" & dia
		return out
	end function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	'funcao para retornar data em microformats
	Public shared function dataMicroformats(data As Date) as string
	Dim out as String = String.Empty
		out = month(data)
		out = year(data) & "-" & month(data) & "-" & day(data) & "T" & hour(data) & ":" & minute(data) & ":" & second(data) & "Z"
		return out
	end function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************		
	public Shared function dataMateria(data As Date) as string
		dim out, dia, mes, ano, hora, minuto as string 
		dia = day(data)
		mes = month(data)
		ano = year(data)
		hora = hour(data)
		minuto = minute(data)
		if dia < 10 then dia = "0" & dia
		if mes < 10 then mes = "0" & mes
		if hora < 10 then hora = "0" & hora
		if minuto < 10 then minuto = "0" & minuto
		out = dia & "/" & mes & "/" & ano & " " & hora & "h" & minuto
		return out
	end function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Shared function dataMomento2(ByVal data as date) as string
		Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
		Dim calBR as New CultureInfo("pt-BR")
	
		Dim dtTweet As Date = Data
		Dim dtAgora As Date = DateTime.Now
		Dim tempo As TimeSpan = dtAgora - dtTweet
		Dim dif as Integer
		Dim ts As TimeSpan = DateTime.Today.Subtract(dtTweet)
		Dim calendario as Calendar = calBR.Calendar
		
		Dim exibir as string = ""
		
		If tempo <= TimeSpan.FromSeconds(60) Then
			return "Um instante atrás"
		end if
		
		If tempo <= TimeSpan.FromMinutes(60) Then
			if tempo.Minutes > 1  then 
				return tempo.Minutes & " minutos atrás"
			else
				return "Um minuto atrás"
			end if
		end if	
		
		If tempo <= TimeSpan.FromHours(24) Then
			if tempo.Hours > 1  then 
				return tempo.Hours & " horas atrás"
			else
				return "Hoje ás " & FormatDateTime(dtTweet, DateFormat.ShortTime) & " - " & tempo.Minutes & " minutos atrás"
			end if
		end if
		
		If tempo <= TimeSpan.FromDays(30) Then
			Dim semanas as Integer = DateDiff(DateInterval.Weekday, dtTweet, dtAgora) 
	
			if tempo.Days > 1  then
				if tempo.Days < 7 then
					return tempo.Days & " dias atrás - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
				else
					if semanas > 1 then return semanas & " semanas atrás - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
					return "semana passada - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
				end if
			else
				return "Ontém - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
			end if
		end if
		
		If tempo <= TimeSpan.FromDays(365) Then
			if tempo.Days > 30  then
				dif = New DateTime((DateTime.Today - dtTweet).Ticks).Month
				if dif > 1 then return dif & " meses atrás - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
				return dif & " meses atrás - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
			else
				return " Mês passado - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
			end if
		end if				
		
		if tempo.Days > 365  then
			dif = New DateTime((DateTime.Today - dtTweet).Ticks).Year - 1
	
			if dif > 1 then return dif & " anos atrás - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
			return dif & " ano atrás - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
			
		else
			return "Um ano atrás - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
		end if
		
		'return exibir
	end function	
	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared sub writeLn(ByVal str as string)
		HttpContext.Current.response.write(str & "<br/>" )
	End Sub	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared function mapPath(ByVal path as string) as string
		return httpContext.Current.server.mappath(path)
	End function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public shared Function LenB(ByVal value As String) As Integer
		if isNothing(value) or value = "" then
			Return 0
		else
			Dim uenc As New System.Text.UnicodeEncoding
			Return uenc.GetByteCount(value)
		end if
	End Function		
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************		
	Public shared function parse(byVal str as String, byVal alternative as String) as string
		parse = alternative
		Dim valor as object = trim(cstr(str) & "")
		if valor = "" then exit function
		Dim type As VariantType = VarType(parse)
		try
			Select case type.ToString()
				case "Integer"
				return clng(valor)
				case "String"
				return str.toString() & ""
			End Select
		catch ex as exception
			_lib.write("Impossível definir o tipo de objeto")
			exit function
		end try
	End function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared sub responseEnd()
		HttpContext.Current.response.end()
	End Sub	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public shared sub responseRedirect(byVal uri as string)
		HttpContext.Current.Response.redirect(uri)
		responseEnd()
	End Sub	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public shared sub responseRedirect301(byVal uri as string)
		HttpContext.Current.Response.Status = "301 Moved Permanently"
		HttpContext.Current.Response.AddHeader("Location",uri)
		responseEnd()
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	'funcao para embed swf
	Public shared Function drawSWF(ByVal src as String, ByVal width as Integer, ByVal height as Integer, ByVal FlashVars as String) as string
		dim swf as String = String.Empty
			swf = "<object codebase=""http://fpdownloadocument.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0"" classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" width=""" & width & """ height=""" & height & """>"
			swf = swf & "<param value=""" & src & """ name=""movie"">"
			swf = swf & "<param value=""transparent"" name=""wmode"">"
			swf = swf & "<param name=""quality"" value=""autohigh"">"
			swf = swf & "<param name=""bgcolor"" value=""#ffffff"">"
			if FlashVars <> "" then swf = swf & "<param name=""flashvars"" value=""" & FlashVars & """>"
			swf = swf & "<param name=""allowfullscreen"" value=""true"">"
			swf = swf & "<param name=""allowScriptAccess"" value=""always"">"
			swf = swf & "<param name=""quality"" value=""high"">"
			swf = swf & "<embed src=""" & src & """ type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" wmode=""transparent"""
			if FlashVars <> "" then swf = swf & " flashvars=""" & FlashVars & """"
			swf = swf & " width=""" & width & """ height=""" & height & """ allowFullScreen=""true"" bgcolor=""#FFFFFF"" allowScriptAccess=""always"" quality=""high"">"
			swf = swf & "</object>"
			return swf
		return swf	
	End function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	'verifica se imagem e retorna caso exista apenas para imagens representativas não membros
	Public shared function npFoto(ByVal img as String, ByVal titulo as String, ByVal href as String, ByVal largura as Integer, ByVal altura as Integer) as string
		dim html as string = String.Empty
		If not IsDBNull(img) and img <> "" then
			html = "<div class=""foto-legenda""><a href=""" & href & """ class=""borda-interna"" title=""" & titulo & """><img alt=""" & titulo & """ src=""http://images.guiadoturista.net/?img=" & img & "&amp;l=" & largura & "&amp;a=" & altura & """ width=""" & largura & """ height=""" & altura & """ /></a></div>"
		end if
		return html
	End function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	'verifica se imagem e retorna caso exista apenas para imagens representativas não membros
	Public shared function npThumb(ByVal img as String, ByVal titulo as String, ByVal href as String, ByVal largura as Integer, ByVal altura as Integer) as string
		dim html as string = String.Empty
		If not IsDBNull(img) and img <> "" then
			img = baseThumb(img,largura,altura)
			html = "<div class=""foto-legenda""><a href=""" & href & """ class=""borda-interna"" title=""" & titulo & """><img alt=""" & titulo & """ src=""" & img & """ width=""" & largura & """ height=""" & altura & """ /></a></div>"
		end if
		return html
	End function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	'verifica se imagem e retorna caso exista apenas para imagens representativas não membros
	Public shared function npThumbTerra(ByVal indice as Integer, ByVal img as String, ByVal Titulo as String) as string
		dim html as string = String.Empty
		dim width, height as integer
		If not IsDBNull(img) and img <> "" then
			select case indice
				case 1
				width = 195
				height = 130
				case 2
				width = 250
				height = 166
				case 3
				width = 89
				height = 59
				case 4
				width = 67
				height = 44
				case 5
				width = 185
				height = 166
				case 6
				width = 89
				height = 59
				case 7
				width = 135
				height = 90
				case 8
				width = 130
				height = 100
				case 9
				width = 150
				height = 100
				case 10
				width = 170
				height = 113
				case 11
				width = 185
				height = 125
				case 12
				width = 170
				height = 113
				case 13
				width = 170
				height = 113
				case 14
				width = 120
				height = 80
				case 15
				width = 80
				height = 53																																																								
			end select
			
			img = baseThumb(img,width,height)
			html = "<img src=""" & img & """ alt=""" & titulo & """ title=""" & titulo & """ />"
			
		end if
		return html
	End function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	'verifica se imagem e retorna caso exista apenas para imagens representativas não membros
	Public shared function npThumbImage(ByVal img as String, ByVal Titulo as String, ByVal width as integer, ByVal height as integer) as string
		dim html as string = String.Empty
		If not IsDBNull(img) and img <> "" then			
			img = baseThumb(img,width,height)
			
			html = "<img src=""" & img & """ alt=""" & titulo & """ title=""" & titulo & """ />"
		end if
		return html
	End function
	
	'******************************************************************************************************************
	'* Function Foto Guia
	'******************************************************************************************************************			
	Public shared Function thumbSite(ByVal img as String, ByVal largura as Integer, ByVal altura as Integer) as string
		dim html : html = ""
		If img <> "" then
			img = baseThumb(img,largura,altura)
			html = img
		end if
		return html
	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'*****************************************************************************************************************
	'funcao para embed swf
	Public shared Function playerSWF(ByVal src as String, ByVal width as Integer, ByVal height as Integer, ByVal FlashVars as String, ByVal nomeObjeto as String) as string
		dim swf as String = String.Empty
			swf = "<object width=""" & width & """ height=""" & height & """ codebase=""http://fpdownloadocument.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0"" classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" id=""" & nomeObjeto & """>"
			swf = swf & "<param value=""" & src & """ name=""movie"">"
			swf = swf & "<param value=""transparent"" name=""wmode"">"
			swf = swf & "<param name=""quality"" value=""autohigh"">"
			swf = swf & "<param name=""bgcolor"" value=""#FFFFFF"">"
			if FlashVars <> "" then swf = swf & "<param name=""flashvars"" value=""" & FlashVars & """>"
			swf = swf & "<param name=""allowfullscreen"" value=""true"">"
			swf = swf & "<param name=""allowScriptAccess"" value=""always"">"
			swf = swf & "<param name=""quality"" value=""high"">"
			swf = swf & "<embed src=""" & src & """ type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" wmode=""transparent"""
			if FlashVars <> "" then swf = swf & " flashvars=""" & FlashVars & """"
			swf = swf & " width=""" & width & """ height=""" & height & """ allowFullScreen=""true"" bgcolor=""#FFFFFF"" allowScriptAccess=""always"" quality=""high"">"
			swf = swf & "</object>"			
		return swf
	End function
	
	'******************************************************************************************************************
	'* Function Foto Legenda
	'******************************************************************************************************************			
	Public shared Function fotoLegenda(ByVal img, ByVal titulo, ByVal href, ByVal largura, ByVal altura, ByVal Legenda, ByVal target)
		dim html : html = ""
		dim alvo : alvo = ""
		if cstr(target) = "_blank" then
			alvo = "target=""_blank"""
		end if
		
		If img <> "" then			
			img = baseThumb(img,largura,altura)
			
			html = "<div class=""foto-legenda"">" & vbNewLine
				html = html & "<a href=""" & href & """ class=""borda-interna"" " & alvo & " title=""" & titulo & """>" & vbNewLine
					html = html & "<img alt=""" & titulo & """ src=""" & img & """ width=""" & largura & """ height=""" & altura & """ />" & vbNewLine
					if legenda <> "" then
						html = html & "<strong class=""bloco-texto""><span class=""meta"">" & Legenda & "</span></strong>" & vbnewline
					end if
				html = html & "</a>" & vbnewLine
				html = html & "<div class=""shadow""></div>"
			html = html & "</div>" & vbnewLine
		end if
		return html
	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Shortens a string and adds a custom string at the end if string is longer than a given value.
	'' @DESCRIPTION:	Useful if you want to display excerpts:
	''					<code><%= str.shorten("some value", 10, "...") % ></code>
	'' @PARAM:			str [string]: string which should be checked against cutting
	'' @PARAM:			maxChars [string]: whats the maximum allowed length of chars
	'' @PARAM:			overflowString [int]: what string should be added at the end of the string if it has been cutted
	'' @RETURN:			[string] cutted string
	'******************************************************************************************************************
	public shared function shorten(byVal str, maxChars, overflowString)
		str = str & ""
		if len(str) > maxChars then str = left(str, maxChars) & overflowString
		return str
	end function
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Shared Sub extractPageFromURI(ByVal uri as String)
	
		Dim Pagina as String = uri
		Dim iTotal as Integer = 0
		
		iTotal = Instr(Pagina, ".html")
		If iTotal > 0 Then
			pageLevel = 3
			pageName = getURLContentPage(uri)
		end if
	
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Shared Function getURLContentPage(ByVal pagina as String) as String
			Dim originalPage as String = pagina
			Dim iTotal as Integer
			Dim urlDefine as Boolean
			'verifica separadores no inicio
			'tabela
			',(virgula) tira o numerador e pega apenas o termo textual
			'_(underline) tira a string e devolve somente o numeral inicial
			'-(traco) tira a string e devolve somente o numeral inicial
			
			IF Not urlDefine then
				
				If ( InStr(1, pagina, ",") ) Then
					
					If ( InStrRev(pagina, ".") ) Then
						pagina = left(pagina, InStrRev(pagina, ".")-1)
					End If
					
					iTotal = InStr(1, pagina, ",")
					pagina = pagina.substring(itotal, pagina.length-itotal)
					
					if isNumeric(pagina) then
						pagina = originalPage
						pagina = pagina.substring(0, itotal-1)
						if isNumeric(pagina) then 
							pagina = originalPage
						else
							urlDefine = true
							return pagina
							Exit Function						
						end if
					Else					
						urlDefine = true
						return pagina
						Exit Function
					End If
					
				End If

			End IF
			
			IF Not urlDefine then
				If ( InStr(1, pagina, "_") ) Then
					
					If ( InStrRev(pagina, ".") ) Then
						pagina = left(pagina, InStrRev(pagina, ".")-1)
					End If
					
					iTotal = InStr(1, pagina, "_")
					pagina = pagina.substring(0, itotal-1)
					
					if isNumeric(pagina) then
						urlDefine = true
						return pagina
						Exit Function
					Else
						pagina = originalPage
					End If
					
				End If
			End IF						
			
			IF Not urlDefine then
				
				If ( InStr(1, pagina, "-") ) Then
					
					If ( InStrRev(pagina, ".") ) Then
						pagina = left(pagina, InStrRev(pagina, ".")-1)
					End If
					
					iTotal = InStr(1, pagina, "-")
					pagina = pagina.substring(0, itotal-1)
					
					if isNumeric(pagina) then
						urlDefine = true
						return pagina
						Exit Function
					Else
						pagina = originalPage
					End If
					
				End If
				
			End IF							
			'################################################################# '
			
			'verifica separadores no final
			'tabela
			',(virgula) tira o numerador e pega apenas o termo textual
			'_(underline) tira a string e devolve somente o numeral final
			'-(traco) tira a string e devolve somente o numeral final			
			IF Not urlDefine then
				
				If InStrRev(pagina,",",-1) then
					
					If ( InStrRev(pagina, ".") ) Then
						pagina = left(pagina, InStrRev(pagina, ".")-1)
					End If
					
					iTotal = InStr(1, pagina, ",")
					pagina = pagina.substring(0, itotal-1)
					
					if isNumeric(pagina) then
						pagina = originalPage
					Else					
						urlDefine = true
						return pagina
						Exit Function
					End If
										
				End if
				
			End IF		
			
			IF Not urlDefine then
				
				If InStrRev(pagina,"_",-1) then
					
					If ( InStrRev(pagina, ".") ) Then
						pagina = left(pagina, InStrRev(pagina, ".")-1)
					End If				
					
					iTotal = InStrRev(pagina,"_",-1)
					
					if isNumeric(pagina.remove(0,iTotal)) then 
						pagina = pagina.remove(0,iTotal)
						urlDefine = true
						return Pagina
						Exit Function
					Else
						pagina = originalPage
					End if
					
				End if
				
			End IF
			
			IF Not urlDefine then
				
				If InStrRev(pagina,"-",-1) then
					
					If ( InStrRev(pagina, ".") ) Then
						pagina = left(pagina, InStrRev(pagina, ".")-1)
					End If				
					
					iTotal = InStrRev(pagina,"-",-1)
					
					if isNumeric(pagina.remove(0,iTotal)) then 
						pagina = pagina.remove(0,iTotal)
						urlDefine = true
						return Pagina
						Exit Function
					Else
						pagina = originalPage
					End if
					
				End if
				
			End IF

			If Not urlDefine then
				If ( InStr(1, pagina, ".") ) Then
					dim np() as String
					np = pagina.split(".")
					pagina = np(0)
					urlDefine = true			
				End If
			End If							
		
			return Pagina		
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public shared sub Erro(msg, erro)
		dim html as string = "<div style=""width:600px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Execução do Aplicativo: <b>" & erro & "</b>, verifique e tente novamente.<br><br> <b>Mensagem:</b><hr/>" & msg & "</div>"
		HttpContext.Current.Response.Clear()
		'if debug then
			Write(html)
			responseEnd()
		'end if
		'responseEnd()
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Shared Sub pageNotFound()
		HttpContext.Current.Response.Clear()
		try
			HttpContext.Current.Server.Execute("/app/web/errors/404.aspx")
		catch ex as exception
			write("<div style=""width:auto;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante o Carregamento da Página, por favor tente novamente mais tarde.<br><br> <b>Atenção!!!</b><hr/>Se o erro persistir informe ao administrador do sistema.</div>" )
		end try
		responseEnd()
	End Sub
	
	'******************************************************************************************
	'' @PARAM:
	'******************************************************************************************
	public Shared function rReplace(val, pattern, replaceWith)
		Dim strTargetString As String = val	
		Dim strRegex as String = pattern
		Dim myRegexOptions As RegexOptions = RegexOptions.IgnoreCase
		Dim myRegex As New Regex(strRegex, myRegexOptions)
		Dim strReplace As String = replaceWith
		
		Return myRegex.Replace(strTargetString, strReplace)
	end function
	
	'******************************************************************************************
	'' @PARAM:
	'******************************************************************************************
	public Shared function returnContent(byVal content as string) as string
		dim html as string = content
		
		html = replace(html,"align=""justify""","")
		html = replace(html,"<p class=""corpo_documento_noticia"">","<p>")
		html = replace(html,"/web_data/modules/mod_mediaplayer/player.swf","/plugin/mod_mediaplayer/player.swf")
				
		if instr(html,"netplayer") > 0 or instr(html,"mod_mediaplayer") > 0 then
			if instr(html,".mp3") > 0 then
				html = replace(html,"http://www.netcampos.com/cidades/cms/netgallery/media/saopaulo/camposdojordao/","http://images.guiadoturista.net/audio/camposdojordao/")
				html = replace(html,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","http://images.guiadoturista.net/audio/camposdojordao/")
			end if
			
			if instr(html,".flv") > 0 then
				html = replace(html,"http://cms.guiadoturista.net/netplayer/netplayer.swf","/plugin/netplayer/netplayer.swf")
				html = replace(html,"/netplayer/skin/","/plugin/netplayer/skin/")
				html = replace(html,"/netplayer/logo.png","/plugin/netplayer/logo.png")
				html = replace(html,"http://www.netcampos.com/cidades/cms/netgallery/media/saopaulo/camposdojordao/videos/flv/","http://images.guiadoturista.net/camposdojordao/videos/flv/")
				html = replace(html,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/videos/flv/","http://images.guiadoturista.net/camposdojordao/videos/flv/")
				
				if instr(html,"mod_mediaplayer") > 0 then
					html = replace(html,"/plugin/mod_mediaplayer/player.swf","/plugin/player/player.swf")
				
				end if
			end if
			
		end if
		
		html = baseImages(html)
		
		if instr(html,"[caption") > 0 then
			html = rReplace(html, ":?(\[caption\s.*?(id=""(.*?)"").*?(align=""(.*?)"").*?(width=""(.*?)"").*?(caption=""(.*?)"").*(\]))?(<a.*[^>]+>?<img.*[^>]+>)?(</a>)?(\[/caption\])", "<div class=""gnc-caption $5 midia-largura-$7"" style=""width:$7px;"">$11<strong>$9</strong></div>")
		end if
		
		return html
	end function
	
	'******************************************************************************************
	'' @PARAM:
	'******************************************************************************************
	public Shared function getImageSite(byVal image as string) as string
		dim html as string = image
		html = baseImages(html)
		return html
	end function	

	'******************************************************************************************
	'' @PARAM:
	'******************************************************************************************
    public Shared function CountWords(ByVal str As String) As Integer
		' Count matches.
		if lenb(str) > 1 then
			Dim collection As MatchCollection = Regex.Matches(str, "\S+")
			Return collection.Count
		else
			return 0
		end if
    End Function
	
End class