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
Imports ComponentSoft.Net
Imports NetScraperLibrary
Imports HtmlAgilityPack

Public Class _lib2
		
	Public Sub New()
		MyBase.New()
	End Sub
	
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Public Shared sub setPageHttpHeader()
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
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function browser() as String
		dim p_browser as string
		dim agent as string = uCase(HttpContext.Current.request.serverVariables("HTTP_USER_AGENT"))
		if instr(agent, "MSIE") > 0 then
			p_browser = "IE"
		elseif instr(agent, "FIREFOX") > 0 then
			p_browser = "FF"
		elseif instr(agent, "CHROME") > 0 then
			p_browser = "CHORME"
		elseif instr(agent, "SAFARI") > 0 then
			p_browser = "SAFARI"				
		elseif instr(agent, "OPERA") > 0 then
			p_browser = "OPERA"								
		end if			
		return lcase(p_browser)
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function browserVersion() as String
		dim p_browser as string
		dim p_version as string = uCase(HttpContext.Current.request.Browser.MajorVersion())
		dim agent as string = uCase(HttpContext.Current.request.serverVariables("HTTP_USER_AGENT"))
		
		Dim strRegex as String = "CHROME/(?'version'(?'major'\d+)(?'minor'\.[.\d]*))"
		Dim myRegexOptions As RegexOptions = RegexOptions.Multiline Or RegexOptions.ECMAScript
		Dim myRegex As New Regex(strRegex, myRegexOptions)
		Dim strTargetString As String = ucase(agent)
		
		if instr(agent, "MSIE") > 0 then
			p_browser = p_version
		elseif instr(agent, "FIREFOX") > 0 then
			p_browser = p_version
		elseif instr(agent, "CHROME") > 0 then
			For Each myMatch As Match In myRegex.Matches(strTargetString)
				If myMatch.Success Then
					p_version = myMatch.Groups("major").toString()
			  	End If
			Next			
			p_browser = p_version
		elseif instr(agent, "SAFARI") > 0 then
			p_browser = p_version			
		elseif instr(agent, "OPERA") > 0 then
			p_browser = p_version						
		end if			
		return lcase(p_browser)
	End Function		
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function Device() as String
		Dim Agent as String = "desktop"
		Dim u As String = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
		Dim strMobile as String = "android.+mobile|avantgo|bada\/|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od)|iris|kindle|lge |maemo|midp|mmp|netfront|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|symbian|treo|up\.(browser|link)|vodafone|wap|windows (ce|phone)|xda|xiino"
		Dim strTablet as String = "ipad"
		Dim Mobile As New Regex(strMobile,RegexOptions.IgnoreCase)
		Dim Tablet As New Regex(strTablet,RegexOptions.IgnoreCase)
		if Mobile.IsMatch(u) then agent = "mobile"
		if Tablet.IsMatch(u) then agent = "tablet"
		return agent
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function isIphone() as boolean
		Dim retorno as boolean
		Dim u As String = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
		Dim strMobile as String = "iphone"
		Dim Mobile As New Regex(strMobile,RegexOptions.IgnoreCase)
		if Mobile.IsMatch(u) then retorno = true
		return retorno
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function requestmethod() as String
		return lcase(HttpContext.Current.request.serverVariables("REQUEST_METHOD"))
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function QS(ByVal str as string) as String
		if isNothing(str) or str = "" then
			return HttpContext.Current.request.querystring().toString()
		else
			return HttpContext.Current.request.querystring(str)
		end if	
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function RF(ByVal str as string) as String
		dim retorno as string
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
	Public shared sub write(ByVal str as string)
		HttpContext.Current.response.write(str)
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a line to the output
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public shared sub writeln(ByVal str as string)
		HttpContext.Current.response.write(str & vbcrlf)
	end sub	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function REWRITE_URL() as string
		dim retorno as string
		retorno = HttpContext.Current.Request.ServerVariables("HTTP_X_REWRITE_URL")
		if lenb(retorno) = 0 then
			retorno = HttpContext.Current.Request.ServerVariables("HTTP_X_ORIGINAL_URL")
		end if
		
		retorno = replace(retorno,"/_app_Site/app.aspx?host=www.netcampos.com&uri=","")
		
		return retorno
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	'Public shared sub writeLn(ByVal str as string)
		'HttpContext.Current.response.write(str & "<br/>" )
	'End Sub	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Shared Function getGUID() as String  
	  Dim guidResult as String = System.Guid.NewGuid().ToString()
	  guidResult = guidResult.Replace("-", String.Empty)
	  Return guidResult
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared function mapPath(ByVal path as string) as string
		return httpContext.Current.server.mappath(path)
	End function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared function ipInternauta() as string
		return HttpContext.Current.request.ServerVariables("REMOTE_ADDR")
	End function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function LenB(ByVal str As String) As Integer
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
	Public shared function trataNome(byval nome as string) as string
		Dim retorno as string
		if _lib2.lenb(nome) > 0 then
			Dim fullName as String() = nome.Split(New Char() {" "c})
			Dim firstName as String = fullName(0)
			Dim lastName as String = nome.substring(firstName.length())
			retorno = StrConv(firstName, VbStrConv.ProperCase) & " " & StrConv(lastName, VbStrConv.ProperCase)
			retorno = replace(retorno, "De ", "de ")
			retorno = replace(retorno, "Dos ", "dos ")
			retorno = replace(retorno, "Do ", "do ")
			retorno = replace(retorno, "Das ", "das ")
			retorno = replace(retorno, "Da ", "da ")			
			retorno = clearSpace(retorno)
		end if
		return retorno
	End function
	
	'******************************************************************************************************************
	'' @RETURN:	
	'******************************************************************************************************************	
	Public Shared Function baseThumb(byval image as string, byval largura as integer, byval altura as integer) as String
		dim html as string = image
		dim base as string = "http://images.netcampos.com/thumb/"
		
		if lenb(image) > 0 then
			html = replace(html,"http://www.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"http://www.netcampos.com/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"http://cms.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			
			if instr(image,"thumbUser") > 0 then
				base = "http://images.netcampos.com/thumbuser/"
				html = replace(html,"/cidades/cms/netgallery/media/thumbUser/","")
			end if
			
		end if
		
		
		html = base & largura & "/" & altura & "/" & html
		return html
	End Function
	
	'******************************************************************************************************************
	'' @RETURN:	
	'******************************************************************************************************************	
	Public Shared Function imageSafe(byval image as string, byval largura as integer, byval altura as integer, byval modo as string) as String
		dim html as string = image
		dim base as string = "/image-safe/"
		if lenb(image) > 0 then
			html = replace(html,"http://","")
			html = replace(html,"thumb.guiadoturista.net/?img=","")
		end if
		html = base & largura & "/" & altura & "/" & html & "&m=" & modo
		return html
	End Function
	
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
	'' @RETURN:	
	'******************************************************************************************************************	
	Public Shared Function getImage(byval image as string, byval s3 as integer) as String
		dim retorno as string
		if lenb(image) > 0 then
			if s3 = 0 then
				retorno = "http://images.netcampos.com/albumfotos/0/0/" & image
			else
				retorno = "/image/" & image
			end if
		end if
		return retorno
	End Function	
	
	'******************************************************************************************************************
	'* Function Foto Legenda
	'******************************************************************************************************************			
	Public shared Function fotoLegenda(ByVal img, ByVal titulo, ByVal href, ByVal largura, ByVal altura, ByVal Legenda, ByVal target) as string
		dim html : html = ""
		dim alvo : alvo = ""
		dim addRel : addRel = ""
		if cstr(target) = "_blank" then alvo = "target=""_blank"""
		
		If img <> "" then
			img = baseThumb(img,largura,altura)		
						
			html = "<div class=""foto-legenda"">" & vbNewLine
				html = html & "<a " & addRel & " href=""" & href & """ class=""borda-interna"" " & alvo & " title=""" & titulo & """>" & vbNewLine
					html = html & "<img title=""" & titulo & """ alt=""" & titulo & """ src=""" & img & """ width=""" & largura & """ height=""" & altura & """ />" & vbNewLine
					if legenda <> "" then
						html = html & "<strong class=""bloco-texto""><span class=""meta""><span>" & Legenda & "&nbsp;&nbsp;</span></span></strong>" & vbnewline
					end if
				html = html & "</a>" & vbnewLine
				html = html & "<div class=""shadow""></div>"
			html = html & "</div>" & vbnewLine
		end if
		return html
	End Function
	
	'******************************************************************************************************************
	'* Function Foto Legenda
	'******************************************************************************************************************			
	Public shared Function fotoNoHref(ByVal img as string, ByVal titulo as string, ByVal largura as integer, ByVal altura as integer, optional byval classe as string = "" ) as string
		dim html, css as string
		
		If img <> "" then
			img = baseThumb(img,largura,altura)	
			
			if lenb(classe) > 0 then css = "class=""" & classe & """"
			
			html = html & "<img " & css & " title=""" & titulo & """ alt=""" & titulo & """ src=""" & img & """ width=""" & largura & """ height=""" & altura & """ />" & vbNewLine
		end if
		return html
	End Function
	
	'******************************************************************************************************************
	'* Function Foto Guia
	'******************************************************************************************************************			
	Public shared Function fotoGuia(ByVal img as String, ByVal largura as Integer, ByVal altura as Integer) as string
		dim html : html = ""		
		If img <> "" then
			img = baseThumb(img,largura,altura)
			html = img
		end if
		return html
	End Function	
	
	'******************************************************************************************************************
	'* Function External Video
	'******************************************************************************************************************			
	Public shared Function externalVideoID(ByVal url as String) as string
		dim retorno : retorno = ""
		Dim strRegex, key as String
		Dim myRegexOptions As RegexOptions
		Dim myRegex As Regex
		'Dim match As Match
		
		If url <> "" then
			if instr(url, "youtube") > 0 then
				strRegex = "(?:youtube\.com/(?:user/.+/|(?:v|e(?:mbed)?)/|.*[?&]v=)|youtu\.be/)([^""&?/ ]{11})"
				myRegexOptions = RegexOptions.Multiline Or RegexOptions.ECMAScript
				myRegex = New Regex(strRegex, myRegexOptions)
				Dim match As Match = myRegex.Match(url)
				If match.Success Then
					key = match.Groups(1).Value
					retorno = key
				end if
			end if
		end if
		return retorno
	End Function	
		
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared function validateEmail(ByVal email as String) as boolean
		dim retorno as boolean
		dim teste as string
		if lenb(email) > 0 then
			Dim client As New EmailValidator()
			client.ValidationLevel = 5
			Dim result As ValidationLevel = client.Validate(email)
			if result.ToString() = "Success" then retorno = true
		end if
		return retorno
	End function	
		
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function PageExecute(ByVal path As String) As Integer
		HttpContext.Current.response.clear()
		HttpContext.Current.server.execute(path,false)
		HttpContext.Current.ApplicationInstance.CompleteRequest()
		'HttpContext.Current.response.end()
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
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
			write("Impossível definir o tipo de objeto")
			exit function
		end try
	End function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
    Public shared Function UnixToTime(ByVal strUnixTime As String) As Date
        UnixToTime = DateAdd(DateInterval.Second, Val(strUnixTime), #1/1/1970#)
        If UnixToTime.IsDaylightSavingTime = True Then
            UnixToTime = DateAdd(DateInterval.Hour, 1, UnixToTime)
        End If
    End Function

	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
    Public shared Function TimeToUnix(ByVal dteDate As Date) As String
        If dteDate.IsDaylightSavingTime = True Then
            dteDate = DateAdd(DateInterval.Hour, -1, dteDate)
        End If
        TimeToUnix = DateDiff(DateInterval.Second, #1/1/1970#, dteDate)
    End Function
	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared sub responseEnd()
		HttpContext.Current.response.end()
	End Sub	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared sub responseRedirect(byVal uri as string)
		HttpContext.Current.Response.redirect(uri)
		responseEnd()
	End Sub	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared sub responseRedirect301(byVal uri as string)
		HttpContext.Current.Response.Status = "301 Moved Permanently"
		HttpContext.Current.Response.AddHeader("Location",uri)
		responseEnd()
	End Sub
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function fileExists(byVal arquivo as string) as boolean
		dim retorno as boolean 
		if lenb(arquivo) > 0 then
			If System.IO.File.Exists(mapPath(arquivo)) Then
				retorno = true
			end if
		end if
		return retorno
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	public shared function shorten(byVal str, maxChars, overflowString)
		str = str & ""
		if len(str) > maxChars then str = left(str, maxChars) & overflowString
		return str
	end function	
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified stylesheet-file
	'******************************************************************************************************************
	Public shared Function loadCSS(byVal url as String, byVal media as string)
		dim html as string = ""
			html = html & ("<link type=""text/css"" rel=""stylesheet"" href=""" & url & """" & IIF(media <> "", " media=""" & media & """", string.empty ) & " />")
		return html
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified stylesheet-file
	'******************************************************************************************************************
	Public shared function loadJS(byVal url as String)
		dim html as string
		html = html & ("<script type=""text/javascript"" src=""" & url & """></script>")
		return html
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public shared Function getJS(byVal libJS as String)
		dim html as string
		
		html = html & loadJS("/media-assets/js/jquery.js,bootstrap.js,cycle.js")

		if lenb(libJS) > 0 then html = html & loadJS("/media-assets/js/" & libJS & ".js")
		
		return html
	End Function
		
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	executes a given javascript. input may be a string or an array. each field = a line
	'' @PARAM:			JSCode [string]. [array]: your javascript-code. e.g. <em>window.location.reload()</em>
	'***********************************************************************************************************
	public shared sub execJS(JSCode)
		writeln("<script>")
		if isArray(JSCode) then
			for i = 0 to uBound(JSCode)
				writeln(JSCode(i))
			next
		else
			writeln(JSCode)
		end if
		writeln("</script>")
	end sub		
				
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	public shared sub Erro(msg, erro)
		dim html as string = "<div style=""width:600px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Execução do Aplicativo: <b>" & erro & "</b>, verifique e tente novamente.<br><br> <b>Mensagem:</b><hr/>" & msg & "</div>"
		HttpContext.Current.Response.Clear()
		'if debug then
			Write(html)
			responseEnd()
		'end if
		'responseEnd()
	end sub
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	public shared sub dbOffline()
		dim html as string = "<div style=""width:600px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Servidor em Manutenção:</h3><hr/>Nosso servidor está temporariamente indisponível, volte em alguns minutos.</div>"
		HttpContext.Current.Response.Clear()
		Write(html)
		responseEnd()
	end sub	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Sub pageNotFound()
		HttpContext.Current.Response.Clear()
		try
			HttpContext.Current.Server.Execute("/app/desktop/errors/404.aspx")
		catch ex as exception
			write("<div style=""width:auto;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Página não Encontrada:</h3><hr/>A Página que você está tentando acessar não existe neste servidor.<br><br> <b>Atenção!!!</b><hr/>Verifique a url digitada.</div>" )
			HttpContext.Current.Response.Status = "404 Not Found"
			HttpContext.Current.Response.StatusCode = 404			
		end try
		responseEnd()
	End Sub
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Sub fileNotFound(Optional ByVal arquivo as string = "")
		HttpContext.Current.Response.Clear()
		try
			HttpContext.Current.Server.Execute("/app/desktop/errors/404.aspx",false)
		catch ex as exception
			write("<div style=""width:auto;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Arquivo Não Encontrado:</h3><hr/>Não foi possível encontrar o Arquivo neste Servidor.<br><br> <b>Nome do Arquivo:</b><hr/>" & arquivo & "</div>" )
			HttpContext.Current.Response.Status = "404 Not Found"
			HttpContext.Current.Response.StatusCode = 404
			responseEnd()
		end try
		
	End Sub
	
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
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'*****************************************************************************************************************
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
	'' @SDESCRIPTION:
	'******************************************************************************************************************	
	Public shared Function noPadding(ByVal v as Integer, ByVal qtd as Integer) as String
		if v mod qtd = 0 then
			return"<li class=""primeiro"">"
		else if (v+1) mod qtd = 0 then
			return "<li class=""ultimo"">"
		else
			return "<li>"
		end if 
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************		
	Public shared function noPaddingClass(ByVal v as Integer, ByVal qtd as Integer) as String
		if v = qtd then
			return "class=""sem-padding"""
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************		
	Public shared Function listaClear(ByVal v as Integer, ByVal qtd as Integer)
		if (v+1) mod qtd = 0 then
			return "<li class=""clear""></li>"
		end if
	End function
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************	
	Public shared Function iifElement(ByVal v as Integer, ByVal qtd as Integer, byVal elementTrue as string, byval elementFalse as string) as String
		if (v) mod qtd = 0 then
			return elementTrue
		else
			return elementFalse
		end if
	end function		
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************	
	Public shared function databr(data As Date) as string
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
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	'funcao para retornar data em microformats
	Public shared function datamicroformats(data As Date) as string
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
	Public Shared function dataMomento(ByVal data as date) as string
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
					return tempo.Days & " dias atrás - " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
				else
					if semanas > 1 then return semanas & " semanas atrás em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
					return "semana passada em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
				end if
			else
				return "Ontem - " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
			end if
		end if	
		
		If tempo <= TimeSpan.FromDays(365) Then
			if tempo.Days > 30  then
				dif = New DateTime((DateTime.Today - dtTweet).Ticks).Month
				if dif > 1 then return dif & " meses atrás em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
				return dif & " mês atrás em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
			else
				return "Mês passado - " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
			end if
		end if				
		
		if tempo.Days > 365  then
			dif = New DateTime((DateTime.Today - dtTweet).Ticks).Year - 1

			if dif > 1 then return dif & " anos atrás, em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " de " & dtTweet.Year & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)				
			return dif & " ano atrás, em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " de " & dtTweet.Year & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
			
		else
			return "Um ano atrás - " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
		end if
		
		'return exibir
	end function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Shared function dataPublishLogDate(ByVal data As Date) as string
		Dim dtTweet As Date = data
		return FormatDateTime(data, DateFormat.LongDate) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
	End function
	
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
				return "Ontem - " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " dia " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & "  de " & dtTweet.Year 
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
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************		
	Public Shared function dataTweet(ByVal data as date) as string		
		Dim retorno as string
		try
			Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
			Thread.CurrentThread.CurrentCulture = New CultureInfo("pt-BR")
		
			Dim dtTweet As DateTime = data
			Dim dtAgora As DateTime = DateTime.Now
			
			 If DateTime.Compare(dtAgora, dtTweet) >= 0 Then
			 	Dim timeSince As TimeSpan = dtAgora.Subtract(dtTweet)
				if (timeSince.TotalMilliseconds < 1) then
					retorno = "alguns segundos atrás"
				elseif (timeSince.TotalMinutes < 1) then
					retorno = "Um instante atrás"
				elseif (timeSince.TotalMinutes < 2) then
					retorno = "há ± 1 minuto atrás"
				elseif (timeSince.TotalMinutes < 60) then
					retorno = string.Format("há {0} minutos", timeSince.Minutes)
				elseif (timeSince.TotalMinutes < 120 ) then
					retorno = "há ± 1 hora atrás"
				elseif (timeSince.TotalHours < 24 ) then
					retorno = string.Format("há {0} horas ", timeSince.Hours)
				elseif (timeSince.TotalDays < 2 ) then
					retorno = string.Format("Ontem às {0} ", FormatDateTime(dtTweet, DateFormat.ShortTime))
				elseif (timeSince.TotalDays < 7 ) then
					retorno = "Na " & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
				elseif (timeSince.TotalDays < 60 ) then
					retorno = dtTweet.ToString("M", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
				elseif (timeSince.TotalDays < 365 ) then
					retorno = dtTweet.ToString("M", new CultureInfo("pt-BR"))
				else
					retorno = FormatDateTime(dtTweet, 1)
				end if
			 end if
		catch ex as exception
		
		end try
		
		return retorno
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************		
	Public Shared function dataTweet_old1(ByVal data as date) as string
		Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
		Thread.CurrentThread.CurrentCulture = New CultureInfo("pt-BR")
	
		Dim dtTweet As Date = data
		Dim dtAgora As Date = DateTime.Now
		Dim tempo As TimeSpan = dtAgora - dtTweet
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
				return "Ás " & FormatDateTime(dtTweet, DateFormat.ShortTime) & " - " & tempo.Minutes & " minutos atrás"
			end if
		end if
		
		If tempo <= TimeSpan.FromDays(30) Then
			if tempo.Days > 1  then
				if tempo.Days < 7 then
					return dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
				else
					return dtTweet.ToString("M", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
				end if
			else
				return "Ontem, às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
			end if
		end if	
		
		If tempo <= TimeSpan.FromDays(365) Then
			if tempo.Days > 30  then 
				return dtTweet.ToString("M", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
			else
				return "Um mês atrás "
			end if
		end if				
		
		if tempo.Days > 365  then 
			return dtTweet.ToString("Y", new CultureInfo("pt-BR")) & " às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
		else
			return "Um ano atrás "
		end if
		
		'return exibir
	end function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************		
	Public Shared function dataTweetShort(ByVal data as date) as string
		Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
		Thread.CurrentThread.CurrentCulture = New CultureInfo("pt-BR")
	
		Dim dtTweet As Date = data
		Dim dtAgora As Date = DateTime.Now
		Dim tempo As TimeSpan = dtAgora - dtTweet
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
				return "Ás " & FormatDateTime(dtTweet, DateFormat.ShortTime) & " - " & tempo.Minutes & " minutos atrás"
			end if
		end if
		
		If tempo <= TimeSpan.FromDays(30) Then
			if tempo.Days > 1  then
				if tempo.Days < 7 then
					return dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " passado"
				else
					return dtTweet.ToString("M", new CultureInfo("pt-BR"))
				end if
			else
				return "Ontem, às " & FormatDateTime(dtTweet, DateFormat.ShortTime)
			end if
		end if	
		
		If tempo <= TimeSpan.FromDays(365) Then
			if tempo.Days > 30  then 
				return dtTweet.ToString("M", new CultureInfo("pt-BR")) 
			else
				return "Um mês atrás "
			end if
		end if				
		
		if tempo.Days > 365  then 
			return dtTweet.ToString("M", new CultureInfo("pt-BR")) & " de " & dtTweet.Year 
		else
			return "Um ano atrás "
		end if
		
		'return exibir
	end function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	'Limpa Tags HTML
	Public Shared Function StripHTMLTags(ByVal html As String) As String
		Dim conteudo As String
		If html <> "" Then
			conteudo = Regex.Replace(html, "<(.|\n)+?>", String.Empty)
		End If
		Return conteudo
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************
	public shared function stripTags(byval html as string) as string			
		Dim conteudo As String
		If html <> "" Then
			conteudo = Regex.Replace(html, "<[^>]*>", String.Empty)
			conteudo = Regex.Replace(conteudo, "\[[^\]]*\]", String.Empty)
		End If
		Return conteudo
	end function		
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************	
    Public Shared Function trimComplete(ByVal html As String)
       	Dim strHtml As String = html
        strHtml = Regex.Replace(strHtml, "^\s+<", " <", RegexOptions.Multiline)
        strHtml = Regex.Replace(strHtml, ">\s+<", "> <", RegexOptions.Multiline)
        Return (strHtml.Trim)
    End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************	
    Public Shared Function clearSpace(ByVal html As String) as String
		Dim conteudo As String
		If html <> "" Then
			conteudo = Regex.Replace(html, "\s{2,}", " ")
			conteudo = replace(conteudo,"&nbsp;"," ")
		End If
		Return conteudo
    End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************	
    Public Shared Function clearSpaceLine(ByVal html As String) as String
		Dim conteudo As String
		If html <> "" Then
			conteudo = Regex.Replace(html, "\s{2,}", " " & vbnewline)
			conteudo = replace(conteudo,"&nbsp;"," "  & vbnewline)
		End If
		Return conteudo
    End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Shared Function getController() as string
		Dim url as String = REWRITE_URL()
		Dim controller as string = getRoute(1)	
		if lenb(controller) = 0 then controller = "home"
		return controller
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
	Public Shared Function MakeLinks(StringValue As String) as String
		dim retorno as string
		
		if _lib2.lenb(StringValue) > 0 then
			Dim uriTube as String = String.Empty
			Dim strContent As String = StringValue
			Dim urlregex As Regex = New Regex("(?<!http://)(?<!https://)(www\.([\w.]+\/?)\S*)", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))
			strContent = urlregex.Replace(strContent, "http://$1")	
			Dim urlregex2 As Regex = New Regex("(http:\/\/([\w.]+\/?)\S*|https:\/\/([\w.]+\/?)\S*)", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))
			strContent = urlregex2.Replace(strContent, "<a href=""$1"" target=""_blank"">$1</a>")
			retorno = strContent
		end if
		
		Return retorno
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Shared Function MakeLinks2(StringValue As String) as String
		dim retorno as string
		
		if _lib2.lenb(StringValue) > 0 then
				Dim strRegex as String = "(?i)\b((?:[a-z][\w-]+:(?:/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'"".,<>?«»“”‘’]))"
				Dim myRegexOptions As RegexOptions = RegexOptions.Singleline
				Dim myRegex As New Regex(strRegex, myRegexOptions)
				Dim strTargetString As String = StringValue
				Dim url, href as string
				
				For Each match As Match In myRegex.Matches(strTargetString)
					if match.Success Then
						url = match.Groups(1).Value
						if instr(url,"http://") > 0 or instr(url,"www.") > 0 then
							if instr(url,"http://") < 1 and instr(url,"www.") > 0 then 
								href = "http://" & url
							else
								href = url
							end if
							
							strTargetString = replace(strTargetString,url, "<a href=""" & href & """ rel=""nofollow"" target=""_blank"">" & url & "</a>" )
						end if
					end if
				next
								
				retorno = strTargetString
		end if
		
		Return retorno
	End Function	
		
	'***********************************************************************************************************
	'* description: retorno firs uri em uma string
	'***********************************************************************************************************
	Public shared Function returnFirstURI(byval StringValue As String) as string
		dim retorno as string
		Dim i as integer 		
		Dim urlregex As Regex = New Regex("(?<!http://)(?<!https://)(www\.([\w.]+\/?)\S*)", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))
		Dim strContent = urlregex.Replace(StringValue, "http://$1")
		
		Dim urlregex2 As Regex = New Regex("(http:\/\/([\w.]+\/?)\S*|https:\/\/([\w.]+\/?)\S*)", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))
		
		For Each match As Match In urlregex2.Matches(strContent)
			if i = 0 and match.Success Then
				retorno = match.Groups(1).Value
			end if
			i += 1
		next			
		return retorno
	End Function
	
	'***********************************************************************************************************
	'* description: retorno firs uri em uma string
	'***********************************************************************************************************
	Public shared Function extractFirstURI(byval StringValue As String) as string
		dim retorno as string
		Dim i as integer 		
		Dim urlregex As Regex = New Regex("(?i)\b((?:[a-z][\w-]+:(?:/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'"".,<>?«»“”‘’]))", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))		
		
		For Each match As Match In urlregex.Matches(StringValue)
			if i = 0 and match.Success Then
				retorno = match.Groups(1).Value
			end if
			i += 1
		next			
		return retorno
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function getURL(byVal group as integer) as string
		dim retorno as string = ""
		Dim urlRoutes As New ArrayList()		
		Dim uri As String = REWRITE_URL()
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
	Public shared Function fileExtension(str as string) as string
		dim retorno, extension as string
		dim dotPosition as integer = str.LastIndexOf(".")
		
		if lenb(str) > 0 and dotPosition > 0 then
			extension = str.Substring(dotPosition + 1)
			retorno = extension
		end if
				
		return retorno
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function slugContent() as string
		Dim retorno as string
		dim slug as string
		Dim uri As String = _lib2.REWRITE_URL()
		Dim parts As String() = uri.Split(New Char() {"/"c})
		
		if ( parts.count - 1 ) = 3 and instr(_lib2.getURL(parts.count - 2), ".html") > 0 then
			slug = _lib2.getURL(parts.count - 2)
			slug = replace(slug,".html","")
			retorno = slug
		elseif ( parts.count - 1 ) = 2 and instr(_lib2.getURL(parts.count - 2), ".html") > 0 then
			slug = _lib2.getURL(parts.count - 2)
			slug = replace(slug,".html","")
			retorno = slug
		elseif( parts.count - 1 ) = 3 and instr(_lib2.getURL(parts.count - 2), ".pdf") > 0 then
			slug = _lib2.getURL(parts.count - 2)
			slug = replace(slug,".pdf","")
			retorno = slug
		elseif ( parts.count - 1 ) = 2 and instr(_lib2.getURL(parts.count - 2), ".pdf") > 0 then
			slug = _lib2.getURL(parts.count - 2)
			slug = replace(slug,".pdf","")
			retorno = slug			
		elseif ( parts.count - 1 ) = 3 and _lib2.lenb(_lib2.getURL(1)) > 0 then
			slug = _lib2.getURL(1)
			retorno = slug
		elseif ( parts.count - 1 ) = 2 and _lib2.lenb(_lib2.getURL(0)) > 0 then
			slug = _lib2.getURL(0)
			retorno = slug
		end if
		
		return retorno
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function idContent(slug) as Integer
		Dim retorno as Integer = 0
		try	
			Dim slashPosition As Integer = slug.LastIndexOf("-")
			Dim id As String = slug.Substring(slashPosition + 1)	
			retorno = id
		Catch ex As Exception
			retorno = 0
		End Try			
		
		return retorno
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public shared Function slugHTML(slug) as String
		Dim retorno as String = ""
		try	
			Dim slashPosition As Integer = slug.IndexOf(".")
			Dim id As String = slug.Substring(slashPosition - 1)
			id = replace(slug,id,"")	
			retorno = id
		Catch ex As Exception
			retorno = ""
		End Try
		return retorno
	End Function
	
	'**********************************************************************************************************************
	'* description: 
	'**********************************************************************************************************************
	Public shared Function joinArrayList(ByVal list As ArrayList, ByVal delimeter As String) As String
	   Dim arrList As New ArrayList
	   
	   For Each item As Object In list
		  arrList.Add(item.ToString())
	   Next
	   
	   Return String.Join(delimeter, arrList.ToArray(GetType(String)))
	End Function
	
	'************************************************************************************************************************
	'* description: 
	'************************************************************************************************************************	
	Public shared Function npNav(pAtual as Integer, iTotalPaginas as Integer, maxLinkNav as Integer, url as String) as String
		dim html as string = ""
		dim uri as string = "/" & url & "/"
		dim intervalo, intInicioPaginas, intFinalPaginas as Integer
		
		intervalo = Int(maxLinkNav / 2)
		intInicioPaginas = pAtual - intervalo
		intFinalPaginas = pAtual + intervalo
		If CInt(intInicioPaginas) < 1 Then
				intInicioPaginas = 1
				intFinalPaginas = maxLinkNav
		End If
	
		If CInt(intFinalPaginas) > iTotalPaginas Then intFinalPaginas = iTotalPaginas
		If intFinalPaginas > maxLinkNav Then intFinalPaginas = intFinalPaginas
		
		html = "<div id=""paginacao""><ul>"
		
		If pAtual > 1 Then
			html = html & ("<li class=""ref""><a href=""" & uri & """ rel=""0"">|«</a></li>")
			html = html & ("<li class=""prev""><a href=""" & uri & "?page=" & pAtual - 1 & """ rel=""" & pAtual - 1 & """>«</a></li>")
		else
			html = html & ("<li class=""previous-off"">«</li>")
		End If
		
		For i = intInicioPaginas To intFinalPaginas
			If CInt(i)=CInt(pAtual) Then
				html = html & ("<li class=""active"">" & i & "</li>")
			End If
			If CInt(i) < CInt(pAtual) Then
				html = html & ("<li><a href=""" & uri & "?page=" & i & """ rel=""" & i & """>" & i & "</a></li>")
			End If	
			If CInt(i) > CInt(pAtual) Then
				html = html & ("<li><a href=""" & uri & "?page=" & i & """ rel=""" & i & """>" & i & "</a></li>")
			End If			
		next
		
		If CInt(pAtual) < CInt(iTotalPaginas) Then
			html = html & ("<li class=""next""><a href=""" & uri & "?page=" & pAtual + 1 & """ rel=""" & pAtual + 1 & """>»</a></li>")
			html = html & ("<li><a href=""" & uri & "?page=" & iTotalPaginas & """ rel=""" & iTotalPaginas & """>»|</a></li>")
		Else
			html = html & ("<li class=""next-off"">»</li>")
		End If
		
		html = html & "</ul></div>"
		
		if (iTotalPaginas < 2 ) then html = ""
	
		return html
	End Function	

	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************
	Public shared Function urlShorter(byval uri as string, Optional byVal servico as string = "gtur") as string
		dim html, url as string
		Dim request As HttpWebRequest
		Dim oResponse As HttpWebResponse
		Dim reader As StreamReader
		select case servico
			case "gtur"
				url = "http://gtur.in/api?url=" & uri
			case "tinyurl"
				url = "http://tinyurl.com/api-create.php?url=" & uri
			case "migre"
				url = "http://migre.me/api.txt?url=" & uri
		end select
		try
			if servico <> "google" then
				request = HttpWebRequest.Create(url)
				request.Method = WebRequestMethods.Http.GET
				oResponse = request.GetResponse()
				reader = New StreamReader(oResponse.GetResponseStream())
				html = reader.ReadToEnd()
				if html = "Please enter a valid url." and servico = "gtur" then  html = "error"
				oResponse.Close()
			else
				html = googleShortner(uri)
			end if
					
		Catch ex As Exception
			html = "error"
		Finally
			
		End Try 		
		
		if html = "error" and servico = "gtur" then html = _lib2.urlShorter(uri,"migre")
		if html = "error" and servico = "migre" then html = _lib2.urlShorter(uri,"tinyurl")
		if html = "error" and servico = "tinyurl" then html = _lib2.urlShorter(uri,"google")
			
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************	
	Public Shared Function googleShortner(ByRef url As String) As String
		Dim tmp, html as string
		Dim encoding As New UTF8Encoding
		Dim request as string = "https://www.googleapis.com/urlshortener/v1/url?key=AIzaSyB714EQg8RqHtdPwvs1rIPl9gs-VOSZsD8"
		Dim req As HttpWebRequest = CType(WebRequest.Create(request), HttpWebRequest)
		Dim byteData As Byte() = encoding.GetBytes("{'longUrl': '" & url & "'}")
		
		try
			req.ContentType = "application/json"
			req.ContentLength = byteData.Length
			req.KeepAlive = True
			req.Method = "POST"
			Dim stream01 As IO.Stream = req.GetRequestStream
			stream01.Write(byteData, 0, byteData.Length)
			
			Dim resp As HttpWebResponse = CType(req.GetResponse(), HttpWebResponse)
			Dim reader As New StreamReader(resp.GetResponseStream())
			tmp = reader.ReadToEnd()
			dim separa = Split(tmp,"""")
			html = separa(7)
		Catch ex As Exception
			html = url
		Finally
			
		End Try 			
		
		Return html
	End Function
	
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
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Shared Function getTableModule(ByVal id As Integer) As String
		Dim retorno As String
		if id > 0 then
			select case id
				case 3
					retorno = "_cms_view_cidades_empresas"			
				case 5
					retorno = "_cms_view_cidades_noticias"
				case 6
					retorno = "_cms_view_cidades_galeria_fotos"				
				case 7
					retorno = "_cms_view_cidades_passeios"
				case 10
					retorno = "_cms_view_cidades_videos"
				case 27
					retorno = "_cms_view_cidades_usuarios_fotos"					
			end select
		End If
		Return retorno
	End Function
	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Shared Function getTableMeta(ByVal id As Integer) As String
		Dim retorno As String
		if id > 0 then
			select case id
				case 3
					retorno = ""			
				case 5
					retorno = ""
				case 6
					retorno = ""				
				case 7
					retorno = ""
				case 10
					retorno = "_cms_view_cidades_videos_meta"
				case 27
					retorno = ""					
			end select
		End If
		Return retorno
	End Function	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Shared Function getSlugModule(ByVal id As Integer) As String
		Dim retorno As String
		if id > 0 then
			select case id
				case 3
					retorno = ""			
				case 5
					retorno = ""
				case 6
					retorno = ""				
				case 7
					retorno = ""
				case 10
					retorno = "videos-campos-do-jordao"
				case 27
					retorno = ""					
			end select
		End If
		Return retorno
	End Function	
	
	'***********************************************************************************************************
	'* Previsao do Tempo
	'***********************************************************************************************************
	public shared function weather_now() as string
		Dim html, html2, data as String
		Dim dtDate As Date
		Dim erro as boolean = false
		Dim fileWeather as string = lcase("/ReferencesCache/") & DateTime.Now.ToString("dd") & "_" & DateTime.Now.ToString("MM") & "_" & DateTime.Now.ToString("yyyy") & ".xml"
		Dim xml as XmlDocument = new XmlDocument()
		Dim m_nodelist As XmlNodeList
		Dim m_node As XmlNode
		
		html2 = html2 & "<a class=""follow"" target=""_blank"" title=""Siga-nos no Twitter"" href=""http://twitter.com/netcampos/"">"
			html2 = html2 & "<span>Siga-nos no Twitter</span>"
		html2 = html2 & "</a>"		
		
		If not System.IO.File.Exists( _lib2.mapPath(fileWeather) ) Then 
			dim remoteFile as string = "http://servicos.cptec.inpe.br/XML/cidade/1219/previsao.xml"
			try
				xml.Load(remoteFile)
				xml.Save(_lib2.mapPath(fileWeather))
			catch ex as exception
				erro = true
			end try
		end if
		
		If not System.IO.File.Exists( _lib2.mapPath(fileWeather) ) Then
			erro = true
		else
			html = ""
			try
				xml.Load( _lib2.mapPath(fileWeather))
				m_nodelist = xml.SelectNodes("/cidade/previsao")
				if m_nodelist.Count > 0 then
					
					html = html & "<div id=""gnc_slider_weather"">"
						For Each m_node In m_nodelist
							data = m_node.ChildNodes.Item(0).InnerText
							dtDate = DateTime.Parse(data, Globalization.CultureInfo.CreateSpecificCulture("pt-BR"))
							data = dtDate.ToString("dd") & "/" & dtDate.ToString("MM")
							
							html = html & "<div class=""inner"">"
								html = html & "<p class=""txt-tempo"">Previsão do Tempo - dia " & data & "</p>"
								html = html & "<img src=""/assets/img/weather/cptec/" & m_node.ChildNodes.Item(1).InnerText & ".png"" alt=""" & m_node.ChildNodes.Item(1).InnerText & """ />"
								html = html & "<div class=""previsao"">"
									html = html & "<div class=""max""><div class=""ico_max""></div><div class=""txt_max"">Max: " & m_node.ChildNodes.Item(2).InnerText & "<span class=""txtgraus"">C</span></div> </div>"
									html = html & "<div class=""min""><div class=""ico_min""></div><div class=""txt_min"">Min: " & m_node.ChildNodes.Item(3).InnerText & "<span class=""txtgraus"">C</span></div> </div>"
								html = html & "</div>"
								html = html & "<p class=""weatherUpdate""><span class=""txt-update"">Fonte: Cptec</span></p>"
							html = html & "</div>"
						next
					html = html & "</div>"
					
				end if
			catch ex as exception
				erro = true
			end try
		end if
		
		if erro then html = html2
		
		return html
	end function
		
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
	Public Shared Function formatURL(byval url as string) as string
		try	
								
			if (not Regex.IsMatch(url, "^http://", RegexOptions.IgnoreCase)) then
				url = "http://" + url
			end if
			
			try	
				Dim baseUri As New Uri(url)
			
			Catch ex As UriFormatException
				url = string.empty		
			End Try
			
		catch ex as exception
			url = string.empty
		end try				
		
		return url
	End Function
	
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
	Public Shared Function getMetaURL(byval url as string, byval meta as string) as string
		dim html as string
		if lenb(url) > 0 and lenb(meta) > 0 then		
			try
				Dim myHtml As New HtmlWeb
				Dim doc As HtmlDocument = myHtml.load(url)
				Dim Nodes = doc.DocumentNode.SelectNodes("//html")
				if not ( nodes Is Nothing ) then
					For Each node As HtmlNode In nodes
						if not ( node.SelectSingleNode(meta) Is Nothing ) and ( instr(meta,"title") < 1 ) then
							html = (node.SelectSingleNode(meta).GetAttributeValue("content", ""))
							exit for
						end if
						
						if ( instr(meta,"title") > 0 ) then
							html = (node.SelectSingleNode(meta).InnerText)
							exit for
						end if
					next
				end if				
			catch ex as exception
				html = ""
			end try
		end if
		
		return html
	End Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
	Public Shared Function loadMetaURL(byval url as string, byval meta as string) as string
		dim html as string
		if lenb(url) > 0 and lenb(meta) > 0 then		
			try
				Dim doc As new HtmlDocument
				doc.LoadHtml(loadHTMLDoc(url))
				Dim Nodes = doc.DocumentNode.SelectNodes("//html")
				if not ( nodes Is Nothing ) then
					For Each node As HtmlNode In nodes
						if not ( node.SelectSingleNode(meta) Is Nothing ) and ( instr(meta,"title") < 1 ) then
							html = (node.SelectSingleNode(meta).GetAttributeValue("content", ""))
							exit for
						end if
						
						if ( instr(meta,"title") > 0 ) then
							html = (node.SelectSingleNode(meta).InnerText)
							exit for
						end if
					next
				end if				
			catch ex as exception
				html = ""
			end try
		end if
		
		return html
	End Function		
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
	Public Shared Function getImageOG(byval url as string) as string
		dim html, img as string
		if _lib2.lenb(url) > 0 then		
			try
				Dim myHtml As New HtmlWeb
				Dim doc As HtmlDocument = myHtml.load(url)
				Dim Nodes = doc.DocumentNode.SelectNodes("//html//body//img")
				if not ( nodes Is Nothing ) then
					For Each node As HtmlNode In nodes
						img = _lib2.RelativeToAbsoluteUrl(url, node.Attributes("src").Value).toString()
						if _lib2.checkImageURI(img) then 
							html = img
							exit for
						end if
						'html = html & node.name & "-" & node.Attributes("src").Value & "-" & RelativeToAbsoluteUrl(url, node.Attributes("src").Value).toString() & "<br>"
					next
				end if				
			catch ex as exception
				html = ""
			end try
		end if
		return html
		
	end Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
	Public Shared Function getXmlImagesURL(byval url as string) as string
		dim html, img as string
		if _lib2.lenb(url) > 0 then		
			try
				Dim myHtml As New HtmlWeb
				Dim doc As HtmlDocument = myHtml.load(url)
				Dim Nodes = doc.DocumentNode.SelectNodes("//html//body//img")
				if not ( nodes Is Nothing ) then
					For Each node As HtmlNode In nodes
						img = _lib2.RelativeToAbsoluteUrl(url, node.Attributes("src").Value).toString()
						if _lib2.checkImageURI(img) then 
							html = html & "<item><![CDATA[" & img & "]]></item>"
						end if
					next
				end if				
			catch ex as exception
				html = ""
			end try
		end if
		return html
		
	end Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
    Public shared Function RelativeToAbsoluteUrl(ByVal uri As string, ByVal RelativeUrl As String) As Uri
       Dim baseUri As New Uri(uri)
	    ' get action tags, relative or absolute
        Dim uriReturn As Uri = New Uri(RelativeUrl, UriKind.RelativeOrAbsolute)
        ' Make it absolute if it's relative
        If Not uriReturn.IsAbsoluteUri Then
            Dim baseUrl As Uri = baseURI
            uriReturn = New Uri(baseUrl, uriReturn)
        End If
        Return uriReturn
    End Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
    Public Shared Function checkImageURI(ByVal uri As string) As Boolean
		Dim retorno as boolean
		try
			Dim webRequest As Net.HttpWebRequest = CType(Net.WebRequest.Create(uri), Net.HttpWebRequest)
			with webRequest
				.AllowAutoRedirect = true
				.Timeout = 1000 * 30
				.UserAgent = "Mozilla/5.0 (Windows NT 5.1; rv:15.0) Gecko/20100101 Firefox/15.0"
				.PreAuthenticate = true
			end with
			
			Dim doc As Net.HttpWebResponse = webRequest.GetResponse()
			if instr(doc.ContentType.toString(),"image") > 0 then 
				Dim img As System.Drawing.Bitmap = New System.Drawing.Bitmap(doc.GetResponseStream)
				if (img.width > 50 and img.width < 400) and (img.height > 50 and img.height < 400) then retorno = true
			end if
			doc.close()
		catch ex as exception
			retorno = false
		end try
		
		return retorno
    End Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
	Public Shared Function loadHTMLDoc(byval url as string) as string 
		Dim uri as New Uri(url)	
		Dim req As HttpWebRequest = CType(WebRequest.Create(uri), HttpWebRequest)
		Dim doc As HttpWebResponse = CType(req.GetResponse(), HttpWebResponse)
		Dim encode As Encoding = System.Text.Encoding.GetEncoding(doc.CharacterSet)
		Dim receiveStream As Stream = doc.GetResponseStream()
		Dim readStream As New StreamReader(receiveStream, encode)
		Dim strHtml As String = readStream.ReadToEnd()
		strHtml = System.Web.HttpUtility.HtmlDecode(strHtml)		
		doc.close()
		return strHtml
	End Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************		
    Public shared Function getEncodeURI(ByVal url As String) As String
		dim retorno as string
		Dim myHtml As New HtmlWeb
		Dim doc As HtmlDocument = myHtml.Load(url)
		retorno = doc.Encoding.toString()
		
		return retorno
    End Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************		
    Public shared Function getValidURL(ByVal url As String) As String
		dim retorno as string
		if lenb(url) > 0 then
			try
				url =  _lib2.formatURL(url)
				Dim baseUri As New Uri(url)
				url = baseUri.AbsoluteUri.toString()
				
				Dim webRequest As Net.HttpWebRequest = CType(Net.WebRequest.Create(url), Net.HttpWebRequest)
				with webRequest
					.AllowAutoRedirect = true
					.Timeout = 1000 * 30
					.UserAgent = "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)"
					.PreAuthenticate = true
				end with
				
				Dim doc As Net.HttpWebResponse = webRequest.GetResponse()										
				url = doc.ResponseUri.toString()	
				doc.close()
				retorno = url
				
			Catch ex As Exception
				retorno = ""
			end try
		end if		
		return retorno
    End Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************		
    Public shared Function getDomain(ByVal url As String) As String
		dim retorno as string
		if lenb(url) > 0 then
			try
				url =  _lib2.formatURL(url)
				Dim baseUri As New Uri(url)
				url = baseUri.Host.ToLower()				
				
				retorno = url
			Catch ex As Exception
				retorno = ""
			end try
		end if		
		return retorno
    End Function	
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Public shared Function formataScrapeHTML(Byval html as string) as string
		dim strHtml as string = html
		strHtml = Regex.Replace(strHtml, "&gt; ", ">" & vbnewline, RegexOptions.Multiline)		
		return strHtml
	End Function
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Public shared Function tiraAscento(ByVal texto as String) as string
		Dim str1 as String
		Dim str2 as String
		Dim i as Integer
		texto = trim(texto)
		i = 0
		str1 = "ÄÅÁÂÀÃäáâàãÉÊËÈéêëèÍÎÏÌíîïìÖÓÔÒÕöóôòõÜÚÛüúûùÇç"
		str2 = "AAAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUuuuuCc"
		Do While i < Len(str1)
			i = i + 1
			texto = Replace(texto, Mid(str1, i, 1), Mid(str2, i, 1))  
		loop
		return lcase((texto))
	end Function		
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
	Public shared Function GenerateSlug(byval phrase As String, optional byval maxLength As Integer = 0) As String
		Dim str As String = phrase.ToLower()
		str = Regex.Replace(str, "[^a-z0-9\s-]", "")
		str = Regex.Replace(str, "[\s-]+", " ").Trim()
		'if maxLength > 0 then str = str.Substring(0, If(str.Length <= maxLength, str.Length, maxLength)).Trim()
		str = Regex.Replace(str, "\s", "-")
		Return str
	End Function	
		
End class