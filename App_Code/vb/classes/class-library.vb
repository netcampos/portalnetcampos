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

Public Class _npLibrary
		
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
			Return System.Configuration.ConfigurationManager.AppSettings.Get("domainName")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property domainAssets() As String
		Get
			Return ConfigurationManager.AppSettings("domainAssets")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property
		
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property domainImages() As String
		Get
			Return ConfigurationManager.AppSettings("domainImages")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property domainThumb() As String
		Get
			Return ConfigurationManager.AppSettings("domainThumb")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property			
		
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Function browser() as String
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
	Public Function browserVersion() as String
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
	Public Function Device() as String
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
	Public Function isIphone() as boolean
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
	Public Function requestmethod() as String
		return lcase(HttpContext.Current.request.serverVariables("REQUEST_METHOD"))
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
	Public Function RF(ByVal str as string) as String
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
	Public shared sub responseEnd()
		HttpContext.Current.response.end()
	End Sub		
	
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
	Public Function getController(byval currentURL as String) as string
		Dim controller as string = getRoute(1, currentURL)	
		if lenb(controller) = 0 then controller = "home"
		return controller
	End Function
		
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Function getRoute(byVal group as integer, byval currentURL as String) as string
		Dim route as String = clearRoute(currentURL)
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
	Public Function clearRoute(byval currentURL as String) as string
		Dim route as String = currentURL
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
	Public Sub pageNotFound()
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
	Public function trataNome(byval nome as string) as string
		Dim retorno as string
		if lenb(nome) > 0 then
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
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'******************************************************************************************************************	
    Public Function clearSpace(ByVal html As String) as String
		Dim conteudo As String
		If html <> "" Then
			conteudo = Regex.Replace(html, "\s{2,}", " ")
			conteudo = replace(conteudo,"&nbsp;"," ")
		End If
		Return conteudo
    End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Function currentURL() as String
		dim retorno as string = domainName() & HttpContext.Current.Request.RawUrl	
		return retorno
	End Function
	
	'******************************************************************************************************************
	'' @RETURN:	
	'******************************************************************************************************************	
	Public Function npThumb(byval image as string) as String
		dim html as string = image
		
		if lenb(image) > 0 then
			html = replace(html,"http://www.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"http://www.netcampos.com/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"http://cms.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")			
		end if
		
		html = domainImages & "/" & html
		return html
	End Function
		
	'******************************************************************************************************************
	'* Function Foto Legenda
	'******************************************************************************************************************			
	Public Function fotoNoHref(ByVal img as string, ByVal titulo as string, ByVal largura as integer, ByVal altura as integer, optional byval classe as string = "" ) as string
		dim html, css as string
		
		If img <> "" then
			img = baseThumb(img,largura,altura)	
			
			if lenb(classe) > 0 then css = "class=""" & classe & """"
			
			html = html & "<img " & css & " title=""" & titulo & """ alt=""" & titulo & """ src=""" & img & """ width=""" & largura & """ height=""" & altura & """ />" & vbNewLine
		end if
		return html
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
	'' @RETURN:	
	'******************************************************************************************************************	
	Public Function baseThumb(byval image as string, byval largura as integer, byval altura as integer) as String
		dim html as string = image
				
		if lenb(image) > 0 then
			html = replace(html,"http://www.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"http://www.netcampos.com/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
			html = replace(html,"http://cms.guiadoturista.net/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
		end if

		html = domainThumb & "/" & largura & "/" & altura & "/" & html
		return html
	End Function
		
	'***********************************************************************************************************
	'* description: Data
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
		
	'******************************************************************************************************************
	'* Function Foto Legenda
	'******************************************************************************************************************			
	Public Function fotoLegenda(ByVal img as string, ByVal titulo as string, ByVal href as string, ByVal largura as integer, ByVal altura as integer, ByVal Legenda as string, ByVal target as string, ByVal follow as string) as string
		dim html : html = ""
		dim alvo : alvo = ""
		dim addRel : addRel = ""
		if cstr(target) = "_blank" then alvo = "target=""_blank"""
		if cstr(follow) = "nofollow" then addRel = "rel=""nofollow"""
		
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
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	'verifica se imagem e retorna caso exista apenas para imagens representativas não membros
	Public function npThumbImage(ByVal img as String, ByVal Titulo as String, ByVal width as integer, ByVal height as integer) as string
		dim html as string = String.Empty
		If not IsDBNull(img) and img <> "" then			
			img = baseThumb(img,width,height)
			
			html = "<img width=""" & width & """ height=""" & height & """ src=""" & img & """ alt=""" & titulo & """ title=""" & titulo & """ />"
		end if
		return html
	End function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'*****************************************************************************************************************
	Public Function playerSWF(ByVal src as String, ByVal width as Integer, ByVal height as Integer, ByVal FlashVars as String, ByVal nomeObjeto as String) as string
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
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public sub redirect301(byVal uri as string)
		HttpContext.Current.Response.Status = "301 Moved Permanently"
		HttpContext.Current.Response.AddHeader("Location",uri)
		responseEnd()
	End Sub	
	
	'***********************************************************************************************************
	'* Previsao do Tempo
	'***********************************************************************************************************
	Public Function previsaoTempo() as string
		Dim html, html2, data as String
		Dim dtDate As Date
		Dim erro as boolean = false
		Dim fileWeather as string = lcase("/App_Cache/") & DateTime.Now.ToString("dd") & "_" & DateTime.Now.ToString("MM") & "_" & DateTime.Now.ToString("yyyy") & ".xml"
		Dim xml as XmlDocument = new XmlDocument()
		Dim m_nodelist As XmlNodeList
		Dim m_node As XmlNode
		
		html2 = html2 & "<a class=""follow"" target=""_blank"" title=""Siga-nos no Twitter"" href=""http://twitter.com/netcampos/"">"
			html2 = html2 & "<span>Siga-nos no Twitter</span>"
		html2 = html2 & "</a>"		
		
		If not System.IO.File.Exists( mapPath(fileWeather) ) Then 
			dim remoteFile as string = "http://servicos.cptec.inpe.br/XML/cidade/1219/previsao.xml"
			try
				xml.Load(remoteFile)
				xml.Save(mapPath(fileWeather))
			catch ex as exception
				erro = true
			end try
		end if
		
		If not System.IO.File.Exists( mapPath(fileWeather) ) Then
			erro = true
		else
			html = ""
			try
				xml.Load(mapPath(fileWeather))
				m_nodelist = xml.SelectNodes("/cidade/previsao")
				if m_nodelist.Count > 0 then
					
					html = html & "<div id=""gnc_slider_weather"">"
						For Each m_node In m_nodelist
							data = m_node.ChildNodes.Item(0).InnerText
							dtDate = DateTime.Parse(data, System.Globalization.CultureInfo.CreateSpecificCulture("pt-BR"))
							data = dtDate.ToString("dd") & "/" & dtDate.ToString("MM")
							
							html = html & "<div class=""inner"">"
								html = html & "<p class=""txt-tempo"">Previsão do Tempo - dia " & data & "</p>"
								html = html & "<img width=""111"" height=""74"" src=""" & domainAssets() & "/img/weather/cptec/" & m_node.ChildNodes.Item(1).InnerText & ".png"" title=""" & traduzTempo(m_node.ChildNodes.Item(1).InnerText) & """ alt=""" & m_node.ChildNodes.Item(1).InnerText & """ />"
								html = html & "<div class=""previsao"">"
									html = html & "<div class=""max""><div class=""ico_max spritePortal""></div><div class=""txt_max"">Max: " & m_node.ChildNodes.Item(2).InnerText & "<span class=""txtgraus"">C</span></div> </div>"
									html = html & "<div class=""min""><div class=""ico_min spritePortal""></div><div class=""txt_min"">Min: " & m_node.ChildNodes.Item(3).InnerText & "<span class=""txtgraus"">C</span></div> </div>"
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
	
	'***********************************************************************************************************
	'* description 
	'***********************************************************************************************************
	Public shared Function traduzTempo(byval str as string) as string
		Dim html as string = ""
			select case str
				case "ec"
					html = "Céu totalmente encoberto com chuvas em algumas regiões, sem aberturas de sol."
				case "ci"
					html = "Muitas nuvens com curtos períodos de sol e chuvas em algumas áreas."
				case "c"
					html = "Muitas nuvens e chuvas periódicas."
				case "in"
					html = "Nebulosidade variável com chuva a qualquer hora do dia."
				case "pp"
					html = "Nebulosidade variável com pequena chance (inferior a 30%) de pancada de chuva."
				case "cm"
					html = "Chuva pela manhã melhorando ao longo do dia."
				case "cn"
					html = "Nebulosidade em aumento e chuvas durante a noite."
				case "pt"
					html = "Predomínio de sol pela manhã. À tarde chove com trovoada."
				case "pm"
					html = "Chuva com trovoada pela manhã. À tarde o tempo abre e não chove."
				case "np"
					html = "Muitas nuvens com curtos períodos de sol e pancadas de chuva com trovoadas."
				case "pc"
					html = "Chuva de curta duração e pode ser acompanhada de trovoadas a qualquer hora do dia."
				case "pn"
					html = "Sol entre poucas nuvens."
				case "cv"
					html = "Muitas nuvens e chuva fraca composta de pequenas gotas d´ água."
				case "ch"
					html = "Nublado com chuvas contínuas ao longo do dia."
				case "t"
					html = "Chuva forte capaz de gerar granizo e ou rajada de vento, com força destrutiva (Veloc. aprox. de 90 Km/h) e ou tornados."
				case "ps"
					html = "Sol na maior parte do período."
				case "e"
					html = "Céu totalmente encoberto, sem aberturas de sol."
				case "n"
					html = "Muitas nuvens com curtos períodos de sol."
				case "cl"
					html = "Sol durante todo o período. Ausência de nuvens."
				case "nv"
					html = "Gotículas de água em suspensão que reduzem a visibilidade."
				case "g"
					html = "Cobertura de cristais de gelo que se formam por sublimação direta sobre superfícies expostas cuja temperatura está abaixo do ponto de congelamento."
				case "ne"
					html = "Vapor de água congelado na nuvem, que cai em forma de cristais e flocos."
				case "nd"
					html = "Não definido."
				case "pnt"
					html = "Chuva de curta duração podendo ser acompanhada de trovoadas à noite."
				case "psc"
					html = "Nebulosidade variável com pequena chance (inferior a 30%) de chuva."
				case "pcm"
					html = "Nebulosidade variável com pequena chance  (inferior a 30%) de chuva pela manhã."
				case "pct"
					html = "Nebulosidade variável com pequena chance  (inferior a 30%) de chuva pela tarde."
				case "pcn"
					html = "Nebulosidade variável com pequena chance  (inferior a 30%) de chuva à noite."
				case "npt"
					html = "Muitas nuvens com curtos períodos de sol e pancadas de chuva com trovoadas à tarde."
				case "npn"
					html = "Muitas nuvens com curtos períodos de sol e pancadas de chuva com trovoadas à noite."
				case "ncn"
					html = "Muitas nuvens com curtos períodos de sol com pequena chance (inferior a 30%) de chuva à noite."
				case "nct"
					html = "Muitas nuvens com curtos períodos de sol com pequena chance (inferior a 30%) de chuva à tarde."
				case "ncm"
					html = "Muitas nuvens com curtos períodos de sol com pequena chance (inferior a 30%) de chuva pela manhã."
				case "npm"
					html = "Muitas nuvens com curtos períodos de sol e chuva com trovoadas pela manhã."
				case "npp"
					html = "Muitas nuvens com curtos períodos de sol com pequena chance (inferior a 30%) de chuva a qualquer hora do dia."
				case "vn"
					html = "Períodos curtos de sol intercalados com períodos de nuvens."
				case "ct"
					html = "Nebulosidade em aumento e chuvas a partir da tarde."
				case "ppn"
					html = "Nebulosidade variável com pequena chance  (inferior a 30%) de chuva à noite."
				case "ppt"
					html = "Nebulosidade variável com pequena chance  (inferior a 30%) de chuva pela tarde."
				case "ppm"
					html = "Nebulosidade variável com pequena chance  (inferior a 30%) de chuva pela manhã."										
			end select
		return html
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Function getSessionID() as String
		dim retorno as string = HttpContext.Current.Session.SessionID		
		return retorno
	End Function			
			
End class