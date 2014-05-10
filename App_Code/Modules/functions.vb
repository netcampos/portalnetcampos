Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.IO
Imports System.Configuration
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography
Imports System.Web
Imports System.Web.Configuration
Imports System.Web.Management
Imports System.Web.Security
Imports System.Web.Caching
Imports System.Threading
Imports System.Globalization
Imports System.Xml
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Sockets

Public Module _Functions

		Public function error_404(ByVal erro as string) as string
			if debug then
				HttpContext.Current.Response.Write(erro)
				HttpContext.Current.Response.StatusCode = 404
				HttpContext.Current.Response.end()
			else
				HttpContext.Current.Response.redirect("/notfound/")
				HttpContext.Current.response.End()
			end if
		End function
		
		Public function error_SQL(ByVal erro as string) as string
			if debug then
				HttpContext.Current.Response.Write(erro)
				HttpContext.Current.Response.StatusCode = 503
				HttpContext.Current.Response.end()
			else
				HttpContext.Current.Response.Write("<div style=""margin:0 auto;text-align:left;width:600px;""><h2 style=""color:#69D2E7;font-size:1.6em;font-weight:normal;letter-spacing:-0.06em;"">ManutenÁ„o Tempor·ria...</h2><p style=""color:#888888;font-size:0.9em;"">Estamos realizando manutenÁ„o tÈcnica em nosso sistema, em instantes estaremos de volta novamente.</p></div>")
				HttpContext.Current.Response.StatusCode = 503
				HttpContext.Current.response.End()
			end if
		End function			
		
				
		Public Function get_md5(ByVal SourceText As String) As String
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
		
		Public Function cortaTexto(ByVal Texto as String, qtd as Integer) as String
			dim aux as string
			if len(trim(texto)) > qtd Then 
				Aux = mid(texto, 1, qtd )
				Aux = rtrim(Aux) & "..."
			else
				Aux = texto
			end if
			return Aux
		End Function
		
		'verifica se imagem e retorna caso exista apenas para imagens representativas n„o membros
		Public function fotoLegenda(ByVal img as String, ByVal titulo as String, ByVal href as String, ByVal largura as Integer, ByVal altura as Integer) as string
			dim html as string = String.Empty
			If not IsDBNull(img) and img <> "" then
				html = "<div class=""foto-legenda""><a href=""" & href & """ class=""borda-interna"" title=""" & titulo & """><img src=""http://images.guiadoturista.net/?img=" & img & "&amp;l=" & largura & "&amp;a=" & altura & """ width=""" & largura & """ height=""" & altura & """ /></a></div>"
			end if
			return html
		End function
		
		'verifica se imagem e retorna com caminho tratada pelo guiadoturista
		Public function fotoLegenda2(ByVal img as String, ByVal largura as Integer, ByVal altura as Integer) as string
			dim html as string = String.Empty
			If not IsDBNull(img) and img <> "" then
				html = "http://images.guiadoturista.net/?img=" & img & "&l=" & largura & "&a=" & altura
			end if
			return html
		End function
		
		'verifica se imagem e retorna com caminho tratada pelo guiadoturista sem dimensoes
		Public function thumbNail(ByVal img as String) as string
			dim html as string = String.Empty
			If not IsDBNull(img) and img <> "" then
				html = "http://images.guiadoturista.net/?img=" & img
			end if
			return html
		End function		
		
		'fotoLeganda New Window
		Public function fotoLegendaNewWindow(ByVal img as String, ByVal titulo as String, ByVal href as String, ByVal largura as Integer, ByVal altura as Integer) as string
			dim html as string = String.Empty
			If not IsDBNull(img) and img <> "" then
				html = "<div class=""foto-legenda""><a href=""" & href & """ class=""borda-interna"" title=""" & titulo & """ target=""_blank""><img src=""http://images.guiadoturista.net/?img=" & img & "&amp;l=" & largura & "&amp;a=" & altura & """ width=""" & largura & """ height=""" & altura & """ /></a></div>"
			end if
			return html
		End function				
		
		'funcao para remover ascento de uma string
		Public Function removeascento(ByVal texto as String) as string
			Dim str1 as String
			Dim str2 as String
			Dim i as Integer
			texto = trim(texto)
			texto = replace(texto, "-", "")
			texto = replace(texto, "  ", " ")
			texto = replace(texto, "   ", " ")
			texto = replace(texto, "∫", "")
			texto = replace(texto, "™", "")
			i = 0
			str1 = "ƒ≈¡¬¿√‰·‚‡„… À»ÈÍÎËÕŒœÃÌÓÔÏ÷”‘“’ˆÛÙÚı‹⁄€¸˙˚˘«Á"
			str2 = "AAAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUuuuuCc"
			Do While i < Len(str1)
				i = i + 1
				texto = Replace(texto, Mid(str1, i, 1), Mid(str2, i, 1))  
			loop
			texto = replace(texto, "  ", " ")
			texto = replace(texto, ",", "")
			texto = replace(texto, ".", "")
			texto = replace(texto, "!", "")
			texto = replace(texto, "?", "")
			texto = replace(texto, "__", "_")
			return lcase((texto))
		end Function
		
		'Remove Espaco
		Public Function removeEspacao(byVal texto as string) as string
			texto = trim(texto)
			dim contador as integer = instr( 1,texto,chr(32) )
			dim i as integer
			for  i = 1 to contador 
				texto = replace(texto, chr(32), "")
				i = i + 1
			next
			return texto 
		end function
		
		'funcao para remover ascento de uma string
		Public Function tiraAscento(ByVal texto as String) as string
			Dim str1 as String
			Dim str2 as String
			Dim i as Integer
			texto = trim(texto)
			i = 0
			str1 = "ƒ≈¡¬¿√‰·‚‡„… À»ÈÍÎËÕŒœÃÌÓÔÏ÷”‘“’ˆÛÙÚı‹⁄€¸˙˚˘«Á"
			str2 = "AAAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUuuuuCc"
			Do While i < Len(str1)
				i = i + 1
				texto = Replace(texto, Mid(str1, i, 1), Mid(str2, i, 1))  
			loop
			return lcase((texto))
		end Function		
		
		Public Function GenerateSlug(phrase As String, maxLength As Integer) As String
			Dim str As String = phrase.ToLower()
			str = Regex.Replace(str, "[^a-z0-9\s-]", "")
			str = Regex.Replace(str, "[\s-]+", " ").Trim()
			str = str.Substring(0, If(str.Length <= maxLength, str.Length, maxLength)).Trim()
			str = Regex.Replace(str, "\s", "-")
			Return str
		End Function		
		
		'Limpa Tags HTML
		Function StripHTMLTags(ByVal html As String) As String
			Dim conteudo As String
			If html <> "" Then
				conteudo = Regex.Replace(html, "<(.|\n)+?>", String.Empty)
				'if stripped.length >= total then
					'stripped = stripped.substring(0,total)
				'end if
				Return conteudo
			Else
				Return ""
			End If
		End Function		
		
		'funcao para embed swf
		Public Function embedSWF(ByVal src as String, ByVal width as Integer, ByVal height as Integer, ByVal FlashVars as String) as string
			dim swf as String = String.Empty
			swf = "<object width=""" & width & """ height=""" & height & """>"
			swf = swf & "<param value=""" & src & """ name=""movie""></param>"
			swf = swf & "<param value=""transparent"" name=""wmode""></param>"
			swf = swf & "<param name=""quality"" value=""autohigh""></param>"
			swf = swf & "<param name=""bgcolor"" value=""#FFFFFF""></param>"
			if FlashVars <> "" then swf = swf & "<param name=""flashvars"" value=""" & FlashVars & """></param>"
			swf = swf & "<param name=""allowfullscreen"" value=""true""></param>"
			swf = swf & "<param name=""allowScriptAccess"" value=""always""></param>"
			swf = swf & "<param name=""quality"" value=""high""></param>"
			swf = swf & "<embed src=""" & src & """ type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" wmode=""transparent"""
			if FlashVars <> "" then swf = swf & " flashvars=""" & FlashVars & """"
			swf = swf & " width=""" & width & """ height=""" & height & """ allowFullScreen=""true"" bgcolor=""#FFFFFF"" allowScriptAccess=""always"" quality=""high""></embed>"
			swf = swf & "</object>"			
			return swf	
		End function
		
		'funcao para embed swf
		Public Function embedPlayer(ByVal src as String, ByVal width as Integer, ByVal height as Integer, ByVal FlashVars as String, ByVal nomeObjeto as String) as string
			dim swf as String = String.Empty
			swf = "<object width=""" & width & """ height=""" & height & """ codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" id=""" & nomeObjeto & """>"
			swf = swf & "<param value=""" & src & """ name=""movie""></param>"
			swf = swf & "<param value=""transparent"" name=""wmode""></param>"
			swf = swf & "<param name=""quality"" value=""autohigh""></param>"
			swf = swf & "<param name=""bgcolor"" value=""#FFFFFF""></param>"
			if FlashVars <> "" then swf = swf & "<param name=""flashvars"" value=""" & FlashVars & """></param>"
			swf = swf & "<param name=""allowfullscreen"" value=""true""></param>"
			swf = swf & "<param name=""allowScriptAccess"" value=""always""></param>"
			swf = swf & "<param name=""quality"" value=""high""></param>"
			swf = swf & "<embed src=""" & src & """ type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" wmode=""transparent"""
			if FlashVars <> "" then swf = swf & " flashvars=""" & FlashVars & """"
			swf = swf & " width=""" & width & """ height=""" & height & """ allowFullScreen=""true"" bgcolor=""#FFFFFF"" allowScriptAccess=""always"" quality=""high""></embed>"
			swf = swf & "</object>"			
			return swf	
		End function		
		
		'################# funcoes para trabalhar com previsao do tempo weather channel ########################'
		Public function traduz_condicao(ByVal txt as string) as String
			select case ucase(txt)
				case "FAIR"
					return "Tempo Bom"
				case "CLEAR"
					return "CÈu Claro"
				case "MOSTLY CLEAR"
					return "Parcialmente Claro"
				case "SUNNY"
					return "Ensolarado"
				case "MOSTLY SUNNY"
					return "Predomin‚ncia de Sol"
				case "RAIN"
					return "Chuva"
				case "LIGHT DRIZZLE"
					return "Chuvisco Fraco"	
				case "LIGHT RAIN"
					return "Chuva Fraca"			
				case "CLOUDY"
					return "Nublado"		
				case "PARTLY CLOUDY"
					return "Parcialmente Nublado"		
				case "MOSTLY CLOUDY"
					return "Encoberto"		
				case "T-STORMS"
					return "Trovoadas"		
				case "SCATTERED T-STORMS"
					return "Trovoadas Esparsas"		
				case "ISOLATED T-STORMS"
					return "Trovoadas Isoladas"		
				case "PM T-STORMS"
					return "Trovoadas a tarde"		
				Case "SHOWERS"
					return "Pancadas"		
				case "SHOWERS EARLY"
					return "Pancadas"		
				case "SCATTERED SHOWERS"
					return "Pancadas Esparsas"		
				case "AM SHOWERS"
					return "Pancadas de manh„"
				case "PM SHOWERS"
					return "Pancadas de tarde"
				case "THUNDER"
					return "Trovoada"
				case ucase("Light Rain with Thunder")
					return "Chuva e trovoada"
			end select
		end function
		
		Public function data_update_wheather(ByVal txt as string) as string
			txt = replace (txt, "Local Time", "")
			txt = "#" & rtrim(txt) & "#"
			Dim data As Date = Date.Parse(txt, System.Globalization.CultureInfo.InvariantCulture) 
				txt = data.toString(System.Globalization.CultureInfo.CreateSpecificCulture("pt-BR"))
			return txt
		
		End function
		
		'##################### Fim funcoes weather ####################################################'		
		
		'Remove Padding do valor inidicado em qtd(ultima coluna)
		Public Function noPadding(ByVal v as Integer, ByVal qtd as Integer) as String
			if v mod qtd = 0 then
				return"<li class=""primeiro"">"
			else if (v+1) mod qtd = 0 then
				return "<li class=""ultimo"">"
			else
				return "<li>"
			end if 
		end function
		
		Public function noPaddingClass(ByVal v as Integer, ByVal qtd as Integer) as String
			if v = qtd then
				return "class=""sem-padding"""
			end if
		end function		
		
		Public Function listaClear(ByVal v as Integer, ByVal qtd as Integer)
			if (v+1) mod qtd = 0 then
				return "<li class=""clear""><strong></strong></li>"
			end if 
		End function
		
		'################### Funcoes para tratamento de Dadtas #####################'
		'funcao para dataBr retorna zero a esquerda
		Public function databr(data As Date) as string
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
		
		'funcao para retornar data em microformats
		public function datamicroformats(data As Date) as string
		Dim out as String = String.Empty
			out = month(data)
			out = year(data) & "-" & month(data) & "-" & day(data) & "T" & hour(data) & ":" & minute(data) & ":00"
			return out
		end function
		
		Public function dataPublishLogDate(ByVal data As Date) as string
			Dim dtTweet As Date = data
			return FormatDateTime(data, DateFormat.LongDate) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime)
		End function		
		
		'retorn data com zero a esquerda e horas ex 20/08/1978 - 22h30
		public function datamateria(data As Date) as string
			dim out, dia, mes, ano, hora, minuto as string 
			dia = day(data)
			mes = month(data)
			ano = year(data)
			hora = hour(data)
			minuto = minute(data)
			if dia < 10 then
				dia = "0" & dia
			end if
			if mes < 10 then
				mes = "0" & mes
			end if
			if hora < 10 then
				hora = "0" & hora
			end if
			if minuto < 10 then
				minuto = "0" & minuto
			end if
			out = dia & "/" & mes & "/" & right(ano,2) & " - " & hora & "h" & minuto
			return out
		end function
		
		public function dataSortable(ByVal data as date) as String
			dim data1 as Date = data
			Dim dateOffset As New DateTimeOffset(data1, TimeZoneInfo.Local.GetUtcOFfset(data1))
			return data1.ToString("s")
		end Function
		
		'retorna data modelo twitter com o momento literal
		Public function dataMomento(ByVal data as date) as string
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
				return "Um instante atr·s"
			end if
			
			If tempo <= TimeSpan.FromMinutes(60) Then
				if tempo.Minutes > 1  then 
					return tempo.Minutes & " minutos atr·s"
				else
					return "Um minuto atr·s"
				end if
			end if	
			
			If tempo <= TimeSpan.FromHours(24) Then
				if tempo.Hours > 1  then 
					return tempo.Hours & " horas atr·s"
				else
					return "Hoje ·s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & " - " & tempo.Minutes & " minutos atr·s"
				end if
			end if
			
			If tempo <= TimeSpan.FromDays(30) Then
				Dim semanas as Integer = DateDiff(DateInterval.Weekday, dtTweet, dtAgora) 

				if tempo.Days > 1  then
					if tempo.Days < 7 then
						return tempo.Days & " dias atr·s - " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
					else
						if semanas > 1 then return semanas & " semanas atr·s em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
						return "semana passada em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
					end if
				else
					return "OntÈm - " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
				end if
			end if	
			
			If tempo <= TimeSpan.FromDays(365) Then
				if tempo.Days > 30  then
					dif = New DateTime((DateTime.Today - dtTweet).Ticks).Month
					if dif > 1 then return dif & " meses atr·s em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
					return dif & " mÍs atr·s em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
				else
					return "MÍs passado - " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
				end if
			end if				
			
			if tempo.Days > 365  then
				dif = New DateTime((DateTime.Today - dtTweet).Ticks).Year - 1

				if dif > 1 then return dif & " anos atr·s, em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " de " & dtTweet.Year & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime)				
				return dif & " ano atr·s, em " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " de " & dtTweet.Year & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime)
				
			else
				return "Um ano atr·s - " & dtTweet.ToString("M", new CultureInfo("pt-BR")) & " (" & dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & ")"
			end if
			
			'return exibir
		end function				
		
		'retorna data modelo twitter com o momento literal
		Public function dataTweet(ByVal data as date) as string
			Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
			Thread.CurrentThread.CurrentCulture = New CultureInfo("pt-BR")
		
			Dim dtTweet As Date = data
			Dim dtAgora As Date = DateTime.Now
			Dim tempo As TimeSpan = dtAgora - dtTweet
			Dim exibir as string = ""
			
			If tempo <= TimeSpan.FromSeconds(60) Then
				return "Um instante atr·s"
			end if
			
			If tempo <= TimeSpan.FromMinutes(60) Then
				if tempo.Minutes > 1  then 
					return tempo.Minutes & " minutos atr·s"
				else
					return "Um minuto atr·s"
				end if
			end if	
			
			If tempo <= TimeSpan.FromHours(24) Then
				if tempo.Hours > 1  then 
					return tempo.Hours & " horas atr·s"
				else
					return "¡s " & FormatDateTime(dtTweet, DateFormat.ShortTime) & " - " & tempo.Minutes & " minutos atr·s"
				end if
			end if
			
			If tempo <= TimeSpan.FromDays(30) Then
				if tempo.Days > 1  then
					if tempo.Days < 7 then
						return dtTweet.ToString("dddd", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime)
					else
						return dtTweet.ToString("M", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime)
					end if
				else
					return "OntÈm ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime)
				end if
			end if	
			
			If tempo <= TimeSpan.FromDays(365) Then
				if tempo.Days > 30  then 
					return dtTweet.ToString("M", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime)
				else
					return "Um mÍs atr·s "
				end if
			end if				
			
			if tempo.Days > 365  then 
				return dtTweet.ToString("Y", new CultureInfo("pt-BR")) & " ‡s " & FormatDateTime(dtTweet, DateFormat.ShortTime)
			else
				return "Um ano atr·s "
			end if
			
			'return exibir
		end function
		
			
		'funcao para verificar se existe um template existente no webconfig utilizado em webtemplate ou no banco de dados		
		Public function iif_template(ByVal webTemplate as string, ByVal SystemTemplate as string) as string
			if not(webTemplate = nothing) then
				return webTemplate
			else
				return SystemTemplate	
			end if
		End function
		
		
		'caminho da url pagina de busca da cidade para formulario
		Public Function urlBusca(ByVal idcidade As Integer) As String
			Dim conn As SqlConnection
			Dim cmd As SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			Try
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand("select c.url_amigavel from tbl_config_modulos_netportal as c inner join tbl_modulos_netportal as m on c.id_modulo = m.id where c.id_cidade = @id_cidade and c.id_modulo = 23", conn)
				cmd.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				cmd.Parameters("@id_cidade").Value = idcidade
				dr = cmd.ExecuteReader()
				If dr.HasRows Then
					dr.Read()
					Return dr(0).toString()
				End If
				dr.close()
			Catch ex As exception
	
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()
				If conn IsNot Nothing Then conn.Dispose()
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try
		End Function
		
		'Funcao Gera CÛdigo para Imagem Captcha
		Public Function GeraCodigoCaptCha() As String
			Dim s As String = ucase(Left(System.Guid.NewGuid.ToString, 6))
			return s
		End Function
		
		'################FunÁıes para ValidaÁ„o e Envio de Email #################################'
		
		'valida com expressao regular
		Public Function validaEmail(ByVal emailAddress As String) As Boolean
			Dim pattern As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
			Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
			If emailAddressMatch.Success Then
				return validaEmailDns(emailAddress)
			Else
				return False
			End If
		End Function
		
		'valida com expressao regular
		Public Function validaEmailSingle(ByVal emailAddress As String) As Boolean
			Dim pattern As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
			Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
			If emailAddressMatch.Success Then
				return true
			Else
				return False
			End If
		End Function 		 
		
		'resolve o host do email informado
		Public Function validaEmailDns(ByVal emailAddress as String ) as Boolean
			try
				Dim host as string() = emailAddress.Split(CChar("@"))
				Dim hostName as String = host(1)
				Dim hostInfo As IPHostEntry = Dns.Resolve(hostName)
				Dim endpoint As New IPEndPoint(hostInfo.AddressList(0), 25)
				Dim s As Socket = New Socket(endpoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp)
				s.Connect(endpoint)
				If s.Connected Then 
					s.close()
					return true
				end if
			Catch ex As Exception
				return false
			End try
		End function
		
		'envia email (Assunto, fromNome, FromEmai, toNome, toEmail, msg)
		Public Function enviarEmail(ByVal Assunto as String, ByVal fromNome as String, ByVal fromEmail as String, ByVal toNome as String, ByVal toEmail as String, ByVal msg as String) as Boolean
			Dim netmail As New System.Net.Mail.MailMessage()
			Dim netmailsmtp As New System.Net.Mail.SmtpClient			

			Try
				with netmail
					.From = New System.Net.Mail.MailAddress("" & fromNome & " <" & fromEmail & ">")
					.To.Add("<" & toEmail & ">")
					.Sender = New System.Net.Mail.MailAddress("<" & fromEmail & ">")		
					.Priority = System.Net.Mail.MailPriority.Normal
					.IsBodyHtml = True
					.Subject = Assunto
					.Body = msg 
					.Headers.Add("X-Mailer", "NetMail .NET (powered by Microsoft Windows 2003)")
					.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
					.SubjectEncoding = System.Text.Encoding.GetEncoding("utf-8")
					.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8")
				end with
				
				with netmailsmtp
					.Host = "localhost"
					.send(netmail)
				end with				
				
			Catch ex As Exception
				return false
				exit function
			End Try
			
			netmail.Dispose()				
			return false
			
		End Function
		
		'################################## Fim Validacao e Envio de Email ############################'
		
		
		'################################## Load JS File ######################################'
		
		public function loadJSFile(byVal url as String, _charset as String) as String
			dim js as string = string.empty
				js = js & "<script type=""text/javascript"" src=""" & url & """></"
				js = js & "script>"			
			return js
		end Function
		
		'Funcao para tratamento de post html em formularios
		Public Function MakeLinks(StringValue As String)
			Dim uriTube as String = String.Empty
			Dim strContent As String = StringValue
			Dim urlregex As Regex = New Regex("(?<!http://)(?<!https://)(www\.([\w.]+\/?)\S*)", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))
			strContent = urlregex.Replace(strContent, "http://$1")	
			Dim urlregex2 As Regex = New Regex("(http:\/\/([\w.]+\/?)\S*|https:\/\/([\w.]+\/?)\S*)", (RegexOptions.IgnoreCase Or RegexOptions.Compiled))
			strContent = urlregex2.Replace(strContent, "<a href=""$1"" target=""_blank"">$1</a>")			
			
			'If (strContent.IndexOf("http://www.youtube.com/watch?v=") <> -1) Then
				'uriTube = urlregex2.Replace(strContent, "<a href=""$1"" target=""_blank"">$1</a>")
				'strContent = replace(strContent,"http://www.youtube.com/watch?v=","http://www.youtube.com/v/")
				'strContent = urlregex2.Replace(strContent, "<br><div>" & embedSWF("/App_Modules/player/player.swf", 350, 240, "file=$1") & "<br><span style=""font-size:11px;"">Link Original: " & uriTube & "</span></div>")
			'else
				
			'end if
			
			Return strContent
		End Function
		
		'FunÁ„o para Star Rating Conte˙do
		Public function starRatingYoutube(v as string)
			if v.toString() <> "" then
				dim t as double
				t = replace(v,".",",")
				v = Math.Round(t).toString()
			else
				v = 0
			end if	
			return v
		end function
		
		Public Function clearStringComments(ByVal StringValue as String) as String
			Dim url As String = StringValue
			dim stripped as string = Regex.Replace(StringValue,"<(.|\n)*?>",string.Empty)
			return HttpUtility.HtmlAttributeEncode(stripped)
		End Function
		
		'Abre P·gina De Conteudo Extraido em Kernel
		Public Sub loadContentPage(byVal area as String, ByVal idConteudo as String, ByVal idModulo as Integer)
			
			Dim pathFileBase as String = httpContext.Current.server.mappath("/App_Themas/" & nome_template & "/paginas/" & area & "/")
			Dim arquivo as string = ("/App_Themas/" & nome_template & "/paginas/" & area & "/pagina.aspx?idconteudo=" & idconteudo & "&idmodulo=" & idModulo)
			Dim caminho as String = arquivo		
			
			if inStr(arquivo, "?") > 0 then
				arquivo = left(arquivo, inStr(arquivo, "?") - 1)
			end if	
			
			arquivo = lcase(httpContext.Current.server.mappath(arquivo))
			
			If System.IO.File.Exists(arquivo) Then	
				execPagina(lcase(caminho))
				Exit Sub
			Else
				If System.IO.File.Exists(pathFileBase & "home.aspx") Then
					System.IO.File.Copy(pathFileBase & "home.aspx", arquivo)
					System.IO.File.Copy(pathFileBase & "home.aspx.vb", replace(arquivo,"aspx","aspx.vb"))	
				else
					error_404("Arquivo Mestre N„o Definido - ImpossÌvel Criar P·gina Filha")
				end if
				error_404("P·gina:" & caminho & ", N„o Encontrada")
				Exit Sub
			End If
			
		End Sub
		
		'Extrai a url da p·gina de Conteudo
		Public Function _getURLContentPage(ByVal pagina as String) as String
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
		
		
       Public Function EncodeBase64String(ByVal Str As String) As String
            Dim enc As System.Text.Encoding = System.Text.Encoding.ASCII
            Dim ByteArray As Byte()
            ByteArray = enc.GetBytes(Str)
            Return System.Convert.ToBase64String(ByteArray)
        End Function
		
end Module

'################ Retorna Coordenadas GeoReferenciadas de um determiando endereco #################################'
Public Class _getGMaps
	Public idCidade as Integer
	Public endereco as String
	Public m_Latitude as String
	Public m_Longitude as String
	Public m_Address as String	
	
	Public Function _renderMapa() as String
		Dim url as string = "http://maps.google.com/maps/geo?q=" & HttpContext.Current.server.UrlEncode(endereco) & "&output=xml&sensor=false&key=" & city.dados(idCidade, "api_google")
		Dim reader as XmlTextReader  = new XmlTextReader(url)
		reader.WhitespaceHandling = WhitespaceHandling.Significant
		Dim xml as XmlDocument= new XmlDocument()
		xml.Load(reader)
		try
			Dim coordenadas as string = xml.GetElementsByTagName("coordinates")(0).InnerText
			Dim addr as String = xml.GetElementsByTagName("address")(0).InnerText
			
			If InStr(1, coordenadas, ",") Then
				dim nn() as String
				nn = coordenadas.split(",")
				Longitude = nn(0)
				Latitude = nn(1)
				address = addr
			End If
			
		catch ex as exception
			HttpContext.Current.response.write(ex.Message())
		end try
		
	End Function
	
	Public property Latitude() as String
        Get
			return m_latitude
        End Get	
        Set(ByVal value As String)
			m_Latitude = value
        End Set		
	End Property
	
	Public property Longitude() as String
        Get
			return m_Longitude
        End Get	
        Set(ByVal value As String)
			m_Longitude = value
        End Set
	End Property
	
	Public property Address() as String
        Get
			return m_Address
        End Get	
        Set(ByVal value As String)
			m_Address = value
        End Set
	End Property	
	
End Class






