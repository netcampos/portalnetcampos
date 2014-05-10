Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail
Imports System.Configuration
Imports System.Text
Imports System.Security.Cryptography
Imports System.Web
Imports System.Web.Configuration
Imports System.Web.Management
Imports System.Web.Security
Imports System.Web.Caching
Imports Mysql.Data
Imports Mysql.Data.MySqlClient
Public Class _sendMail

	Public idCidade as String
	Public nomeFrom as string
	Public emailFrom as string
	Public nomeReplyTo as String	
	Public replyTo as string
	Public nomeTo as string	
	Public emailTo as string
	Public assunto as String
	Public corpoHTML as String
	Public messageID as String
		
	Public Sub followNotify(ByVal memberFrom as String, ByVal memberTo as string)
		messageID = get_md5(DateTime.Now.toString())	
		if not(_checkNotifyFollow(memberFrom, memberTo) ) then
			'Corpo do Email
			Dim urlPerfil as String = city.dados(idCidade, "urlcidade") & "/membros/" & dados_user(memberFrom, "id") & "," & GenerateSlug(tiraAscento( dados_user(memberFrom, "nome") ),255) & ".html"
			Dim logoEmail as String = city.dados(idCidade, "pasta_padrao") & "templates/" & city.dados(idCidade, "ntemplate") & "/" & city.dados(idCidade, "template") & "/images/logo-email.png"
			Dim corTemplate as String = city.dados(idCidade, "template")
			Dim bg as String = "#cede9d"
			Dim bgBox as String = "#f3f6ea"
			Dim bgBtn as String = "#003333"
			Dim borda as String = "#cccccc"			
			select case corTemplate
				case "verde"
				bg = "#cede9d"
				case Else
				bg = bg
			end select
			Dim htmlEmail as String = String.Empty
			
			htmlEmail = htmlEmail & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional //EN""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""><title>" & city.dados(idCidade, "nome_amigavel") & "</title></head><body style=""margin: 0; padding: 0;"" dir=""ltr"">"
			htmlEmail = htmlEmail & "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""40""><tr><td bgcolor=""#f7f7f7"" width=""100%"" style=""font-family: 'lucida grande', tahoma, verdana, arial, sans-serif;"">"
			htmlEmail = htmlEmail & "<table width=""620"" border=""0"" cellSpacing=""0"" cellPadding=""0"" style=""border:solid 1px " & bg & ";"">"
				htmlEmail = htmlEmail & "<tr><td style=""background: " & bg & ";padding: 4px 4px 4px 10px""><a href=""" & city.dados(idCidade, "urlCidade") & """><img src=""" & logoEmail & """ style=""border:0;display:block;""/></a></td></tr>"
				htmlEmail = htmlEmail & "<tr><td style=""padding: 15px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;font-size:12px;background:#ffffff;"">"
					htmlEmail = htmlEmail & "<table width=""100%""><tr>"
						'Coluna Esquerda
						htmlEmail = htmlEmail & "<td width=""440px"" style=""font-size: 12px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"" valign=""top"" align=""left"">"
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 15px 0;font-size: 13px;"">"
							htmlEmail = htmlEmail & "Ol&aacute; " & dados_user(memberTo, "nome") & ","
							htmlEmail = htmlEmail & "</div>"
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 15px 0;"">"
							htmlEmail = htmlEmail & dados_user(memberFrom, "nome") & " " & dados_user(memberFrom, "sobreNome") & " passou a seguir suas atualiza&ccedil;&otilde;es no " & city.dados(idCidade, "nome_amigavel")
							htmlEmail = htmlEmail & "</div>"
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 15px 0;"">"
								htmlEmail = htmlEmail & "<div style=""border-bottom: solid 1px #ccc; line-height:5px"">&nbsp;</div><br />"
								htmlEmail = htmlEmail & "<table style=""margin: 5px 0 0 0"" cellPadding=""0"">"
									htmlEmail = htmlEmail & "<tr vAlign=""top"">"
										htmlEmail = htmlEmail & "<td style=""padding: 0 3px 10px 0;""><a href=""" & urlPerfil & """><img src=""" & foto_user(memberFrom, 50, 50) & """ width=""50"" height=""50"" style=""border:solid 1px #cccccc;display:block;padding:5px;""/></a></td>"
										htmlEmail = htmlEmail & "<td style=""padding: 0 0 10px 0;font-size: 11px;color:#999;""><a href=""" & urlPerfil & """ style=""color: #3b5998; font-size: 13px; font-weight: bold; text-decoration: none"" >" & dados_user(memberFrom, "nome") & " " & dados_user(memberFrom, "sobrenome") & "</a></td>"
									htmlEmail = htmlEmail & "</tr>"
								htmlEmail = htmlEmail & "</table>"
								htmlEmail = htmlEmail & "<div style=""border-bottom: solid 1px #ccc; line-height:5px"">&nbsp;</div><br />"
								htmlEmail = htmlEmail & "<div style=""margin: 0px"">Atenciosamente,<br/><br/>Equipe do " & city.dados(idCidade, "nome_amigavel") & "<br/>" & city.dados(idCidade, "cidade") & "/" & city.dados(idCidade, "uf") & "</div>"								
							htmlEmail = htmlEmail & "</div>"							
						htmlEmail = htmlEmail & "</td>"
						
						'Coluna Direita
						htmlEmail = htmlEmail & "<td style=""padding: 0 10px 0 15px;"" vAlign=""top"" width=""150"" align=""left"">"
							htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
								htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:10px; background:" & bgBox & """>"
									htmlEmail = htmlEmail & "<div style=""margin:0 0 15px 0; font-size:13px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"">Para ver o Perfil de " & dados_user(memberFrom, "nome") & " " & dados_user(memberFrom, "sobreNome") & "</div>"
									htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
										htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:0px; background:" & bgBtn & """>"
											htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
												htmlEmail = htmlEmail & "<tr><td style=""padding:6px 15px;""><a href=""" & urlPerfil & """ style=""color:#ffffff;font-size:12px;text-decoration:none;text-transform:uppercase;font-weight:bold;"">Entrar</a></td></tr>"
											htmlEmail = htmlEmail & "</table>"
										htmlEmail = htmlEmail & "</td></tr>"
									htmlEmail = htmlEmail & "</table>"
								htmlEmail = htmlEmail & "</td></tr>"
							htmlEmail = htmlEmail & "</table>"
						htmlEmail = htmlEmail & "</td>"
						
					htmlEmail = htmlEmail & "</tr></table>"

					'Lembre Box link url Perfil'
					htmlEmail = htmlEmail & "<br/>"
					htmlEmail = htmlEmail & "<table cellspacing=""0"" cellpadding=""0"" style=""border-collapse:collapse;width:100%;"">"
					htmlEmail = htmlEmail & "<tr><td style=""font-family:'lucida grande', tahoma, verdana, arial, sans-serif;padding:10px;background:" & bgBox & ";border:1px solid " & borda & """>"
						htmlEmail = htmlEmail & "<div style=""font-weight: bold; margin-bottom: 2px; font-size: 11px;"">Para visualizar o perfil de " & dados_user(memberFrom, "nome") & " " & dados_user(memberFrom, "sobreNome") & ", acesse:</div>"
						htmlEmail = htmlEmail & "<a href=""" & urlPerfil & """ style=""color:#3b5998;text-decoration:none;font-size:11px;"">" & urlPerfil & "</a>"
					htmlEmail = htmlEmail & "</td></tr>"
					htmlEmail = htmlEmail & "</table>"
											
				htmlEmail = htmlEmail & "</td></tr>"
			htmlEmail = htmlEmail & "</table>"
			
			'Rodape
			htmlEmail = htmlEmail & "<table width=""620"" border=""0"" cellSpacing=""0"" cellPadding=""0"" style=""padding:10px 0 0 0"">"
				htmlEmail = htmlEmail & "<tr><td style=""padding:10px;color:#999999;font-size:10px;font-family: 'lucida grande', tahoma, verdana, arial, sans-serif"">"
					htmlEmail = htmlEmail & "<div>Esta mensagem foi enviada para " & dados_user(memberTo, "usuario") & " pelo " & city.dados(idCidade, "nome_amigavel") & ". O " & city.dados(idCidade, "nome_amigavel") & " &eacute; um Portal s&eacute;rio, respeitado e que disponibiliza informa&ccedil;&otilde;es para um p&uacute;blico que realmente visita e se interessa por " & city.dados(idCidade, "cidade") &", para ver nossas diretrizes sobre seguran&ccedil;a de seus dados Acesse nossa <a href=""" & city.dados(idCidade, "urlCidade") & "/press-release/politica-de-privacidade.html"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">Pol&iacute;tica de Privacidade</a> e para mais informa&ccedil;&otilde;es sobre termos de uso do " & city.dados(idCidade, "nome_amigavel") & " acesse nosso <a href=""" & city.dados(idCidade, "urlCidade") & "/press-release/aviso-legal.html"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">Aviso Legal.</a></div><br/>"
					htmlEmail = htmlEmail & "<div>Por favor, n&atilde;o responda a esta mensagem, pois foi enviada de um endere&ccedil;o de email n&atilde;o monitorado. Esta mensagem &eacute; um e-mail dos servi&ccedil;os relacionados ao seu uso no " & city.dados(idCidade, "nome_amigavel") & ". Para informa&ccedil;&otilde;es gerais ou para pedir ajuda sobre sua conta no " & city.dados(idCidade, "nome_amigavel") & ", por favor acesse nosso <a href=""" & city.dados(idCidade, "urlCidade") & "/contato/"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">formul&aacute;rio de contato</a>.</div><br/>"
					htmlEmail = htmlEmail & "<div>" & city.dados(idCidade, "nome_amigavel") & "<br/>" & city.dados(idCidade, "endFranquia") & " - " & city.dados(idCidade, "bairroFranquia") & "<br/>" & city.dados(idCidade, "cepFranquia") & " - " & city.dados(idCidade, "cidadeFranquia") & " / " & city.dados(idCidade, "estadoFranquia") & "</div>"					
				htmlEmail = htmlEmail & "</td></tr>"
			htmlEmail = htmlEmail & "</table>"
			
			htmlEmail = htmlEmail & "</td></tr></table>"
			htmlEmail = htmlEmail & "<img src=""http://cms.guiadoturista.net/logSendMail/logFollow.asp?key=" & messageID & """ alt="""" style=""border: 0; height:1px; width:1px;"" />"
			htmlEmail = htmlEmail & "</body></html>"
			
			corpoHTML = htmlEmail			
			
			Dim netMail As New System.Net.Mail.MailMessage()
			Try
				with netmail
					.From = New System.Net.Mail.MailAddress("" & nomeFrom & " <" & emailFrom & ">")
					.To.Add("" & nomeTo & " <" & emailTo & ">")
					.ReplyTo = New System.Net.Mail.MailAddress("" & nomeReplyTo & " <" & replyTo & ">")		
					.Priority = System.Net.Mail.MailPriority.Normal
					.IsBodyHtml = True
					.Subject = assunto
					.Body = corpoHTML 
					.Headers.Add("X-Mailer", "NetMail .NET (powered by Microsoft Windows Server)")
					.Headers.Add("X-MessageType", 1) 'follow notification					
					.Headers.Add("X-MessageUID", messageID)
					.Headers.Add("X-MessageFrom", memberFrom)
					.Headers.Add("X-MessageTo", memberTo)
					.Headers.Add("X-CidadeId", idCidade)					
					.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
					.SubjectEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
					.BodyEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
				end with
			
				Dim netmailsmtp As New System.Net.Mail.SmtpClient
				with netmailsmtp
					.Host = "localhost"
					.send(netmail)
				end with
				
				Exit Sub
				
			Catch ex As Exception
				Exit Sub
			End Try
		End If
	
	End Sub
	
	Private Function _checkNotifyFollow(ByVal memberFrom as String, ByVal memberTo as string) as Boolean
		Dim total as Integer
		Dim conn as MySqlConnection
		Dim cmd as MySqlCommand
		Dim strCon as string = ConfigurationManager.ConnectionStrings("mysqlNetMailCidades").ConnectionString
		conn = New MySqlConnection(strCon)
		
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New MySqlCommand()
			with cmd
				.commandtext = "select count(*) from tbl_logEnvioFollow where memberFrom = @memberFrom and memberTo = @memberTO and idCidade = @idCidade"
				.connection = conn
				.Parameters.Add(New MySqlParameter("@memberFrom", MySqlDbType.String))
				.Parameters("@memberFrom").Value = memberFrom
				.Parameters.Add(New MySqlParameter("@memberTO", MySqlDbType.String))
				.Parameters("@memberTO").Value = memberTo
				.Parameters.Add(New MySqlParameter("@idCidade", MySqlDbType.Int32))
				.Parameters("@idCidade").Value = idCidade
			end with
			total = cmd.ExecuteScalar()
			
			if total = 0 then 
				with cmd
					.commandtext = "insert into tbl_logEnvioFollow (memberFrom, memberTo, messageID, idCidade) values (@memberFrom_Insert, @memberTO_Insert, @messageID, @idCidade_Insert)"
					.connection = conn
					.Parameters.Add(New MySqlParameter("@memberFrom_Insert", MySqlDbType.String))
					.Parameters("@memberFrom_Insert").Value = memberFrom
					.Parameters.Add(New MySqlParameter("@memberTO_Insert", MySqlDbType.String))
					.Parameters("@memberTO_Insert").Value = memberTo
					.Parameters.Add(New MySqlParameter("@messageID", MySqlDbType.String))
					.Parameters("@messageID").Value = messageID
					.Parameters.Add(New MySqlParameter("@idCidade_Insert", MySqlDbType.Int32))
					.Parameters("@idCidade_Insert").Value = idCidade
					.ExecuteNonQuery()
				end with
				return false
			end if
			
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try	
		
		return true
	
	End Function
					  		
End Class

Public Class _sendMailEmpresa

	'Dados privados
	Public idCidade as String
	Public idInternauta as String
	Public idEmpresa as String
	
	'Dados email
	Public nomeFrom as string
	Public emailFrom as string
	Public nomeReplyTo as String	
	Public replyTo as string
	Public nomeTo as string	
	Public emailTo as string
	Public assuntoMail as String
	Public corpoHTML as String
	Public messageID as String
	
	'Dados Adicionais
	Public DataEntrada as String
	Public DataSaida as String
	Public nAdultos as String
	Public nCriancas as String
	Public nQuartos as String
	Public assuntoMensagem as String
	Public Mensagem as String
		
	Public Sub sendMailHospedagem() 
		messageID = get_md5(DateTime.Now.toString())
		nomeFrom = dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobrenome")
		emailFrom = dados_user(idInternauta, "usuario")
		
		'Corpo do Email
		Dim urlPerfil as String = city.dados(idCidade, "urlcidade") & "/membros/" & dados_user(idInternauta, "id") & "," & GenerateSlug(tiraAscento( dados_user(idInternauta, "nome") ),255) & ".html"
		Dim logoEmail as String = city.dados(idCidade, "pasta_padrao") & "templates/" & city.dados(idCidade, "ntemplate") & "/" & city.dados(idCidade, "template") & "/images/logo-email.png"
		Dim corTemplate as String = city.dados(idCidade, "template")
		Dim bg as String = "#cede9d"
		Dim bgBox as String = "#f3f6ea"
		Dim bgBtn as String = "#003333"
		Dim borda as String = "#cccccc"			
		select case corTemplate
			case "verde"
			bg = "#cede9d"
			case Else
			bg = bg
		end select		
		
			Dim htmlEmail as String = String.Empty
			Dim cidadeMembro as String = String.Empty
			Dim idCidadeMembro as Integer = dados_user(idInternauta, "idCidade")
			If (Not IsNothing(idCidadeMembro) and idCidadeMembro <> 0) then
				cidadeMembro = "Cidade: " & city.dados(idCidade, "cidade") & "," & city.dados(idCidade, "uf")
			End If						
			
			htmlEmail = htmlEmail & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional //EN""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""><title>" & city.dados(idCidade, "nome_amigavel") & "</title></head><body style=""margin: 0; padding: 0;"" dir=""ltr"">"
			htmlEmail = htmlEmail & "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""40""><tr><td bgcolor=""#f7f7f7"" width=""100%"" style=""font-family: 'lucida grande', tahoma, verdana, arial, sans-serif;"">"
			htmlEmail = htmlEmail & "<table width=""620"" border=""0"" cellSpacing=""0"" cellPadding=""0"" style=""border:solid 1px " & bg & ";"">"
				htmlEmail = htmlEmail & "<tr><td style=""background: " & bg & ";padding: 4px 4px 4px 10px""><a href=""" & city.dados(idCidade, "urlCidade") & """><img src=""" & logoEmail & """ style=""border:0;display:block;""/></a></td></tr>"
				htmlEmail = htmlEmail & "<tr><td style=""padding: 15px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;font-size:12px;background:#ffffff;"">"
					htmlEmail = htmlEmail & "<table width=""100%""><tr>"
						'Coluna Esquerda
						htmlEmail = htmlEmail & "<td width=""440px"" style=""font-size: 12px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"" valign=""top"" align=""left"">"
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 15px 0;font-size: 13px;"">"
							htmlEmail = htmlEmail & "Ol&aacute; <strong>" & empresa.dados(idEmpresa, "Empresa") & "</strong>,"
							htmlEmail = htmlEmail & "</div>"
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 5px 0;"">"
							htmlEmail = htmlEmail & "<strong>" & dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobreNome") & "</strong> fez uma nova solicita&ccedil;&atilde;o de reserva pelo <strong>" & city.dados(idCidade, "nome_amigavel") & "</strong>, segue abaixo os dados do internauta bem como os de sua solicita&ccedil;&atilde;o."
							htmlEmail = htmlEmail & "</div>"
							
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 10px 0;"">"
								'Dados Internauta
								htmlEmail = htmlEmail & "<div style=""border-bottom: solid 1px #ccc; line-height:5px"">&nbsp;</div><br /><div><strong>Dados do Internauta:</strong></div>"
								htmlEmail = htmlEmail & "<table style=""margin:0;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"" cellPadding=""0"">"
									htmlEmail = htmlEmail & "<tr vAlign=""top"">"
										htmlEmail = htmlEmail & "<td style=""padding: 0 3px 0 0;""><a href=""" & urlPerfil & """><img src=""" & foto_user(idInternauta, 50, 50) & """ width=""50"" height=""50"" style=""border:solid 1px #cccccc;display:block;padding:5px;""/></a></td>"
										htmlEmail = htmlEmail & "<td style=""padding: 0 0 0 0;font-size: 11px;color:#999;""><a href=""" & urlPerfil & """ style=""color: #3b5998; font-size: 13px; font-weight: bold; text-decoration: none"" >" & dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobrenome") & "</a><br/>Membro Desde: " & databr(dados_user(idInternauta, "data_cadastro")) & "<br/>" & cidadeMembro & "</td>"
									htmlEmail = htmlEmail & "</tr>"
								htmlEmail = htmlEmail & "</table>"
								
								'Dados da Solicitacao
								htmlEmail = htmlEmail & "<div style=""border-bottom: solid 1px #ccc; line-height:2px"">&nbsp;</div><br /><div><strong>Dados da Solicita&ccedil;&atilde;o:</strong></div>"
								htmlEmail = htmlEmail & "<table style=""margin:0;"" cellPadding=""0"">"
									htmlEmail = htmlEmail & "<tr vAlign=""top"">"
										htmlEmail = htmlEmail & "<td style=""padding: 0 0 10px 0;font-size: 11px;color:#000000;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"">"
											htmlEmail = htmlEmail & "<div><strong>Data e Hora da Mensagem:</strong> " & dateTime.Now() & "</div>"
											htmlEmail = htmlEmail & "<div><strong>IP do Internauta:</strong> <a href=""http://www.infosniper.net/?ip_address=" & HttpContext.Current.request.ServerVariables("REMOTE_ADDR") & """>" & HttpContext.Current.request.ServerVariables("REMOTE_ADDR") & "</a></div>"											
											htmlEmail = htmlEmail & "<div style=""margin: 0 0 10px 0;""><strong>Email do Internauta:</strong> " & dados_user(idInternauta, "usuario") & "</div>"
											htmlEmail = htmlEmail & "<div><strong>Data Entrada:</strong> " & dataEntrada & "</div>"
											htmlEmail = htmlEmail & "<div><strong>Data Sa&iacute;da:</strong> " & dataSaida & "</div>"
											htmlEmail = htmlEmail & "<div><strong>N&ordm; de Noites:</strong> " & DateDiff( "d", cDate(dataEntrada), cDate(dataSaida) )  & "</div>"
											htmlEmail = htmlEmail & "<div><strong>N&ordm; de Quartos:</strong> " & nQuartos & "</div>"
											htmlEmail = htmlEmail & "<div><strong>N&ordm; de Adultos:</strong> " & nAdultos & "</div>"
											htmlEmail = htmlEmail & "<div style=""margin: 0 0 10px 0;""><strong>N&ordm; de Crian&ccedil;as:</strong> " & nCriancas & "</div>"
											htmlEmail = htmlEmail & "<div style=""margin: 0 0 5px 0;""><strong>Observa&ccedil;&otilde;es feita por " & dados_user(idInternauta, "nome") & "</strong>:</div>"
											htmlEmail = htmlEmail & "<div>" & Mensagem & "</div>"
										htmlEmail = htmlEmail & "</td>"
									htmlEmail = htmlEmail & "</tr>"
								htmlEmail = htmlEmail & "</table>"
								'Atenciosamente
								htmlEmail = htmlEmail & "<div style=""border-bottom: solid 1px #ccc; line-height:2px"">&nbsp;</div><br />"
								htmlEmail = htmlEmail & "<div style=""margin: 0px"">Atenciosamente,<br/><br/>Equipe do " & city.dados(idCidade, "nome_amigavel") & "<br/>" & city.dados(idCidade, "cidade") & "/" & city.dados(idCidade, "uf") & "</div>"								
							htmlEmail = htmlEmail & "</div>"						
						htmlEmail = htmlEmail & "</td>"
						
						'Coluna Direita
						htmlEmail = htmlEmail & "<td style=""padding: 0 10px 0 15px;"" vAlign=""top"" width=""150"" align=""left"">"
							htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
								htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:10px; background:" & bgBox & """>"
									htmlEmail = htmlEmail & "<div style=""margin:0 0 15px 0; font-size:13px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"">Para ver o Perfil de " & dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobreNome") & "</div>"
									htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
										htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:0px; background:" & bgBtn & """>"
											htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
												htmlEmail = htmlEmail & "<tr><td style=""padding:6px 15px;""><a href=""" & urlPerfil & """ style=""color:#ffffff;font-size:12px;text-decoration:none;text-transform:uppercase;font-weight:bold;"">Entrar</a></td></tr>"
											htmlEmail = htmlEmail & "</table>"
										htmlEmail = htmlEmail & "</td></tr>"
									htmlEmail = htmlEmail & "</table>"
								htmlEmail = htmlEmail & "</td></tr>"
							htmlEmail = htmlEmail & "</table><br/><br/>"
							
							
							htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
								htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:10px; background:" & bgBox & """>"
									htmlEmail = htmlEmail & "<div style=""margin:0 0 15px 0; font-size:13px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"">Para responder esse email</div>"
									htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
										htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:0px; background:" & bgBtn & """>"
											htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
												htmlEmail = htmlEmail & "<tr><td style=""padding:4px 10px;""><a href=""mailto:"&dados_user(idInternauta, "usuario")&"?Subject=Resposta Solicita&ccedil;&atilde;o de Reserva efetuada pelo " & city.dados(idCidade, "nome_amigavel") & "&body=" & dados_user(idInternauta, "nome") & ", para o per&iacute;odo informado de " & dataEntrada & " &agrave; " & dataSaida & " segue abaixo: "" style=""color:#ffffff;font-size:11px;text-decoration:none;text-transform:uppercase;font-weight:bold;"">Clique aqui</a></td></tr>"
											htmlEmail = htmlEmail & "</table>"
										htmlEmail = htmlEmail & "</td></tr>"
									htmlEmail = htmlEmail & "</table>"
								htmlEmail = htmlEmail & "</td></tr>"
							htmlEmail = htmlEmail & "</table>"							
							
						htmlEmail = htmlEmail & "</td>"
						
					htmlEmail = htmlEmail & "</tr></table>"

					'Lembre Box link url Perfil'
					htmlEmail = htmlEmail & "<br/>"
					htmlEmail = htmlEmail & "<table cellspacing=""0"" cellpadding=""0"" style=""border-collapse:collapse;width:100%;"">"
					htmlEmail = htmlEmail & "<tr><td style=""font-family:'lucida grande', tahoma, verdana, arial, sans-serif;padding:10px;background:" & bgBox & ";border:1px solid " & borda & """>"
						htmlEmail = htmlEmail & "<div style=""font-weight: bold; margin-bottom: 2px; font-size: 11px;"">Para visualizar o perfil de " & dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobreNome") & ", acesse:</div>"
						htmlEmail = htmlEmail & "<a href=""" & urlPerfil & """ style=""color:#3b5998;text-decoration:none;font-size:11px;"">" & urlPerfil & "</a>"
					htmlEmail = htmlEmail & "</td></tr>"
					htmlEmail = htmlEmail & "</table>"
											
				htmlEmail = htmlEmail & "</td></tr>"
			htmlEmail = htmlEmail & "</table>"
			
			'Rodape
			htmlEmail = htmlEmail & "<table width=""620"" border=""0"" cellSpacing=""0"" cellPadding=""0"" style=""padding:10px 0 0 0"">"
				htmlEmail = htmlEmail & "<tr><td style=""padding:10px;color:#999999;font-size:10px;font-family: 'lucida grande', tahoma, verdana, arial, sans-serif"">"
					htmlEmail = htmlEmail & "<div>Esta mensagem foi enviada para " & empresa.dados(idEmpresa, "email") & " pelo " & city.dados(idCidade, "nome_amigavel") & ". O " & city.dados(idCidade, "nome_amigavel") & " &eacute; um Portal s&eacute;rio, respeitado e que disponibiliza informa&ccedil;&otilde;es para um p&uacute;blico que realmente visita e se interessa por " & city.dados(idCidade, "cidade") &", para ver nossas diretrizes sobre seguran&ccedil;a de seus dados Acesse nossa <a href=""" & city.dados(idCidade, "urlCidade") & "/press-release/politica-de-privacidade.html"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">Pol&iacute;tica de Privacidade</a> e para mais informa&ccedil;&otilde;es sobre termos de uso do " & city.dados(idCidade, "nome_amigavel") & " acesse nosso <a href=""" & city.dados(idCidade, "urlCidade") & "/press-release/aviso-legal.html"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">Aviso Legal.</a></div><br/>"
					htmlEmail = htmlEmail & "<div>Por favor, n&atilde;o responda a esta mensagem, pois foi enviada de um endere&ccedil;o de email n&atilde;o monitorado. Esta mensagem &eacute; um e-mail dos servi&ccedil;os relacionados ao seu uso no " & city.dados(idCidade, "nome_amigavel") & ". Para informa&ccedil;&otilde;es gerais ou para pedir ajuda sobre sua conta no " & city.dados(idCidade, "nome_amigavel") & ", por favor acesse nosso <a href=""" & city.dados(idCidade, "urlCidade") & "/contato/"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">formul&aacute;rio de contato</a>.</div><br/>"
					htmlEmail = htmlEmail & "<div>" & city.dados(idCidade, "nome_amigavel") & "<br/>" & city.dados(idCidade, "endFranquia") & " - " & city.dados(idCidade, "bairroFranquia") & "<br/>" & city.dados(idCidade, "cepFranquia") & " - " & city.dados(idCidade, "cidadeFranquia") & " / " & city.dados(idCidade, "estadoFranquia") & "</div>"					
				htmlEmail = htmlEmail & "</td></tr>"
			htmlEmail = htmlEmail & "</table>"
			
			htmlEmail = htmlEmail & "</td></tr></table>"
			htmlEmail = htmlEmail & "<img src=""http://cms.guiadoturista.net/logSendMail/contatoEmpresa.asp?key=" & messageID & """ alt="""" style=""border: 0; height:1px; width:1px;"" />"
			htmlEmail = htmlEmail & "</body></html>"
			
			corpoHTML = htmlEmail
					
			nomeTo = empresa.dados(idEmpresa, "empresa")
			emailTo = empresa.dados(idEmpresa, "email")
			nomeReplyTo = city.dados(idCidade, "nome_amigavel")
			replyTo = "noreply@netcampos.com"
			
			Dim fromAddress As New Net.Mail.MailAddress(replyTo, nomeReplyTo)
			Dim toAddress As New Net.Mail.MailAddress(emailTo, nomeTo)
			
			Dim netMail As New System.Net.Mail.MailMessage()
			Try
				with netmail
					.From = fromAddress
					.To.Add(toAddress)					
					.Sender = New System.Net.Mail.MailAddress("" & nomeReplyTo & " <" & replyTo & ">")
					.ReplyTo = New System.Net.Mail.MailAddress("" & nomeFrom & " <" & emailFrom & ">")		
					.Priority = System.Net.Mail.MailPriority.High					
					.IsBodyHtml = True
					.Subject = assuntoMail
					.Body = corpoHTML 
					.Headers.Add("X-Mailer", "NetMail .NET (powered by Microsoft Windows Server)")
					.Headers.Add("X-MessageType", 2) 'contato empresa notification
					.Headers.Add("X-MessageUID", messageID)
					.Headers.Add("X-MessageFrom", idEmpresa)
					.Headers.Add("X-MessageTo", idInternauta)
					.Headers.Add("X-CidadeId", idCidade)					
					.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
					.SubjectEncoding = System.Text.Encoding.GetEncoding("utf-8")
					.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8")
				end with
			
				Dim netmailsmtp As New System.Net.Mail.SmtpClient
				with netmailsmtp
					.Host = "localhost"
					.send(netmail)
				end with
				
			Catch ex As Exception
				Exit Sub
			End Try
			
			Dim conn as SqlConnection 
			Dim cmd as SqlCommand
			conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
			
			try	
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand()
				with cmd
					.commandtext = "insert into tbl_contato_empresas_netportal (idCidade,idEmpresa,idInternauta,assunto,mensagem,ipInternauta) values (@idCidade,@idEmpresa,@idInternauta,@assunto,@mensagem,@ipInternauta)"
					.connection = conn
					.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
					.Parameters("@idCidade").Value = idCidade
					.Parameters.Add(New SqlParameter("@idEmpresa", SqlDbType.Int))
					.Parameters("@idEmpresa").Value = idEmpresa	
					.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))
					.Parameters("@idInternauta").Value = idInternauta	
					.Parameters.Add(New SqlParameter("@assunto", SqlDbType.VarChar))
					.Parameters("@assunto").Value = assuntoMail
					.Parameters.Add(New SqlParameter("@mensagem", SqlDbType.VarChar))
					.Parameters("@mensagem").Value = corpoHTML
					.Parameters.Add(New SqlParameter("@ipInternauta", SqlDbType.VarChar))
					.Parameters("@ipInternauta").Value = HttpContext.Current.request.ServerVariables("REMOTE_ADDR")					
					.ExecuteNonQuery()
				end with
				
			Catch ex As Exception		
				Exit Sub
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()
				If conn IsNot Nothing Then conn.Dispose()
			End Try	
			
				
			Exit Sub												
	
	End Sub
	
	Public Sub sendMailEmpresa()
	
		messageID = get_md5(DateTime.Now.toString())
		nomeFrom = dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobrenome")
		emailFrom = dados_user(idInternauta, "usuario")
		
		'Corpo do Email
		Dim urlPerfil as String = city.dados(idCidade, "urlcidade") & "/membros/" & dados_user(idInternauta, "id") & "," & GenerateSlug(tiraAscento( dados_user(idInternauta, "nome") ),255) & ".html"
		Dim logoEmail as String = city.dados(idCidade, "pasta_padrao") & "templates/" & city.dados(idCidade, "ntemplate") & "/" & city.dados(idCidade, "template") & "/images/logo-email.png"
		Dim corTemplate as String = city.dados(idCidade, "template")
		Dim bg as String = "#cede9d"
		Dim bgBox as String = "#f3f6ea"
		Dim bgBtn as String = "#003333"
		Dim borda as String = "#cccccc"			
		select case corTemplate
			case "verde"
			bg = "#cede9d"
			case Else
			bg = bg
		end select		
		
			Dim htmlEmail as String = String.Empty
			Dim cidadeMembro as String = String.Empty
			Dim idCidadeMembro as Integer = dados_user(idInternauta, "idCidade")
			If (Not IsNothing(idCidadeMembro) and idCidadeMembro <> 0) then
				cidadeMembro = "Cidade: " & city.dados(idCidade, "cidade") & "," & city.dados(idCidade, "uf")
			End If						
			
			htmlEmail = htmlEmail & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional //EN""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""><head><title>" & city.dados(idCidade, "nome_amigavel") & "</title></head><body style=""margin: 0; padding: 0;"" dir=""ltr"">"
			htmlEmail = htmlEmail & "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""40""><tr><td bgcolor=""#f7f7f7"" width=""100%"" style=""font-family: 'lucida grande', tahoma, verdana, arial, sans-serif;"">"
			htmlEmail = htmlEmail & "<table width=""620"" border=""0"" cellSpacing=""0"" cellPadding=""0"" style=""border:solid 1px " & bg & ";"">"
				htmlEmail = htmlEmail & "<tr><td style=""background: " & bg & ";padding: 4px 4px 4px 10px""><a href=""" & city.dados(idCidade, "urlCidade") & """><img src=""" & logoEmail & """ style=""border:0;display:block;""/></a></td></tr>"
				htmlEmail = htmlEmail & "<tr><td style=""padding: 15px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;font-size:12px;background:#ffffff;"">"
					htmlEmail = htmlEmail & "<table width=""100%""><tr>"
						'Coluna Esquerda
						htmlEmail = htmlEmail & "<td width=""440px"" style=""font-size: 12px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"" valign=""top"" align=""left"">"
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 15px 0;font-size: 13px;"">"
							htmlEmail = htmlEmail & "Ol&aacute; <strong>" & empresa.dados(idEmpresa, "Empresa") & "</strong>,"
							htmlEmail = htmlEmail & "</div>"
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 5px 0;"">"
							htmlEmail = htmlEmail & "<strong>" & dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobreNome") & "</strong> fez um novo contato pelo <strong>" & city.dados(idCidade, "nome_amigavel") & "</strong>, segue abaixo os dados do internauta bem como os de sua mensagem."
							htmlEmail = htmlEmail & "</div>"
							
							htmlEmail = htmlEmail & "<div style=""margin: 0 0 10px 0;"">"
								'Dados Internauta
								htmlEmail = htmlEmail & "<div style=""border-bottom: solid 1px #ccc; line-height:5px"">&nbsp;</div><br /><div><strong>Dados do Internauta:</strong></div>"
								htmlEmail = htmlEmail & "<table style=""margin:0;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"" cellPadding=""0"">"
									htmlEmail = htmlEmail & "<tr vAlign=""top"">"
										htmlEmail = htmlEmail & "<td style=""padding: 0 3px 0 0;""><a href=""" & urlPerfil & """><img src=""" & foto_user(idInternauta, 50, 50) & """ width=""50"" height=""50"" style=""border:solid 1px #cccccc;display:block;padding:5px;""/></a></td>"
										htmlEmail = htmlEmail & "<td style=""padding: 0 0 0 0;font-size: 11px;color:#999;""><a href=""" & urlPerfil & """ style=""color: #3b5998; font-size: 13px; font-weight: bold; text-decoration: none"" >" & dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobrenome") & "</a><br/>Membro Desde: " & databr(dados_user(idInternauta, "data_cadastro")) & "<br/>" & cidadeMembro & "</td>"
									htmlEmail = htmlEmail & "</tr>"
								htmlEmail = htmlEmail & "</table>"
								
								'Dados da Solicitacao
								htmlEmail = htmlEmail & "<div style=""border-bottom: solid 1px #ccc; line-height:2px"">&nbsp;</div><br /><div><strong>Dados da Mensagem:</strong></div>"
								htmlEmail = htmlEmail & "<table style=""margin:0;"" cellPadding=""0"">"
									htmlEmail = htmlEmail & "<tr vAlign=""top"">"
										htmlEmail = htmlEmail & "<td style=""padding: 0 0 10px 0;font-size: 11px;color:#000000;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"">"
											htmlEmail = htmlEmail & "<div><strong>Data e Hora da Mensagem:</strong> " & dateTime.Now() & "</div>"
											htmlEmail = htmlEmail & "<div><strong>IP do Internauta:</strong> <a href=""http://www.infosniper.net/?ip_address=" & HttpContext.Current.request.ServerVariables("REMOTE_ADDR") & """>" & HttpContext.Current.request.ServerVariables("REMOTE_ADDR") & "</a></div>"											
											htmlEmail = htmlEmail & "<div style=""margin: 0 0 10px 0;""><strong>Email do Internauta:</strong> " & dados_user(idInternauta, "usuario") & "</div>"
											htmlEmail = htmlEmail & "<div><strong>Assunto:</strong> " & assuntoMensagem & "</div>"
											htmlEmail = htmlEmail & "<div style=""margin: 0 0 5px 0;""><strong>Observa&ccedil;&otilde;es feita por " & dados_user(idInternauta, "nome") & "</strong>:</div>"
											htmlEmail = htmlEmail & "<div>" & Mensagem & "</div>"
										htmlEmail = htmlEmail & "</td>"
									htmlEmail = htmlEmail & "</tr>"
								htmlEmail = htmlEmail & "</table>"
								'Atenciosamente
								htmlEmail = htmlEmail & "<div style=""border-bottom: solid 1px #ccc; line-height:2px"">&nbsp;</div><br />"
								htmlEmail = htmlEmail & "<div style=""margin: 0px"">Atenciosamente,<br/><br/>Equipe do " & city.dados(idCidade, "nome_amigavel") & "<br/>" & city.dados(idCidade, "cidade") & "/" & city.dados(idCidade, "uf") & "</div>"								
							htmlEmail = htmlEmail & "</div>"						
						htmlEmail = htmlEmail & "</td>"
						
						'Coluna Direita
						htmlEmail = htmlEmail & "<td style=""padding: 0 10px 0 15px;"" vAlign=""top"" width=""150"" align=""left"">"
							htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
								htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:10px; background:" & bgBox & """>"
									htmlEmail = htmlEmail & "<div style=""margin:0 0 15px 0; font-size:13px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"">Para ver o Perfil de " & dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobreNome") & "</div>"
									htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
										htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:0px; background:" & bgBtn & """>"
											htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
												htmlEmail = htmlEmail & "<tr><td style=""padding:6px 15px;""><a href=""" & urlPerfil & """ style=""color:#ffffff;font-size:12px;text-decoration:none;text-transform:uppercase;font-weight:bold;"">Entrar</a></td></tr>"
											htmlEmail = htmlEmail & "</table>"
										htmlEmail = htmlEmail & "</td></tr>"
									htmlEmail = htmlEmail & "</table>"
								htmlEmail = htmlEmail & "</td></tr>"
							htmlEmail = htmlEmail & "</table><br/><br/>"
							
							
							htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
								htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:10px; background:" & bgBox & """>"
									htmlEmail = htmlEmail & "<div style=""margin:0 0 15px 0; font-size:13px;font-family:'lucida grande', tahoma, verdana, arial, sans-serif;"">Para responder esse email</div>"
									htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
										htmlEmail = htmlEmail & "<tr><td style=""border: solid 1px " & borda & "; padding:0px; background:" & bgBtn & """>"
											htmlEmail = htmlEmail & "<table style=""border-collapse: collapse"" cellSpacing=""0"" cellPadding=""0"">"
												htmlEmail = htmlEmail & "<tr><td style=""padding:4px 10px;""><a href=""mailto:"&dados_user(idInternauta, "usuario")&"?Subject=Resposta Contato efetuado pelo " & city.dados(idCidade, "nome_amigavel") & "&body=" & dados_user(idInternauta, "nome") & ", "" style=""color:#ffffff;font-size:11px;text-decoration:none;text-transform:uppercase;font-weight:bold;"">Clique aqui</a></td></tr>"
											htmlEmail = htmlEmail & "</table>"
										htmlEmail = htmlEmail & "</td></tr>"
									htmlEmail = htmlEmail & "</table>"
								htmlEmail = htmlEmail & "</td></tr>"
							htmlEmail = htmlEmail & "</table>"							
							
						htmlEmail = htmlEmail & "</td>"
						
					htmlEmail = htmlEmail & "</tr></table>"

					'Lembre Box link url Perfil'
					htmlEmail = htmlEmail & "<br/>"
					htmlEmail = htmlEmail & "<table cellspacing=""0"" cellpadding=""0"" style=""border-collapse:collapse;width:100%;"">"
					htmlEmail = htmlEmail & "<tr><td style=""font-family:'lucida grande', tahoma, verdana, arial, sans-serif;padding:10px;background:" & bgBox & ";border:1px solid " & borda & """>"
						htmlEmail = htmlEmail & "<div style=""font-weight: bold; margin-bottom: 2px; font-size: 11px;"">Para visualizar o perfil de " & dados_user(idInternauta, "nome") & " " & dados_user(idInternauta, "sobreNome") & ", acesse:</div>"
						htmlEmail = htmlEmail & "<a href=""" & urlPerfil & """ style=""color:#3b5998;text-decoration:none;font-size:11px;"">" & urlPerfil & "</a>"
					htmlEmail = htmlEmail & "</td></tr>"
					htmlEmail = htmlEmail & "</table>"
											
				htmlEmail = htmlEmail & "</td></tr>"
			htmlEmail = htmlEmail & "</table>"
			
			'Rodape
			htmlEmail = htmlEmail & "<table width=""620"" border=""0"" cellSpacing=""0"" cellPadding=""0"" style=""padding:10px 0 0 0"">"
				htmlEmail = htmlEmail & "<tr><td style=""padding:10px;color:#999999;font-size:10px;font-family: 'lucida grande', tahoma, verdana, arial, sans-serif"">"
					htmlEmail = htmlEmail & "<div>Esta mensagem foi enviada para " & empresa.dados(idEmpresa, "email") & " pelo " & city.dados(idCidade, "nome_amigavel") & ". O " & city.dados(idCidade, "nome_amigavel") & " &eacute; um Portal s&eacute;rio, respeitado e que disponibiliza informa&ccedil;&otilde;es para um p&uacute;blico que realmente visita e se interessa por " & city.dados(idCidade, "cidade") &", para ver nossas diretrizes sobre seguran&ccedil;a de seus dados Acesse nossa <a href=""" & city.dados(idCidade, "urlCidade") & "/press-release/politica-de-privacidade.html"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">Pol&iacute;tica de Privacidade</a> e para mais informa&ccedil;&otilde;es sobre termos de uso do " & city.dados(idCidade, "nome_amigavel") & " acesse nosso <a href=""" & city.dados(idCidade, "urlCidade") & "/press-release/aviso-legal.html"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">Aviso Legal.</a></div><br/>"
					htmlEmail = htmlEmail & "<div>Por favor, n&atilde;o responda a esta mensagem, pois foi enviada de um endere&ccedil;o de email n&atilde;o monitorado. Esta mensagem &eacute; um e-mail dos servi&ccedil;os relacionados ao seu uso no " & city.dados(idCidade, "nome_amigavel") & ". Para informa&ccedil;&otilde;es gerais ou para pedir ajuda sobre sua conta no " & city.dados(idCidade, "nome_amigavel") & ", por favor acesse nosso <a href=""" & city.dados(idCidade, "urlCidade") & "/contato/"" style=""color:#999999;font-weight:bold; text-decoration:underline;"">formul&aacute;rio de contato</a>.</div><br/>"
					htmlEmail = htmlEmail & "<div>" & city.dados(idCidade, "nome_amigavel") & "<br/>" & city.dados(idCidade, "endFranquia") & " - " & city.dados(idCidade, "bairroFranquia") & "<br/>" & city.dados(idCidade, "cepFranquia") & " - " & city.dados(idCidade, "cidadeFranquia") & " / " & city.dados(idCidade, "estadoFranquia") & "</div>"					
				htmlEmail = htmlEmail & "</td></tr>"
			htmlEmail = htmlEmail & "</table>"
			
			htmlEmail = htmlEmail & "</td></tr></table>"
			htmlEmail = htmlEmail & "<img src=""http://cms.guiadoturista.net/logSendMail/contatoEmpresa.asp?key=" & messageID & """ alt="""" style=""border: 0; height:1px; width:1px;"" />"
			htmlEmail = htmlEmail & "</body></html>"
			
			corpoHTML = htmlEmail	
			
			nomeTo = empresa.dados(idEmpresa, "empresa")
			emailTo = empresa.dados(idEmpresa, "email")
			nomeReplyTo = city.dados(idCidade, "nome_amigavel")
			replyTo = "noreply@netcampos.com"
			
			Dim fromAddress As New Net.Mail.MailAddress(replyTo, nomeReplyTo)
			Dim toAddress As New Net.Mail.MailAddress(emailTo, nomeTo)
			
			Dim netMail As New System.Net.Mail.MailMessage()
			Try
				with netmail
					.From = fromAddress
					.To.Add(toAddress)
					.Sender = New System.Net.Mail.MailAddress("" & nomeReplyTo & " <" & replyTo & ">")
					.ReplyTo = New System.Net.Mail.MailAddress("" & nomeFrom & " <" & emailFrom & ">")		
					.Priority = System.Net.Mail.MailPriority.High
					.IsBodyHtml = True
					.Subject = assuntoMail
					.Body = corpoHTML 
					.Headers.Add("X-Mailer", "NetMail .NET (powered by Microsoft Windows Server)")
					.Headers.Add("X-MessageType", 2) 'contato empresa notification
					.Headers.Add("X-MessageUID", messageID)
					.Headers.Add("X-MessageFrom", idEmpresa)
					.Headers.Add("X-MessageTo", idInternauta)
					.Headers.Add("X-CidadeId", idCidade)
					.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
					.SubjectEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
					.BodyEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
				end with
			
				Dim netmailsmtp As New System.Net.Mail.SmtpClient
				with netmailsmtp
					.Host = "localhost"
					.send(netmail)					
				end with

				
			Catch ex As Exception
				HttpContext.Current.response.write(ex.StackTrace())	
				Exit Sub
			End Try		
	
	
			Dim conn as SqlConnection 
			Dim cmd as SqlCommand
			conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
			
			try	
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand()
				with cmd
					.commandtext = "insert into tbl_contato_empresas_netportal (idCidade,idEmpresa,idInternauta,assunto,mensagem,ipInternauta) values (@idCidade,@idEmpresa,@idInternauta,@assunto,@mensagem,@ipInternauta)"
					.connection = conn
					.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
					.Parameters("@idCidade").Value = idCidade
					.Parameters.Add(New SqlParameter("@idEmpresa", SqlDbType.Int))
					.Parameters("@idEmpresa").Value = idEmpresa	
					.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))
					.Parameters("@idInternauta").Value = idInternauta	
					.Parameters.Add(New SqlParameter("@assunto", SqlDbType.VarChar))
					.Parameters("@assunto").Value = assuntoMail
					.Parameters.Add(New SqlParameter("@mensagem", SqlDbType.VarChar))
					.Parameters("@mensagem").Value = corpoHTML
					.Parameters.Add(New SqlParameter("@ipInternauta", SqlDbType.VarChar))
					.Parameters("@ipInternauta").Value = HttpContext.Current.request.ServerVariables("REMOTE_ADDR")					
					.ExecuteNonQuery()
				end with
				
			Catch ex As Exception		
				Exit Sub
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()
				If conn IsNot Nothing Then conn.Dispose()
			End Try		
				
		Exit Sub		
	End Sub

End Class