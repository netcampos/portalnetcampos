Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.IO
Imports System.Xml
Imports System.Configuration
Imports System.Text
Imports System.Security.Cryptography
Imports System.Web
Imports System.Web.Configuration
Imports System.Web.Management
Imports System.Web.Security
Imports System.Web.Caching
Imports Newtonsoft.Json
Imports System.Collections.Generic
Imports Facebook
Imports Facebook.Web
Imports TwitterVB2

Public Class Ajax
	
	Private idCidade as Integer = ConfigurationSettings.AppSettings("id_cidade")
	Private action as string = HttpContext.Current.request("action")
	Private str as new StringOperations()
	Private user as new npUsers()
	Private db as new database()
	Public _db as new _npDbViews()
	
	'Private Property
	Private Property sqlServer() As String
		Get
			Return ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
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
	Private sub setHttpHeader()
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.ClearContent()
		HttpContext.Current.Response.ClearHeaders()
		HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
		Dim ts As New TimeSpan(0,60,0)
		HttpContext.Current.Response.Cache.SetMaxAge(ts)
		HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddSeconds(3600))
		HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
		HttpContext.Current.Response.Cache.SetValidUntilExpires(true)
		HttpContext.Current.Response.Cache.VaryByHeaders("Accept-Language") = true
		HttpContext.Current.Response.Cache.VaryByHeaders("User-Agent") = true
		HttpContext.Current.Response.AppendHeader("X-Powered-By","Grupo NetCampos Tecnologia")
	End sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function draw()
		setHttpHeader()
		select case action
			case "mural-home"
			muralHome()
			case "userOnline"
			userOnline()
			case "userConnected"
			userConnected()
			case "modalProtect"
			modalProtect()
			case "validateUserName"
			validateUserName()
			case "cadNewUserModal"
			cadNewUserModal()
			case "passwordrecovery"
			passwordRecovery()
			case "notaConteudo"
			notaConteudo()
			case "publishComments"
			publishComments()
			case "checkFBPermission"
			checkFBPermission()
			case "checkTWPermission"
			checkTWPermission()
			case "loadMoreComments"
			loadMoreComments()
			case "deleteComments"
				deleteComments()
			case "loadAD"
				'loadAD()			
		end select
		
	End Function
	
	'***********************************************************************************************************
	'* Sub Mural de Recados Home
	'***********************************************************************************************************
	Private Sub MuralHome()
		dim html as string
		html = html & db.muralHome(10)
		_lib.write(html)
	End Sub
	
	'***********************************************************************************************************
	'* Sub Mural de Recados Home
	'***********************************************************************************************************
	Private Sub passwordRecovery()
		dim html as string
		html = html & "<html>"
		html = html & "<head>"
			html = html & "<meta name=""robots"" content=""noindex,nofollow"" />"
			html = html & "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />"
			html = html & "<title>Recuperar Senha</title>"
			html = html & str.loadCSS("/assets/css/gnc.popup.css",string.empty)
		html = html & "</head>"
		html = html & "<body id=""recuperar-senha"">"
			html = html & "<div id=""gnc-loading""></div>"
			html = html & "<form id=""gnc-form"" name=""gnc-form"">"
			html = html & "<div id=""gnc-wrapper"">"
				html = html & "<div id=""gnc-header"">"
					html = html & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
						html = html & "<tr>"
							html = html & "<td height=""25"" valign=""top""><span class=""label"">Digite seu email abaixo para recuperar sua senha.</span></td>"
						html = html & "</tr>"
						html = html & "<tr>"
							html = html & "<td>"
								html = html & "<div class=""container_campo inputemail"">"
									html = html & "<input name=""email"" type=""text"" id=""email"" value="""" size=""50"" maxlength=""100"" tabindex=""1"" />"
								html = html & "</div>"
							html = html & "</td>"
					html = html & "</tr>"
					html = html & "</table>"
				html = html & "</div>"
				html = html & "<div id=""gnc-body"">"
					html = html & "<table width=""100%"" border=""0"" cellspacing=""3"" cellpadding=""0"">"
					html = html & "<tr>"
						html = html & "<td>"
							html = html & "<div class=""container_captcha"">"
								html = html & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
									html = html & "<tr><td><div class=""img_captcha""><img id=""img_capcha"" alt=""captcha"" src=""/images/captcha.gif?t=" & _lib.getMD5(now()) & """ /></div></td></tr>"
									html = html & "<tr><td><div class=""reload_captcha""><a href=""#"" id=""reloadCaptcha""><img src=""/assets/images/reloadCapchaCyan.gif"" /><span>atualizar imagem</span></a></div></td></tr>"
								html = html & "</table>"
							html = html & "</div>"
						html = html & "</td>"
					html = html & "<td valign=""top"">"
						html = html & "<div class=""container_input_captcha"">"
							html = html & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
								html = html & "<tr><td><span class=""tt1"">Digite as palavras ao lado para recuperar sua senha</span></td></tr>"
								html = html & "<tr>"
									html = html & "<td>"
										html = html & "<div class=""campo caracteres"">"
											html = html & "<input name=""captcha"" type=""text"" id=""captcha"" size=""6"" maxlength=""6"" class="""" tabindex=""2"" />"
										html = html & "</div>"
									html = html & "</td>"
								html = html & "</tr>"
							html = html & "</table>"
						html = html & "</div>"
					html = html & "</td>"
					html = html & "<td>"
						html = html & "<div class=""enviar"">"
							html = html & "<input type=""submit"" name=""gncEnviar"" id=""gncEnviar"" value=""Enviar Senha"" tabindex=""3"" />"
						html = html & "</div>"
					html = html & "</td>"
					html = html & "</tr>"
					html = html & "</table>"
				html = html & "</div>"
			html = html & "</div>"
			html = html & "</form>"
			html = html & str.loadJS("//ajax.googleapis.com/ajax/libs/jquery/1.7.1/jquery.min.js")
			html = html & "<script>window.jQuery || document.write('<script src=""/assets/js/jquery.min.js""><\/script>')</script>"		
			html = html & str.loadJS("/assets/js/jquery.md5.js,jquery.gnc.user.js")
			html = html & "<script>$.recuperarSenha();</script>"
		html = html & "</body>"
		html = html & "</html>"
		_lib.write(html)
	End Sub		
	
	'***********************************************************************************************************
	'* Check User Connected
	'***********************************************************************************************************
	Private Sub userConnected()
		Dim conectado as String = "false"
		If HttpContext.Current.User.Identity.IsAuthenticated Then conectado = "true"
		Dim logonUser As New logonUser()
		logonUser.Connected = conectado

		Dim output As String = JsonConvert.SerializeObject(logonUser,Newtonsoft.Json.Formatting.Indented)
		_lib.write(output)
        HttpContext.Current.Response.ContentType = "application/json"
		_lib.responseEnd()
	End Sub
	
	'***********************************************************************************************************
	'* Check User Connected
	'***********************************************************************************************************
	Private Sub userOnline()
		Dim conectado as String = "false"
		If HttpContext.Current.User.Identity.IsAuthenticated Then conectado = "true"
		Dim logonUser As New logonUser()
		logonUser.Connected = conectado

		Dim output As String = JsonConvert.SerializeObject(logonUser,Newtonsoft.Json.Formatting.Indented)
		_lib.write(output)
        HttpContext.Current.Response.ContentType = "application/json"
		_lib.responseEnd()	
	End Sub	
	
	'***********************************************************************************************************
	'* Modal Protect
	'***********************************************************************************************************
	Private sub modalProtect()	
		dim html as string = user.modalWindowProtect()
		_lib.write(html)
        HttpContext.Current.Response.ContentType = "text/html"		
		_lib.responseEnd()
	End Sub
	
	'***********************************************************************************************************
	'* Validate UserName
	'***********************************************************************************************************
	Private sub validateUserName()
		dim html as string = "false"
		dim usuario as string = _lib.RF("u")
		dim passwd as string = _lib.RF("p")
		
		if _lib.Lenb(usuario) > 0 and _lib.Lenb(passwd) > 0 then 
			if db.checkLoginUserName(usuario,passwd).toString() then
				if db.getUserByLogin(usuario,"ativo") = 1 then
					user.userAuthenticationTicket(usuario)
					db.updateUltimoAcessoUsuario(usuario)
					html = "true"
				else
					html = "inative"					
				end if
			end if
		end if
	
		_lib.write(html)
        HttpContext.Current.Response.ContentType = "text/html"		
		_lib.responseEnd()	
	
	End Sub
	
	'***********************************************************************************************************
	'* Cad New User Modal
	'***********************************************************************************************************
	Private sub cadNewUserModal()
		Dim nome, dataNasc, sexo, email, senha, md5, passwdEmail,ipInternauta as string
		Dim existUser as boolean = true
		nome = _lib.RF("nome").toString()
		dataNasc = _lib.RF("data_nasc").toString()
		sexo = _lib.RF("sexo").toString()
		email = _lib.RF("email") .toString()
		senha = _lib.RF("senha").toString()
		md5 = _lib.getMD5(DateTime.Now.toString())
		passwdEmail = _lib.RF("pw-mail").toString()
		ipInternauta = HttpContext.Current.request.ServerVariables("REMOTE_ADDR")
		
		if _lib.LENB(email) > 0 then existUser = db.checkUserName(email)
		
		if existUser then 
			_lib.write("false")
			HttpContext.Current.Response.ContentType = "text/html"		
			_lib.responseEnd()			
		else
			_lib.write("true")
			HttpContext.Current.Response.ContentType = "text/html"		
			_lib.responseEnd()			
		end if
		
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************	
	Private Sub notaConteudo()
		dim html, retorno as string
		dim db as new database()
		dim id as integer = HttpContext.Current.request("id")
		dim modulo as integer = HttpContext.Current.request("modulo")
		dim nota as integer = HttpContext.Current.request("nota")
		
		retorno = db.notaConteudo(id,modulo,nota)
		if retorno = "true" then retorno = "Obrigado pela sua participação, seu voto foi computado com sucesso..."
		if retorno = "duplicado" then retorno = "Obrigado, mais você só pode votar uma única vez..."
		if retorno = "false" then retorno = "Descuple, mais ocorreu um erro inesperado, tente novamente mais tarde..."
		
		HttpContext.Current.response.Clear()
		HttpContext.Current.response.ContentType = "text/xml"
		HttpContext.Current.response.ContentEncoding = System.Text.Encoding.UTF8
				
		html = html & "<gnc-retorno>"
			html = html & "<resposta>" & retorno & "</resposta>"
		html = html & "</gnc-retorno>"
		
		HttpContext.Current.response.write(html)
	End Sub
	
	'***********************************************************************************************************
	'* Check User Connected
	'***********************************************************************************************************	
	Private sub publishComments()
		dim html as string
		dim publish as boolean
		If user.userON() Then
			dim userKey as string = user.getCurrentUserKey()
			dim actionUser as string = lcase(_lib.RF("actionUser"))
			dim fbActive as string = _lib.RF("fbActive")
			dim twActive as string = _lib.RF("twActive")
			dim review as integer = _lib.RF("hidden_star_rating")
			dim idConteudo as integer = _lib.RF("idConteudo")
			dim idModulo as integer = _lib.RF("idModulo")
			dim tituloPage as string = _lib.RF("tituloPage")
			dim uriPage as string = _lib.RF("uriPage")
			dim comments as string = _lib.RF("comments")
			comments = _lib.clearCommentsPost(comments)
			
			Dim userID as string = user.getCurrentUserID()
			Dim nomeUser as String = user.currentUserName()
			Dim fotoUser as String = user.currentUserPicture("square")
			Dim postkey as string = _lib.getMD5(DateTime.Now)
			Dim dataPublicacao As Date = DateTime.Now
			Dim acaoUsuario as String
			Dim userIdCity as String = db.getUser(userID, "idCidade")
			Dim userIdState as String = db.getUser(userID, "idEstado")
			Dim userCity as String = db.getUser(userID, "cidade")
			Dim userState as String = db.getUser(userID, "uf")
			dim localUsuario as String
			
			if _lib.lenb(userIdCity) > 0 and _lib.lenb(userIdState) > 0 then
				localUsuario = db.viewCidade(userIdCity,"Cidade") & ", " & ucase(db.viewCidade(userIdCity,"uf"))
			end if
			
			'grava post db
			if actionUser = "comentou" then acaoUsuario = "Comentou sobre: <a href=""" & uriPage & """>" & tituloPage & "</a>"
			
			publish = db.insertComments(userKey,postkey,idModulo,idConteudo,comments,acaoUsuario,review,"","","","","")
			
			if publish then
				
				'exibe post usuario
				html = html & "<li id=""post-" & postkey & """ class=""ui-user-post"">"
					
					html = html & "<div class=""comments-wrapper"">"
						html = html & "<div id=""comments-header"" class=""gnc-left"">"
							html = html & "<div class=""picture"">"
								html = html & fotoUser
							html = html & "</div>"
						html = html & "</div>"
						
						html = html & "<div id=""comments-content"" class=""gnc-right"">"
							html = html & "<div class=""comments-author"">"
								html = html & "<strong>" & nomeUser & "</strong>"
								html = html & "<span class=""user-local gnc-hidden"">" & localUsuario & "</span>"
								html = html & "<span class=""comments-momment gnc-hidden""><abbr class=""timestamp"" title=""" & dataPublishLogDate(dataPublicacao.toString()) & """>" & dataTweet(dataPublicacao.toString()) & "</abbr></span>"
								html = html & "<span class=""gnc-delete gnc-hidden""><span title=""Excluir Comentário"" data-key=""" & postKey & """ class=""comments-del gnc_tip""></span></span>"
								html = html & "<span class=""gnc-improprio gnc-hidden""><span title=""Marcar como Impróprio"" class=""comments-improprio gnc_tip""></span></span>"
							html = html & "</div>"
							html = html & "<div class=""comments"">"
								html = html & "<p>" & comments & "</p>"
							html = html & "</div>"
						html = html & "</div>"
					
						html = html & "<div style=""clear:both""></div>"
					html = html & "</div>"
					
					html = html & "<div class=""commets-footer"">"
						html = html & "<div class=""gnc-left"">"
							html = html & "<span class=""comment-rating""><span class=""rating review_" & review & """></span></span>"
						html = html & "</div>"
						
						html = html & "<div class=""gnc-right"">"
							html = html & "<div class=""barra-list-comments-share""><label>Compartilhar</label><button type=""button"" title=""Compartilhar no Facebook"" class=""comments-share-bt-facebook gnc_tip"">Compartilhar no Facebook</button><button type=""button"" title=""Compartilhar no Twitter"" class=""comments-share-bt-twitter gnc_tip"">Compartilhar no Twitter</button><button type=""button"" title=""Compartilhar no Google+"" class=""comments-share-bt-googleplus gnc_tip"">Compartilhar no Google+</button></div>"
						html = html & "</div>"
						
					html = html & "</div>"
				html = html & "</li>"
				
				_lib.write(html)
				
				if fbActive = "active" then 
					dim fbToken as string = db.getUserMeta(userKey,"fbToken")
					if _lib.lenb(fbToken) > 0 then
						Dim FBApp as FacebookClient = new FacebookClient(fbToken)
						try
							Dim info = DirectCast(FBApp.Get("me"), IDictionary(Of String, Object))
							Dim Nome As String = DirectCast(info("first_name"), String)
							
							Dim fbparams As New Dictionary(Of String, Object)
							fbparams.Add("message", comments)
							fbparams.Add("name", tituloPage)
							fbparams.Add("link", uriPage)
							Dim PostCommentFB As Facebook.JsonObject = FBApp.Post("/me/feed", fbparams)
							
							'_lib.write("fbOK")
						Catch ex As Exception
							if instr(ex.message(), "Error validating access token:") > 0 then
								'_lib.write("fbError")
							end if
							
						End Try					
					end if				
				end if
				
				if twActive = "active" then
					try
						Dim estatus as string = ( _lib.shorten(comments, 150, "...") & " " & uriPage)					
						Dim twToken as string = db.getUserMeta(userKey,"twToken")
						Dim twTokenSecret as string = db.getUserMeta(userKey,"twTokenSecret")						
						Dim ConsumerKey As String = db.getOAuth(2,"appID")
						Dim ConsumerKeySecret As String = db.getOAuth(2,"secretKey")				
						Dim tw As New TwitterVB2.TwitterAPI
						tw.AuthenticateWith(ConsumerKey,ConsumerKeySecret,twToken,twTokenSecret)
						tw.Update(estatus)
						Dim tweet As TwitterVB2.TwitterStatus
						tweet = tw.HomeTimeline.Item(0)
						'_lib.write("twOK")
					Catch ex As Exception
						'_lib.write("twError")
					End Try
				End if
				
			Else
				_lib.write("erro")
			end if	
		end if
	End Sub

	'***********************************************************************************************************
	'* Check User Connected
	'***********************************************************************************************************	
	private sub loadMoreComments()
		dim id : id = _lib.RF("id")
		dim idModulo : idModulo = _lib.RF("idModulo")
		dim inicio : inicio = _lib.RF("inicio")
		dim qtd : qtd = _lib.RF("qtd")
		dim html as String
		
		if _lib.lenb(id) > 0 and _lib.lenb(idModulo) > 0 and _lib.lenb(inicio) > 0 and _lib.lenb(qtd) > 0 then
			html = db.getCommentsContent(id,idModulo,inicio,qtd)
		end if
		
		_lib.write(html)	
	end sub
	
	'***********************************************************************************************************
	'* Check User Connected
	'***********************************************************************************************************	
	private sub deleteComments()
		dim id as string = _lib.RF("id")
		if _lib.lenb(id) > 0 then
			Dim conn as SqlConnection
			Dim cmd as SqlCommand
			conn = New SqlConnection(sqlServer())
			try	
				conn.Open()
				cmd = New SqlCommand("delete tbl_comentarios_netportal where md5 = @id", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@id", SqlDbType.VarChar))
					.Parameters("@id").Value = id
					.ExecuteNonQuery()
				end with
				
				cmd = New SqlCommand("delete tbl_rating_netportal where id_comments = @id", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@id", SqlDbType.VarChar))
					.Parameters("@id").Value = id
					.ExecuteNonQuery()
				end with				
				conn.close()
			Catch ex As Exception
				HttpContext.Current.Response.Clear()
				HttpContext.Current.Response.StatusCode = 500
				HttpContext.Current.Response.end()
			Finally
				conn.Close()
				conn.Dispose()
			End Try
		else
			HttpContext.Current.Response.Clear()
			HttpContext.Current.Response.StatusCode = 500
			HttpContext.Current.Response.end()		
		end if
	end sub	
	
	'***********************************************************************************************************
	'* Check User Connected
	'***********************************************************************************************************
	private sub checkFBPermission()
		If user.userON() Then
			dim userKey as string = user.getCurrentUserKey()
			dim fbToken as string = db.getUserMeta(userKey,"fbToken")
			if _lib.lenb(fbToken) > 0 then
				try
					Dim FBApp as FacebookClient = new FacebookClient(fbToken)				
					Dim info = DirectCast(FBApp.Get("me"), IDictionary(Of String, Object))
					Dim Nome As String = DirectCast(info("first_name"), String)
					_lib.write("fbOK")
				Catch ex As Exception
					if instr(ex.message(), "Error validating access token:") > 0 then
						_lib.write("fbError")
					end if
				End Try
			else
				_lib.write("fbNotFound")
			end if
		end if
	End Sub
	
	'***********************************************************************************************************
	'* Check User Connected
	'***********************************************************************************************************
	private sub checkTWPermission()
		If user.userON() Then
			dim userKey as string = user.getCurrentUserKey()
			dim twToken as string = db.getUserMeta(userKey,"twToken")
			dim twTokenSecret as string = db.getUserMeta(userKey,"twTokenSecret")
			if _lib.lenb(twToken) > 0 then
				try					
					Dim ConsumerKey As String = db.getOAuth(2,"appID")
					Dim ConsumerKeySecret As String = db.getOAuth(2,"secretKey")				
					Dim tw As New TwitterVB2.TwitterAPI
					tw.AuthenticateWith(ConsumerKey,ConsumerKeySecret,twToken,twTokenSecret)
					Dim twAccount As TwitterUser = tw.AccountInformation 
					_lib.write("twOK")
				Catch ex As Exception
					_lib.write("twError")
				End Try
			else
				_lib.write("twNotFound")
			end if
		end if
	End Sub	
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************
	Private Sub loadAD()
		dim area, local as integer
		dim html, ads as string
		
		area = HttpContext.Current.request("area")
		local = HttpContext.Current.request("local")
		
		if area > 0 and local > 0 then ads = _db.ads(area,local)
		
		if _lib.lenb(ads) > 0 then
			html = html & "<stream>"
				html = html & "<ads><![CDATA[" & ads & "]]></ads>"
			html = html & "</stream>"
		else
			HttpContext.Current.Response.StatusCode = 500
		end if		
		
		
		'HttpContext.Current.Response.Clear()
		'HttpContext.Current.Response.ContentType = "text/xml"
		'HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8		
		'_lib.write(html)
		'HttpContext.Current.Response.end()
		_lib.write("teste")
	end sub		
	
End Class

'***********************************************************************************************************
'* Check User Connected
'***********************************************************************************************************
public class logonUser
	Private m_Logon As String
	
	Public Property Connected() As String
		Get
			Return m_Logon
		End Get
		Set
			m_Logon = Value
		End Set
	End Property
	
End Class

