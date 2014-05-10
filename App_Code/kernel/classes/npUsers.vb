Imports Microsoft.VisualBasic
Imports System
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
Imports System.Web.Caching
Imports System.Globalization

Public Class npUsers
	
	Private str as new StringOperations()
	Private db as new database()
		
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
	Public Function userON() as Boolean
		dim retorno as boolean = false
		retorno = HttpContext.Current.User.Identity.IsAuthenticated
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function getCurrentUserID() as Integer
		dim retorno as integer
		if userON() then
			Dim user As MembershipUser
			Dim userKey as string
			try	
				user = Membership.GetUser()
				userKey = user.providerUserKey
				retorno = db.getInfoUserMD5(userKey, "id")						
			Catch ex As Exception
				retorno = 0
			End Try	
		end if
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function getCurrentUserMail() as String
		dim retorno as String
		if userON() then
			Dim user As MembershipUser
			try	
				user = Membership.GetUser()
				retorno = user.userName																
			Catch ex As Exception
				retorno = 0
			End Try	
		end if
		return retorno
	end Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function getCurrentUserKey() as String
		dim retorno as String
		if userON() then
			Dim user As MembershipUser
			Dim userKey as string
			try	
				user = Membership.GetUser()
				userKey = user.providerUserKey
				retorno = userKey					
			Catch ex As Exception
				retorno = 0
			End Try	
		end if
		return retorno
	end Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function currentUserPicture(byval tipo as string) as String
		dim retorno as string
		dim width as integer
		dim height as integer
		if _lib.lenb(tipo) = 0 then tipo = "square"
		select case tipo
			case "small"
				width = 50
				height = 0
			case "square"
				width = 50
				height = 50			
			case "large"
				width = 200
				height = 0			
		end select
		if userON() then
			dim picture as string
			picture = db.getUser(getCurrentUserID(), "foto")
			if _lib.lenb(picture) > 0 then
				if instr(picture, "/cidades/cms/netgallery/media/profiles/") > 0 then 
					picture = replace(picture,"/cidades/cms/netgallery/media/profiles/","")
					picture = "/user-picture/" & width & "/" & height & "/" & picture
				end if
				picture = "<img src=""" & picture & """ class=""gnc-user-picture"">"
			elseif _lib.lenb(db.getUserMeta(getCurrentUserKey(),"fbid")) > 0 then
				picture = db.getUserMeta(getCurrentUserKey(),"fbid")
				picture = "http://graph.facebook.com/" & picture & "/picture?type=" & tipo
				picture = "<img src=""" & picture & """ class=""gnc-user-picture"">"
			else
				dim sexo as string = lcase(db.getUser(getCurrentUserID(), "sexo"))
				picture="gnc-masc"
				if _lib.lenb(sexo) > 0 then
					if sexo = "feminino" or sexo="f" then picture="gnc-fem"
				end if
				picture = "<div class=""gnc-no-picture " & picture & """><img class=""gnc-user-no-picture""></div>"
			end if
			retorno = picture
		end if
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function UserPicture(ByVal id as Integer, byVal userKey as String, byval tipo as string) as String
		dim retorno as string
		dim width as integer
		dim height as integer
		if _lib.lenb(tipo) = 0 then tipo = "square"
		select case tipo
			case "small"
				width = 50
				height = 0
			case "square"
				width = 50
				height = 50			
			case "large"
				width = 200
				height = 0
		end select
		if _lib.lenb(id) > 0 and _lib.lenb(userKey) > 0 then
			dim picture as string
			picture = db.getUser(id, "foto")
			if _lib.lenb(picture) > 0 then
				if instr(picture, "/cidades/cms/netgallery/media/profiles/") > 0 then 
					picture = replace(picture,"/cidades/cms/netgallery/media/profiles/","")
					picture = "/user-picture/" & width & "/" & height & "/" & picture
				end if
				picture = "<img src=""" & picture & """ class=""gnc-user-picture"">"
			elseif _lib.lenb(db.getUserMeta(userKey,"fbid")) > 0 then
				picture = db.getUserMeta(userKey,"fbid")
				picture = "http://graph.facebook.com/" & picture & "/picture?type=" & tipo
				picture = "<img src=""" & picture & """ class=""gnc-user-picture"">"
			else
				dim sexo as string = lcase(db.getUser(id, "sexo"))
				picture="gnc-masc"
				if _lib.lenb(sexo) > 0 then
					if sexo = "feminino" or sexo="f" then picture="gnc-fem"
				end if
				picture = "<div class=""gnc-no-picture " & picture & """><img class=""gnc-user-no-picture""></div>"
			end if
			retorno = picture
		end if
		return retorno
	end Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function currentUserName() as String
		dim retorno as string
		if userON() then
			dim nome as string
			nome = db.getUser(getCurrentUserID(), "nome") & " " & db.getUser(getCurrentUserID(), "sobrenome")
			if _lib.lenb(nome) = 0 then nome = db.getUser(getCurrentUserID(), "usuario")
			retorno = nome
		end if
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function nomeUsuario(byval id as integer) as String
		dim retorno as string
		if _lib.lenb(id) > 0 then
			dim nome as string
			nome = db.getUser(id, "nome") & " " & db.getUser(id, "sobrenome")
			if _lib.lenb(nome) = 0 then nome = db.getUser(id, "usuario")
			retorno = nome
		end if
		return retorno
	end Function		
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Function top_toolbar() as string
		dim html as string
		html = html & "<div id=""barra-topo"">"& vbnewline
			html = html & "<div class=""gbar"">"& vbnewline
				if userOn() then
					
					Dim user As MembershipUser
					dim userMembro,idMembro,customerID as string
					try	
						user = Membership.GetUser()
						userMembro = user.userName
						idMembro = user.providerUserKey
						customerID = dados_user(idMembro, "id")						
					Catch ex As Exception
						userLogout("/cadastro/login/")
					End Try				
									
					html = html & "<div class=""gbar-l"">"& vbnewline
						html = html & "<ul>"& vbnewline
						html = html & "<li class=""icoHome""><a title=""Página Inicial"" href=""http://www.netcampos.com""><span>home</span></a></li>" & vbnewline
						html = html & "<li><a href=""http://www.netcampos.com/cadastro/inicio/"">" & userMembro & "</a></li>" & vbnewline
						html = html & "</ul>"& vbnewline
												
					html = html & "</div>"& vbnewline
					
					html = html & "<div class=""gbar-r"">"& vbnewline
						html = html & "<ul>"& vbnewline
						html = html & "<li><a href=""/me/desconectar/"">desconectar</a></li>"& vbnewline
						html = html & "<li><a class=""menu-nav-pipe"" href=""/cadastro/dados/"">meus dados</a></li>"& vbnewline
						html = html & "<li><a class=""menu-nav-pipe"" href=""http://www.netcampos.com/cadastro/favoritos/"">favoritos</a></li>"& vbnewline
						html = html & "<li><a class=""menu-nav-pipe"" href=""http://www.netcampos.com/membros/"">membros</a></li>"& vbnewline
						html = html & "<li><a class=""menu-nav-pipe"" href=""http://www.netcampos.com/cadastro/inicio/"">inicio</a></li>"& vbnewline
						html = html & "</ul>"	& vbnewline				
					html = html & "</div>"	& vbnewline				
				else			
					dim iTotalReg as integer = totalMembers() 
					html = html & "<div class=""gbar-l"">"& vbnewline
						html = html & "<ul>"& vbnewline
							html = html & "<li class=""icoHome""><a href=""/"" title=""Página Inicial"" rel=""nofollow""><span>home</span></a></li>"& vbnewline
							html = html & "<li id=""gbarLike""></li>"	& vbnewline						
							if iTotalReg > 0 then html = html & "<li class=""topTotalMembers""><strong>" & formatnumber(iTotalReg,0) & "</strong> pessoas já estão aqui! Participe você também.</li>"& vbnewline
						html = html & "</ul>"& vbnewline
					html = html & "</div>"& vbnewline
					
					html = html & "<div class=""gbar-r"">"& vbnewline
						html = html & "<ul>"& vbnewline
							html = html & "<li><a href=""/cadastro/"" title=""Cadastre-se"" style=""color:#5DB2D9; text-decoration:underline;"">cadastre-se</a></li>"& vbnewline
							html = html & "<li><a href=""/membros/"" title=""Membros"" class=""menu-nav-pipe"">membros</a></li>" & vbnewline           
							html = html & "<li><a id=""gnc-login"" href=""http://camposdojordao.netcampos.com"" title=""Entrar"" class=""menu-nav-pipe"" rel=""nofollow"">entrar</a></li>" & vbnewline            
						html = html & "</ul>"& vbnewline
					html = html & "</div>"& vbnewline
				end if
			html = html & "</div>"& vbnewline
		html = html & "</div>"& vbnewline
		return html
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Function gncUserTopBar() as string
		dim html as string
		html = html & "<div id=""barra-topo"">"& vbnewline
			html = html & "<div class=""gbar"">"& vbnewline
				if userOn() then
					
					Dim user As MembershipUser
					dim userMembro,idMembro,customerID as string
					try	
						user = Membership.GetUser()
						userMembro = user.userName
						idMembro = user.providerUserKey
						customerID = dados_user(idMembro, "id")						
					Catch ex As Exception
						userLogout("http://camposdojordao.netcampos.com")
					End Try
						
					html = html & "<div class=""gbar-l"">"& vbnewline
						html = html & "<ul>"& vbnewline
						html = html & "<li class=""icoHome""><a title=""Página Inicial"" href=""http://www.netcampos.com""><span>home</span></a></li>" & vbnewline
						html = html & "<li><a href=""http://www.netcampos.com/cadastro/inicio/"">" & userMembro & "</a></li>" & vbnewline
						html = html & "</ul>"& vbnewline
												
					html = html & "</div>"& vbnewline
					
					html = html & "<div class=""gbar-r"">"& vbnewline
						html = html & "<ul>"& vbnewline
						html = html & "<li><a href=""/me/desconectar/"">desconectar</a></li>"& vbnewline
						html = html & "<li><a class=""menu-nav-pipe"" href=""/cadastro/dados/"">meus dados</a></li>"& vbnewline
						html = html & "<li><a class=""menu-nav-pipe"" href=""http://www.netcampos.com/cadastro/favoritos/"">favoritos</a></li>"& vbnewline
						html = html & "<li><a class=""menu-nav-pipe"" href=""http://www.netcampos.com/membros/"">membros</a></li>"& vbnewline
						html = html & "<li><a class=""menu-nav-pipe"" href=""http://www.netcampos.com/cadastro/inicio/"">inicio</a></li>"& vbnewline
						html = html & "</ul>"	& vbnewline				
					html = html & "</div>"	& vbnewline				
				else			
					dim iTotalReg as integer = totalMembers() 
					html = html & "<div class=""gbar-l"">"& vbnewline
						html = html & "<ul>"& vbnewline
							html = html & "<li class=""icoHome""><a href=""/"" title=""Página Inicial"" rel=""nofollow""><span>home</span></a></li>"& vbnewline
							html = html & "<li id=""gbarLike""></li>"	& vbnewline						
							if iTotalReg > 0 then html = html & "<li class=""topTotalMembers""><strong>" & formatnumber(iTotalReg,0) & "</strong> pessoas já estão aqui! Participe você também.</li>"& vbnewline
						html = html & "</ul>"& vbnewline
					html = html & "</div>"& vbnewline
					
					html = html & "<div class=""gbar-r"">"& vbnewline
						html = html & "<ul>"& vbnewline
							html = html & "<li><a href=""/cadastro/"" title=""Cadastre-se"" style=""color:#5DB2D9; text-decoration:underline;"">cadastre-se</a></li>"& vbnewline
							html = html & "<li><a href=""/membros/"" title=""Membros"" class=""menu-nav-pipe"">membros</a></li>" & vbnewline           
							html = html & "<li><a id=""gnc-login"" href=""/cadastro/login/"" title=""Entrar"" class=""menu-nav-pipe"" rel=""nofollow"">entrar</a></li>" & vbnewline            
						html = html & "</ul>"& vbnewline
					html = html & "</div>"& vbnewline
				end if
			html = html & "</div>"& vbnewline
		html = html & "</div>"& vbnewline
		return html
	end Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function totalMembers() as Integer
		dim retorno as integer = 0
		retorno = IIf(not IsNothing(httpContext.Current.session("totalMembers")), httpContext.Current.session("totalMembers"), city.getTotalMembers(4787) )
		return retorno
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function usuariosOnline() as Integer
		dim retorno as integer = 0
		retorno = System.Web.Security.Membership.GetNumberOfUsersOnline
		return retorno
	end function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function modalWindowProtect() as String
		dim html as string
		html = html & "<div id=""gncModalProtect"" class=""modalWindowProtect"">"
			
			html = html & "<div class=""closeModalProtect simplemodal-close"" title=""fechar"">Fechar</div>"
			html = html & "<h1>Disponível Apenas para Membros Cadastrados</h1>"
			html = html & "<div class=""modalWindowProtectBody"">"
				html = html & "<div class=""cadeado""></div>"
				html = html & "<div class=""chamada""><h2 class=""chamadatitle"">Se você já é um membro cadastrado faça seu login abaixo, ou cadastre-se gratuitamente parar ficar por dentro de tudo que acontece em <strong>Campos do Jordão</strong> e ainda concorrer a diversos prêmios e promoções exclusivas.</h2></div>"
				html = html & "<div class=""containnerLogin"">"
					'coluna 01
					html = html & "<div id=""rightCollapsingContent"" class=""rightContent"">"
						html = html & "<h2 class=""bigTopSpace"">Ainda não é Membro?</h2>"
						html = html & "<p><strong>É Grátis!</strong> Você poderá participar do conteúdo do portal, escrever seus comentários, enviar fotos, receber promoções especiais e muito mais.</p>"
						html = html & "<a id=""actionNewAccount"">cadastre-se</a>"
					html = html & "</div>"
					'coluna 02
					html = html & "<div id=""leftColapsingContent"" class=""leftContent"">"
						html = html & "<div id=""ColapsingContainner"" class=""innerContentWrapper"">"
							html = html & "<div class=""innerContent"">"
								'social connect
								html = html & "<div id=""gncloginsocial"" class=""connectSocial"">"
									html = html & "<div id=""loginFacebook"">"
										html = html & "<div class=""inset"">"
											html = html & "<a href=""#"" rel=""facebook"" class=""login_button fb"">"
												html = html & "<span>Entrar usando o Facebook</span>"
											html = html & "</a>"
										html = html & "</div>"
										html = html & "<span class=""resume"">Economize Tempo! Utilize sua conta do Facebook para entrar.</span>"
									html = html & "</div>"
								html = html & "</div>"
								
								'gnc form login
								html = html & "<div class=""gnc-form"">"
									html = html & "<form id=""frmLogin"" name=""frmLogin"">"
										html = html & "<h2 id=""signInChoicesHeader"" class=""bigTopSpace"">Já possui cadastro? informe seu Login e Senha:</h2>"
										html = html & "<div class=""columnarForm"">"
											html = html & "<div class=""form-login"">"
												html = html & "<label class=""login"">Login / E-mail:<input type=""text"" id=""inputLoginModalProtect"" /><span class=""canto-input""></span></label>"
												html = html & "<label class=""senha"">Senha:<input type=""password"" id=""inputSenhaModalProtect"" /><span class=""canto-input""></span></label>"
												html = html & "<a id=""frmLoginActionPopup"">entrar</a>"
												
												html = html & "<div class=""error""></div>"
												
												html = html & "<div class=""login-help"">"
													html = html & "( <a title=""Esqueceu sua senha?"" id=""passwordRecovery"" href=""#"">Esqueceu sua senha?</a> )"
												html = html & "</div>"
												
												html = html & "<div class=""email-active"">"
													html = html & "( <a title=""Reenviar Email de Ativação"" id=""ativacaoRecovery"" href=""#"">Reenviar Email Ativação</a> )"
												html = html & "</div>"
												
												html = html & "<div class=""loginStatus""><img alt=""Back"" src=""/assets/images/loaders/ajax_load_facebook.gif""><span>por favor aguarde, processando...</span></div>"												
												
											html = html & "</div>"
										html = html & "</div>"
									html = html & "</form>"
								html = html & "</div>"
							html = html & "</div>"
						html = html & "</div>"
						
						html = html & "<div class=""arrowDivision""><img src=""/assets/images/members/login_divider.gif"" /></div>"
						html = html & "<span class=""divisionText"" style=""visibility: visible; opacity: 1;"">OU</span>"
						html = html & "<span class=""backText"" id=""modalBackButton""><img alt=""Back"" src=""/assets/images/members/modal_back_arrow.gif""></span>"
						
					html = html & "</div>"
					
					'gnc form new user	
					html = html & "<div id=""ContainnerfrmNewUser"">"
						html = html & "<form id=""frmNewUser"" name=""frmNewUser"" method=""post"">"
							
							html = html & "<h2>Formulário de Cadastro:</h2>"
							html = html & "<div id=""form_container"" class=""frm"">"
								
								html = html & "<div class=""frmNewUserLeft"">"
									html = html & "<div class=""row horizontal""><label for=""nome"">Nome:</label> <input type=""text"" id=""nome"" name=""nome"" class=""field_input {newUser:{required:true, minlength:4}}"" value="""" /> </div>"
									html = html & "<div class=""row horizontal""><label for=""sobrenome"">Sobrenome:</label> <input type=""text"" id=""sobrenome"" name=""sobrenome"" class=""field_input {newUser:{required:true, minlength:4}}"" value="""" /> </div>"
									html = html & "<div class=""row horizontal""><label for=""data_nasc"">Data Nascimento:</label> <span class=""fld""> <input type=""text"" id=""data_nasc"" name=""data_nasc"" class=""field_datanasc field_data {newUser:{required:true}}"" value=""""/></span></div>"
									html = html & "<div class=""row horizontal""><label for=""email"">Seu Email:</label> <input type=""text"" id=""email"" name=""email"" class=""field_input {newUser:{required:true, email:true}}"" value="""" /></div>"
									html = html & "<div class=""row horizontal""><label for=""confirmar_email"">Confirmar Email:</label> <input type=""text"" id=""confirmar_email"" name=""confirmar_email"" class=""field_input {newUser:{required:true, equalTo:'#email'}}"" value=""""/></div>"
									html = html & "<div class=""row horizontal""><label for=""senha"">Senha:</label> <input type=""password"" id=""senha"" name=""senha"" class=""field_input {newUser:{required:true, minlength:6}}"" value=""""/> <span class=""helpSenha"">(mínimo de 6 caracteres)</span></div>"
								html = html & "</div>"
								
								html = html & "<div class=""frmNewUserRigth"">"
									html = html & "<div class=""row labelGender""><label for=""form-widgets-sexo"">Sexo:</label><span class=""fld""><label class=""seletor""><input type=""radio"" id=""sexo_masc"" name=""sexo"" value=""Masculino"" class=""{newUser:{required:true}}"" />Masculino</label> <label class=""seletor""><input type=""radio"" id=""sexo_femin"" name=""sexo"" value=""Feminino"" />Feminino</label> <em htmlfor=""sexo"" class=""error"" style=""display: none;"">Por favor informe seu sexo.</em></span></div>"
									html = html & "<div class=""row labelAgreement""><label for=""form-widgets-agreement""><input type=""checkbox"" value=""true"" class=""checkbox {newUser:{required:true}}"" name=""agreement"" id=""agreement"">Eu estou de acordo e aceito o <a name=""sign_up.terms_of_use"" id=""sign_up/terms_of_use"" target=""_blank"" href=""/press-release/aviso-legal.html"">Termo de Uso</a> &amp; <a name=""sign_up.privacy_policy"" id=""sign_up/privacy_policy"" target=""_blank"" href=""/press-release/politica-de-privacidade.html"">Politica de Privacidade</a><span class=""requiredField""></span></label></div>"
									html = html & "<div class=""row horizontal""><button class=""actionButton"" type=""button"" id=""signUpButtonNewUser"" name=""signUpButton"">Efetuar Cadastro</button> <div class=""statusPost""><img alt=""Back"" src=""/assets/images/loaders/ajax_load_facebook.gif""></div></div>"
									html = html & "<div id=""ShowErroNewUser""><div class=""row labelfrmSubmit""><div id=""reg_error_inner"">Você deve preencher todos os campos.</div></div></div>"
								html = html & "</div>"
								
								html = html & "<input type=""hidden"" id=""pw-mail"" name=""pw-mail"" value="""" />"
								html = html & "<input type=""hidden"" name=""action"" value=""cadNewUserModal"" />"
								
							html = html & "</div>"
						html = html & "</form>"
						
						html = html & "<div id=""frmNewUserSuccess""></div>"
						
					html = html & "</div>"
					
				html = html & "</div>"
			html = html & "</div>"
			
		html = html & "</div>"
		
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Autentica username existente
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Sub userAuthenticationTicket(byVal userName as String)
		FormsAuthentication.Initialize()						
		Dim ticket As New FormsAuthenticationTicket(1, userName, DateTime.Now, DateTime.Now.AddMonths(3), true, "onlineUser", FormsAuthentication.FormsCookiePath)
		Dim encTicket As String = FormsAuthentication.Encrypt(ticket)
		Dim cookie as New HttpCookie(FormsAuthentication.FormsCookieName, encTicket)
		if (ticket.IsPersistent) then cookie.Expires = ticket.Expiration
		HttpContext.Current.Response.Cookies.Add(cookie)		
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Autentica username existente
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Sub userLogout(byVal uri as String)
		try	
			Dim user As MembershipUser
			user = Membership.GetUser()
			db.updateUltimaAtividadeUsuario(user.userName)
		Catch ex As Exception
		
		end Try
		idMembro = ""
		HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddMinutes(-1))
		HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache)		
		HttpContext.Current.Session.Abandon()
		FormsAuthentication.SignOut()
		HttpContext.Current.response.redirect(uri, false)		
	End Sub	
	
End Class