Imports Microsoft.VisualBasic
Imports System

Public Class _npUsers
	
	Private _lib as new _npLibrary()
	Private _db as new _npDbViews()
		
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
	Public Function totalMembers() as String
		dim retorno as string
		retorno = IIf(not IsNothing(httpContext.Current.session("totalMembers")), httpContext.Current.session("totalMembers"), _db.getTotalMembers() )
		return formatnumber(retorno,0)
	end function
	
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
				retorno = _db.getInfoUserMD5(userKey, "id")						
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
				retorno = "<div class=""alert alert-error""><button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & ex.message() & "</div>"
			End Try	
		end if
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public sub updateUserActivity()
		if userON() then
			Dim user As MembershipUser
			Dim userKey as string
			try	
				user = Membership.GetUser()
				userKey = user.providerUserKey
			Catch ex As Exception

			End Try			
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function currentUserName() as String
		dim retorno as string
		dim iduser as integer
		if userON() then
			retorno = getUserName(getCurrentUserID())
		end if
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function currentUserURI() as String
		dim retorno as string
		if userON() then
			retorno = userURI(getCurrentUserKey())
		end if
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
    Public Function userURI(ByVal userKey As String, Optional ByVal idUser as integer = 0) As String
		dim uri as string = "/perfil/"
		dim retorno as string
			if _lib.lenb(userKey) > 0 then
				retorno = uri & userKey & "/"
			elseif idUser > 0 then
				retorno = uri & _db.getUser(idUser,"md5") & "/"
			end if
		return retorno
    End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getUserName(byval idUser as integer) as String
		dim nome, sobrenome as string
		Dim retorno as string
			if _lib.lenb(idUser) > 0 then
				nome = _db.getUser(idUser,"nome")
				sobrenome = _db.getUser(idUser,"sobrenome")
				if instr(nome, sobrenome) > 0 or instr(sobrenome,"@") > 0 then sobrenome = ""
				retorno = nome & " " & sobrenome
				retorno = _lib.trataNome(retorno)
			end if
		return retorno
   	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function currentUserLevel() as integer
		dim retorno as integer = 1
		if userON() then
			retorno = _db.getUser(getCurrentUserID(), "tipo")
		end if
		return retorno
	end Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function currentUserThumb(byval tipo as string) as String
		dim retorno as string
		if userON() then
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
				case "normal"
					width = 120
					height = 120						
				case "large"
					width = 200
					height = 0
				case "hoverCard"
					width = 134
					height = 100							
			end select
				
			dim currentUserId as string = getCurrentUserID()
			dim picture as string = _db.getUser(currentUserId, "foto")
			dim UserKey as string = getCurrentUserKey()
			dim userFBID as string = _db.getUserMeta(UserKey,"fbid")
			dim userImageProfile as string = _db.getUserMeta(UserKey,"imageProfile")
			
			if _lib.lenb(userImageProfile) > 0 then
				picture = userImageProfile			
				picture = "http://camposdojordao.netcampos.com/images/" & width & "/" & height & "/" & picture & "?mode=crop"
			elseif _lib.lenb(picture) > 0 then
				if instr(picture, "/cidades/cms/netgallery/media/profiles/") > 0 then 
					picture = replace(picture,"/cidades/cms/netgallery/media/profiles/","")
					picture = "http://www.netcampos.com/user-picture/" & width & "/" & height & "/" & picture
				end if
			elseif _lib.lenb(userFBID) > 0 then
				picture = userFBID			
				if tipo = "hoverCard" then
					picture = "http://graph.facebook.com/" & picture & "/picture?type=large"
				else
					picture = "http://graph.facebook.com/" & picture & "/picture?type=" & tipo
				end if
			else
				dim sexo as string = lcase(_db.getUser(currentUserId, "sexo"))
				picture="nophoto-male.jpg"
				if _lib.lenb(sexo) > 0 then
					if sexo = "feminino" or sexo="f" then picture="nophoto-female.jpg"
				end if
				picture = "/assets/images/" & picture
				
			end if
			
			retorno = picture
				
		end if
		return retorno
	end Function
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function userThumb(byval userId as integer, byval tipo as string) as String	
		dim retorno, UserKey as string
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
			case "normal"
				width = 120
				height = 120						
			case "large"
				width = 200
				height = 0
			case "hoverCard"
				width = 134
				height = 100							
		end select
		
		if _lib.lenb(userId) > 0 and _lib.lenb(tipo) > 0 then
			dim picture, nome as string
			picture 	= _db.getUser(userId, "foto")
			UserKey  	= _db.getUser(userId, "md5")
			nome = _db.getUser(userId,"nome")			
			dim userImageProfile as string = _db.getUserMeta(UserKey,"imageProfile")
		
			if _lib.lenb(userImageProfile) > 0 then
				picture = userImageProfile			
				picture = "http://camposdojordao.netcampos.com/images/" & width & "/" & height & "/" & picture & "?mode=crop"
				if tipo = "hoverCard" then
					picture = "<img alt=""" & nome & """ width=""" & width & """ height=""" & height & """ src=""" & picture & """ class=""gnc-user-picture"">"
				else
					picture = "<img alt=""" & nome & """ src=""" & picture & """ class=""gnc-user-picture"">"
				end if			
			elseif _lib.lenb(picture) > 0 then
				if instr(picture, "/cidades/cms/netgallery/media/profiles/") > 0 then 
					picture = replace(picture,"/cidades/cms/netgallery/media/profiles/","")
					picture = "http://www.netcampos.com/user-picture/" & width & "/" & height & "/" & picture
				end if
				if tipo = "hoverCard" then
					picture = "<img alt=""" & nome & """ width=""" & width & """ height=""" & height & """ src=""" & picture & """ class=""gnc-user-picture"">"
				else
					picture = "<img alt=""" & nome & """ src=""" & picture & """ class=""gnc-user-picture"">"
				end if			
			elseif _lib.lenb(_db.getUserMeta(UserKey,"fbid")) > 0 then
				picture = _db.getUserMeta(UserKey,"fbid")
								
				if tipo = "hoverCard" then
					picture = "http://graph.facebook.com/" & picture & "/picture?type=large"
					picture = "<img alt=""" & nome & """ width=""" & width & """ height=""" & height & """ src=""" & picture & """ class=""gnc-user-picture"">"
				else
					picture = "http://graph.facebook.com/" & picture & "/picture?type=" & tipo
					picture = "<img alt=""" & nome & """ src=""" & picture & """ class=""gnc-user-picture"">"
				end if			
			else
				dim sexo as string = lcase(_db.getUser(userId, "sexo"))
				picture="nophoto-male.jpg"
				if _lib.lenb(sexo) > 0 then
					if sexo = "feminino" or sexo="f" then picture="nophoto-female.jpg"
				end if
				picture = "<img alt=""" & nome & """ src=""" & _db.domainAssets() & "/img/icons/" & tipo & "/" & picture & """ class=""gnc-user-picture"">"			
			end if
			retorno = picture
		end if
		
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
    Public Function urlHoverCard(ByVal userId as integer) As String
		dim retorno, UserKey as String
		if userId > 0 then
			UserKey  = _db.getUser(userId, "md5")
			retorno = "/ajax/?action=userHoverCard&id=" & userKey
		end if
		return retorno
    End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: Desconecta usuario
	'******************************************************************************************************************
	Public Sub logout(optional byval uri as string = "/" )
		try	
			Dim user As MembershipUser
			user = Membership.GetUser()
			_lib.write(user.userName)
			_db.userLogout(user.userName)
			
			HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddMinutes(-1))
			HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache)		
			HttpContext.Current.Session.Abandon()
			FormsAuthentication.SignOut()
			HttpContext.Current.response.redirect(uri, false)				
			
		Catch ex As Exception
		
		end Try
		
	End Sub		
		
End Class