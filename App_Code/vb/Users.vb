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

Public Class _npUsers2
	
	Private str as new _netPortalStrings()
	Private db as new _database()
		
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
	Public Function userON() as Boolean
		dim retorno as boolean = false
		retorno = HttpContext.Current.User.Identity.IsAuthenticated
		return retorno
	end Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function countUsers() as Integer
		dim retorno as integer = 0
		retorno = IIf(not IsNothing(httpContext.Current.session("totalUsers")), httpContext.Current.session("totalUsers"), db.countUsers() )
		return formatnumber(retorno)
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
	Public Function currentUserURI2() as String
		dim retorno as string
		if userON() then
			retorno = userURI2(getCurrentUserID())
		end if
		return retorno
	end Function	
	
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
	Public Function currentUserPicture(byval tipo as string) as String
		dim retorno as string
		dim width as integer
		dim height as integer
		if _lib2.lenb(tipo) = 0 then tipo = "square"
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
		end select
		if userON() then
			dim picture as string = db.getUser(getCurrentUserID(), "foto")
			
			dim UserKey as string = getCurrentUserKey()
			dim userImageProfile as string = db.getUserMeta(UserKey,"imageProfile")
			
			if _lib2.lenb(userImageProfile) > 0 then
				picture = userImageProfile			
				picture = "http://camposdojordao.netcampos.com/images/" & width & "/" & height & "/" & picture & "?mode=crop"
				picture = "<img src=""" & picture & """ class=""gnc-user-picture"">"
			elseif _lib2.lenb(picture) > 0 then
				if instr(picture, "/cidades/cms/netgallery/media/profiles/") > 0 then 
					picture = replace(picture,"/cidades/cms/netgallery/media/profiles/","")
					picture = "http://www.netcampos.com/user-picture/" & width & "/" & height & "/" & picture
				end if
				picture = "<img src=""" & picture & """ class=""gnc-user-picture"">"
			elseif _lib2.lenb(db.getUserMeta(getCurrentUserKey(),"fbid")) > 0 then
				picture = db.getUserMeta(getCurrentUserKey(),"fbid")
				picture = "http://graph.facebook.com/" & picture & "/picture?type=" & tipo
				picture = "<img src=""" & picture & """ class=""gnc-user-picture"">"
			else
				dim sexo as string = lcase(db.getUser(getCurrentUserID(), "sexo"))
				picture="gnc-masc"
				if _lib2.lenb(sexo) > 0 then
					if sexo = "feminino" or sexo="f" then picture="gnc-fem"
				end if
				picture = "<div class=""gnc-no-picture " & picture & """><img class=""gnc-user-no-picture""></div>"
			end if
			retorno = picture
		end if
		return retorno
	end Function
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getUserName(byval idUser as integer) as String
		dim nome, sobrenome as string
		Dim retorno as string
			if _lib2.lenb(idUser) > 0 then
				nome = db.getUser(idUser,"nome")
				sobrenome = db.getUser(idUser,"sobrenome")
				if instr(nome, sobrenome) > 0 or instr(sobrenome,"@") > 0 then sobrenome = ""
				retorno = nome & " " & sobrenome
				retorno = _lib2.trataNome(retorno)
			end if
		return retorno
   	End Function		
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
    Public Function userURI(ByVal userKey As String, Optional ByVal idUser as integer = 0) As String
		dim uri as string = "http://camposdojordao.netcampos.com/perfil/"
		dim retorno as string
			if _lib2.lenb(userKey) > 0 then
				retorno = uri & userKey & "/"
			elseif idUser > 0 then
				retorno = uri & db.getUser(idUser,"md5") & "/"
			end if
		return retorno
    End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
    Public Function userURI2(ByVal idUser as integer) As String
		dim uri as string = "http://www.netcampos.com/membros/"
		dim retorno as string
			if idUser > 0 then
				retorno = uri & idUser & "-" & GenerateSlug(tiraAscento(getUserName(idUser)),255) & ".html" 
			end if			
		return retorno
    End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************		
	Public Function GenerateSlug(phrase As String, maxLength As Integer) As String
		Dim str As String = phrase.ToLower()
		str = Regex.Replace(str, "[^a-z0-9\s-]", "")
		str = Regex.Replace(str, "[\s-]+", " ").Trim()
		str = str.Substring(0, If(str.Length <= maxLength, str.Length, maxLength)).Trim()
		str = Regex.Replace(str, "\s", "-")
		Return str
	End Function	
	
End Class