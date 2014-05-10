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
Imports System.Threading
Imports System.Globalization

Public Module _Membros

	Public function dados_user(idMembro as string, coluna as string) as string

		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select * from tbl_login_clientes_netportal where md5 = @idMembro", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idMembro", SqlDbType.VarChar))
				.Parameters("@idMembro").Value = idMembro			
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			httpContext.Current.response.Write(ex.Message.toString())
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try		
			
	end function
	
	'Check Social NetWork Connections
	Public Function check_Social_Network_User(ByVal idMembro as String, ByVal idServico as Integer) as boolean
		'ID dos Serviços
		'	1 - Facebook
		'	2 - Twitter
		'	3 - Google
		'	4 - Youtube
		'	5 - Windows Live
		'	6 - Flickr		
		
		Dim connect as boolean
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As Integer
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_login_clientes_netportal_OAuth where idMembro = @idMembro and idServico = @idServico", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idMembro", SqlDbType.VarChar))
				.Parameters("@idMembro").Value = idMembro
				.Parameters.Add(New SqlParameter("@idServico", SqlDbType.VarChar))
				.Parameters("@idServico").Value = idServico							
			end with
			dr = cmd.ExecuteScalar()
			if dr = 1 then
				connect = true
			end if
		Catch ex As Exception
			connect = false
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()
		End Try	
		
		Return connect			
		
	End Function
	
	'Delete Social NetWork Connections
	Public Function delete_Social_Network_User(ByVal idMembro as String, ByVal idServico as Integer) as boolean
		'ID dos Serviços
		'	1 - Facebook
		'	2 - Twitter
		'	3 - Google
		'	4 - Youtube
		'	5 - Windows Live
		'	6 - Flickr		
		
		Dim _delete as boolean
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("delete tbl_login_clientes_netportal_OAuth where idMembro=@idMembro and idServico = @idServico", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idMembro", SqlDbType.VarChar))
				.Parameters("@idMembro").Value = idMembro
				.Parameters.Add(New SqlParameter("@idServico", SqlDbType.VarChar))
				.Parameters("@idServico").Value = idServico	
				.ExecuteNonQuery()						
			end with
			_delete = true
		Catch ex As Exception
			_delete = false
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()
		End Try	
		
		Return _delete			
		
	End Function
	
	'Return Mail for Associate NetWork
	Public Function get_UserName_Social_Network(ByVal idServico as Integer, ByVal userName as String, Coluna as String) as String
		'ID dos Serviços
		'	1 - Facebook
		'	2 - Twitter
		'	3 - Google
		'	4 - Youtube
		'	5 - Windows Live
		'	6 - Flickr
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select * from tbl_login_clientes_netportal_OAuth where userName=@userName and idServico=@idServico", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@userName", SqlDbType.VarChar))
				.Parameters("@userName").Value = userName
				.Parameters.Add(New SqlParameter("@idServico", SqlDbType.VarChar))
				.Parameters("@idServico").Value = idServico
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr(coluna).toString()
				exit Function
			end if
			dr.close
		Catch ex As Exception
			return ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()
		End Try	
		
		Return ""			
		
	End Function		
	
	function foto_user(idMembro as string, largura as integer, altura as integer) as string
		Dim foto, sexo as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select foto, sexo from tbl_login_clientes_netportal where md5 = @idMembro", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idMembro", SqlDbType.VarChar))
				.Parameters("@idMembro").Value = idMembro			
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				foto = dr(0).toString()
				sexo = dr(1).toString()
			end if
			dr.close
		Catch ex As Exception
			httpContext.Current.response.Write(ex.Message.toString())
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try		
	
		if foto <> "" then
			
			if altura = 0 and largura <> 0 then
				foto = "http://images.guiadoturista.net/?img=" & foto & "&amp;l=" & largura.toString()
			elseif altura <> 0 and largura = 0 then
				foto = "http://images.guiadoturista.net/?img=" & foto & "&amp;a=" & altura.toString()
			elseif altura <> 0 and largura <> 0 then
				foto = "http://images.guiadoturista.net/?img=" & foto & "&amp;l=" & largura.toString() & "&amp;a=" & altura.toString()
			else
				foto = "http://images.guiadoturista.net/?img=" & foto
			end if
			
		end if
		
		if foto = "" then
			foto = "http://images.guiadoturista.net/?img=/cidades/cms/netgallery/media/images/ml_nophoto_male.png&amp;l=" & largura.toString() & "&amp;a=" & altura.toString()
			if sexo = "feminino" then
				foto = "http://images.guiadoturista.net/?img=/cidades/cms/netgallery/media/images/ml_nophoto_female.png&amp;l=" & largura.toString() & "&amp;a=" & altura.toString()
			end if
		end if			
	
		return foto
	end function
	
	function dataNascMember(ByVal idMember as String) as string
		Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
		Dim calBR as New CultureInfo("pt-BR")	
		Dim data as String = dados_user(idMember, "data_nasc")
		Dim html as String = String.Empty
		
		if (Not IsNothing(data) and data <> "") then
			Dim novaData As Date = Data
			html = html & "<div class=""birthday"">"
			html = html & "<dt>Data de nascimento:</dt>"
			html = html & "<dd>" & novaData.ToString("M", new CultureInfo("pt-BR")) & "</dd>"
			html = html & "</div>"
		end if
		return html
		
	end function
	
	Function cidadeMembro(ByVal idMember as String) as String
		Dim idCidade as Integer = dados_user(idMember, "idCidade")
		Dim html as String = String.Empty
		
		If (Not IsNothing(idCidade) and idCidade <> 0) then
			html = html & "<div class=""current_city"">"
			html = html & "<dt>Cidade:</dt>"
			html = html & "<dd>" & city.dados(idCidade, "cidade") & "," & city.dados(idCidade, "uf") & "</dd>"
			html = html & "</div>"
		End If
		return html
		
	End Function
	
	Function checkSeguindoMembro(ByVal idMembroLogado as String, ByVal idMembro as String) as Boolean
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim Seguindo As Integer
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
			
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_clientes_seguindo_netportal where idMembro = @idMembroLogado and idContato = @idMembro", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idMembroLogado", SqlDbType.VarChar))
				.Parameters("@idMembroLogado").Value = idMembroLogado				
				.Parameters.Add(New SqlParameter("@idMembro", SqlDbType.VarChar))
				.Parameters("@idMembro").Value = idMembro			
			end with
			Seguindo = cmd.executeScalar
		Catch ex As Exception
			httpContext.Current.response.Write(ex.Message.toString())
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose() 
		End Try	
		
		if seguindo = 1 then			
			return true
		end if
		
	End Function
	
    Public Function ValidateUserAjax(ByVal username As String, ByVal password As String) As Boolean

		Dim boolReturn As Boolean = False
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select md5 from tbl_login_clientes_netportal where usuario = @usuario and senha = @senha", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@usuario", SqlDbType.VarChar))
				.Parameters("@usuario").Value = username
				.Parameters.Add(New SqlParameter("@senha", SqlDbType.VarChar))
				.Parameters("@senha").Value = password
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				boolReturn = True
			end if
			dr.close
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try	
		
		Return boolReturn

    End Function
	
	Public Sub acessoRestrito()
	
		If not (HttpContext.Current.User.Identity.IsAuthenticated) Then
			FormsAuthentication.SignOut()
			Dim url as string = HttpContext.Current.Request.RawUrl
			Dim it as integer = url.LastIndexOf( "?" ) + 1
			url = url.Substring(it)
			HttpContext.Current.Response.redirect("~/login.aspx?ReturnUrl=/?" & HttpUtility.UrlEncode(url), false)
		end if	
	
	End Sub	

end Module
