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
Imports System.Net
Imports System.Net.Mail
Imports System.Drawing
Imports System.Drawing.Imaging

Public Class _dadosMembro

    Public debug As Boolean = True
    Public idmembro As String

    Public Function dados(ByVal coluna As String) As String
        Dim retorno As String = rsDados(coluna)
        Return retorno
    End Function

    Public Function rsDados(ByVal coluna As String) As String
        Dim conn As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)

        Try
            If conn.State <> ConnectionState.Open Then conn.Open()
            cmd = New SqlCommand("exec proc_tbl_login_clientes @id_membro", conn)
            With cmd
                .Parameters.Add(New SqlParameter("@id_membro", SqlDbType.varChar))
                .Parameters("@id_membro").Value = idmembro
            End With
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                dr.Read()
                Return dr(coluna).toString()
            End If
        Catch ex As exception
            error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Montagem do Sistema de Menus: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>")
        Finally
            If conn.State = ConnectionState.Open Then conn.Close()
            If conn IsNot Nothing Then conn.Dispose()
            If conn IsNot Nothing Then conn.Dispose()
            If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
        End Try

    End Function

End Class


Public Class _dadosPageMembros

	Public idConteudo as Integer = 0
	
	public function dados(coluna as string) as string
		Dim retorno as string = rsDados(coluna)
		return retorno
	end function
		
	public function rsDados(coluna as string) as string
		'@idConteudo = retorna conteudo filho da pagina mestra em idModulo
		
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)		
		
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select p.id, p.nome_pagina, p.url_amigavel, p.titulo from tbl_paginas_clientes_netportal as p where p.id = @idConteudo", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int ))
				.Parameters("@idConteudo").Value = idConteudo												
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr(coluna).toString()
			end if						
		catch ex as exception
			error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Montagem do Sistema de Menus: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try			
		
		return coluna
	end function

End Class

Public Class _cadInternauta

	Public _idCity as Integer
	Public _nomeDe as String
	Public _emailDe as String
	Public _urlCidade as String
	Public _assuntoMail as String
	
	'dados novo usuario
	Public _usuario as String
	Public _nome as String
	Public _sobrenome as String
	Public _sexo as String
	Public _data_nasc as String	
	Public _senha as String
	Public _cidade_mae as Integer
	Public _md5 as String
	Public _foto as String
	
	Public Function _novoInternautaOAuth() as boolean
		_senha = mid(_md5,1,6)
		_cidade_mae = _idCity
		_nomeDe = city.dados(_idCity, "nome_amigavel")
		_emailDe = city.dados(_idCity, "emailcidade")
		_urlCidade = city.dados(_idCity, "urlcidade")		
		
		if registro then
			if send_email then
				return true
			end if
		end if		
	End Function

	Private Function registro() as Boolean
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)		
		try
			conn = New SqlConnection(System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_sql"))
			conn.Open()
			cmd = New SqlCommand( "insert tbl_login_clientes_netportal (usuario,senha,md5,nome,sobrenome,sexo,cidade_mae) values (@usuario,@senha,@md5,@nome,@sobrenome,@sexo,@cidade_mae)", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@usuario", SqlDbType.VarChar))	
				.Parameters("@usuario").Value = _usuario
				.Parameters.Add(New SqlParameter("@senha", SqlDbType.VarChar))	
				.Parameters("@senha").Value = _senha
				.Parameters.Add(New SqlParameter("@md5", SqlDbType.VarChar))	
				.Parameters("@md5").Value = _md5
				.Parameters.Add(New SqlParameter("@nome", SqlDbType.VarChar))	
				.Parameters("@nome").Value = _nome
				.Parameters.Add(New SqlParameter("@sobrenome", SqlDbType.VarChar))	
				.Parameters("@sobrenome").Value = _sobrenome
				.Parameters.Add(New SqlParameter("@sexo", SqlDbType.VarChar))	
				.Parameters("@sexo").Value = lcase(_sexo)
				.Parameters.Add(New SqlParameter("@cidade_mae", SqlDbType.Int))	
				.Parameters("@cidade_mae").Value = _cidade_mae					
				.ExecuteNonQuery()
			end with
		Catch ex as exception
			return false
			exit Function
		Finally	
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()
		End Try
		return true
	end Function
	
	Private function send_email()
		dim html as string
		_assuntoMail = "Dados Registro - " & _nomeDe
		html = "<hr/>"
		html = html + "<span style=""font-family: arial""><p><b>" & _nome & ",</b></p>"
		html = html + "<p>Parabéns pelo seu cadastro no <strong>" & _nomeDe & ".</strong></p><p>O " & _nomeDe & " é parte integrante do Portal Guia do Turista.Net, o qual tem como principal objetivo divulgar as diversas opções de destinos turísticos, bem como ser uma rede de relacionamentos entre pessoas que desejam viajar e conhecer o Brasil.</p>"
		html = html + "<p>Com o seu cadastro, você poderá acessar qualquer um dos sites pertencentes a rede de portais Guia do Turista e utilizar nossos serviços para colaborar com o conteúdo das cidades disponíveis, publicando comentários, enviar fotos, escrever relatos, trocar experiências com outros internautas, criar listas de estabelecimentos favoritos, participar das inúmeras promoções exclusivas em cada cidade, e ainda será informado sobre todos os acontecimentos nas cidades que você desejar.</p>"
		html = html + "<p>Para confirmar seu registro, <a href=""" & _urlCidade & "/cadastro/confirma/" & _md5 & "/" & _usuario & """><b>clique aqui</b></a>.</p>"
		html = html + "<p>Abaixo estão suas informações para acessar as áreas restritas existentes no " & _nomeDe & ", guarde-as pois serão úteis.</p>"
		html = html + "<p><b>Login de Acesso:</b> " & _usuario & "<br>"
		html = html + "<b>Sua Senha:</b> " & _senha & "</p><br>"
		html = html + "<p>Em caso de dúvidas, envie um email para <a href=""mailto:" & _emailDe & """>" & _emailDe & "</a> que estaremos à sua inteira disposição.</p>"
		html = html + "<br><p>Nossas boas vindas!</p>"
		html = html + "<br><p><b>Atenciosamente,</b></p>"
		html = html + "<p>Equipe " & _nomeDe & "<br>"
		html = html + "" & _urlCidade &"<br>"
		html = html + "Tel:(12)3663-7321</p></span>"
		
		Dim netmail As New System.Net.Mail.MailMessage()
		Try	
			with netmail
				.From = New System.Net.Mail.MailAddress("" & _nomeDe & " <" & _emailDe & ">")
				.To.Add("<" & _usuario & ">")
				.Sender = New System.Net.Mail.MailAddress("<contato@guiadoturista.net>")		
				.Priority = System.Net.Mail.MailPriority.Normal
				.IsBodyHtml = True
				.Subject = _assuntoMail
				.Body = html 
				.Headers.Add("X-Mailer", "NetMail .NET (powered by Microsoft Windows 2003)")
				.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
				.SubjectEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
				.BodyEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
			end with
			
			Dim netmailsmtp As New System.Net.Mail.SmtpClient
			with netmailsmtp
				.Host = "localhost"
				.send(netmail)
			end with
			
		Catch exc As Exception
			return false
		End Try	
				
		return true
		netmail.Dispose()
	
	end function 	

End Class

Public Class _AlbumFotosInternauta

	Public idInternauta as String 'md5
	Public idCidade as Integer
	Public idModulo as Integer
	Public idConteudo as Integer
	
	Public Function _getIDAlbumFotos(ByVal Nome as String, ByVal local as String, ByVal descricao as String) as Integer
		
		If _checkAlbum then
			return _idAlbumFotos
			Exit Function
		Else
			If _createAlbum(nome, local, descricao) then
				return _idAlbumFotos
				Exit Function
			End If
		End If
		
	End Function
	
	Private Function _checkAlbum() as Boolean
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_Album_fotos_clientes where idInternauta = @idInternauta and idCidade = @idCidade and idModulo = @idModulo and idConteudo = @idConteudo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))
				.Parameters("@idInternauta").Value = idInternauta
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo				
			end with
			if cmd.ExecuteScalar > 0 then
				return True
			End If
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose() 
		End Try	
		
	End Function
	
	Private Function _createAlbum(ByVal Nome as String, ByVal local as String, ByVal descricao as String) as Boolean
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)		
		try
			conn = New SqlConnection(System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_sql"))
			conn.Open()
			cmd = New SqlCommand( "insert tbl_Album_fotos_clientes (idInternauta,idCidade,idModulo,idConteudo,Nome,local,descricao) values (@idInternauta,@idCidade,@idModulo,@idConteudo,@Nome,@local,@descricao)", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))	
				.Parameters("@idInternauta").Value = idInternauta
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))	
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))	
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))	
				.Parameters("@idConteudo").Value = idConteudo
				.Parameters.Add(New SqlParameter("@Nome", SqlDbType.VarChar))	
				.Parameters("@Nome").Value = Nome
				.Parameters.Add(New SqlParameter("@local", SqlDbType.VarChar))	
				.Parameters("@local").Value = local
				.Parameters.Add(New SqlParameter("@descricao", SqlDbType.VarChar))	
				.Parameters("@descricao").Value = descricao
				.ExecuteNonQuery()
			end with
			return true	
			exit Function
		Catch ex as exception
			return false
			exit Function
		Finally	
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()
		End Try
	
	End Function
	
	
	Private Function _idAlbumFotos() as Integer

		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select id from tbl_Album_fotos_clientes where idInternauta = @idInternauta and idCidade = @idCidade and idModulo = @idModulo and idConteudo = @idConteudo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))
				.Parameters("@idInternauta").Value = idInternauta
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo				
			end with
			if cmd.ExecuteScalar > 0 then
				return cmd.ExecuteScalar
			End If
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose() 
		End Try	

	End Function
	
	'########################## Inserir Foto na Galeria em Questao ######################################'
	Public Sub _insertFotoAlbum(ByVal idAlbum as Integer, ByVal idMembro as String, ByVal legenda as String, ByVal imagem as String)
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("insert into tbl_fotos_clientes_netportal(idAlbum,idInternauta,legenda,imagem) values(@idAlbum,@idMembro,@legenda,@imagem)", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idAlbum", SqlDbType.Int))
				.Parameters("@idAlbum").Value = idAlbum
				.Parameters.Add(New SqlParameter("@idMembro", SqlDbType.VarChar))
				.Parameters("@idMembro").Value = idMembro
				.Parameters.Add(New SqlParameter("@legenda", SqlDbType.VarChar))
				.Parameters("@legenda").Value = legenda
				.Parameters.Add(New SqlParameter("@imagem", SqlDbType.VarChar))
				.Parameters("@imagem").Value = imagem
			end with
			cmd.ExecuteNonQuery()
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose() 
		End Try		
	
	End Sub
	
	'########################## Get Informações sobre o Album de fotos do Usuário ######################################'
	Public Function _getInfoAlbumByID(ByVal idAlbum as Integer, ByVal coluna as String) as String
        Dim dr As SqlDataReader	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec np_Album_Fotos_Clientes @id", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = idAlbum
			end with
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                dr.Read()
                Return dr(coluna).toString()
            End If
			dr.Close()
		Catch ex As Exception
			Return ex.Message()
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
            If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try		
	
	End Function
	
	'########################## Retorna lista de fotos e um determinado Album de Fotos ######################################'
	Public function _getFotosByAlbumID(ByVal idAlbum as Integer) as System.Data.DataTable						
		Dim cmd As SqlCommand
		Dim conn As SqlConnection
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand( "exec np_Fotos_Album_Clientes @idAlbum", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idAlbum", SqlDbType.Int))
				.Parameters("@idAlbum").Value = idAlbum
			end with
			
			Dim da As sqlDataAdapter = New sqlDataAdapter(cmd)
			Dim dt As DataTable = New DataTable()
			da.Fill(dt)
			return dt
			
		Catch ex As Exception
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try
	
	End Function		
	
End Class