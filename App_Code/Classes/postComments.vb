Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System.Web.Configuration
Imports System.Xml

Public Class _postFacebook

	Public idServico as Integer
	Public idCidade as string
	Public idUsuario as string
	Public mediaType as string
	Public mediasrc as string
	Public mediahref as string
	Public mediaTitulo as String
	Public mediaDescription as String
	Public message as string
	'Public api as API 
	'Public user as User
	'Public friends As IList(Of user)	
		
	Public appID, apiKey, secretKey, userCode as String 		
		
	Public Sub postFacebook()
		idServico = 1
		setKeyFacebook()
		_bindFacebook()
	End Sub
	
	Public Sub setKeyFacebook()
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)

		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select * from cidades_OAuth where idCidade = @idCidade and idServico = @idServico", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idServico", SqlDbType.Int))
				.Parameters("@idServico").Value = idServico	
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				appID = dr("appId").toString()
				apiKey = dr("apiKey").toString()
				secretKey = dr("secretKey").toString()
				userCode = getCodeMembro()
			end if
			dr.close()
		catch ex as exception

		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try		
	
	End Sub
	
	
	Private Function getCodeMembro() as String
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)

		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select * from tbl_login_clientes_netportal_OAuth where idCidade = @idCidade and idMembro = @idMembro and idServico = @idServico", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idMembro", SqlDbType.VarChar))
				.Parameters("@idMembro").Value = idUsuario
				.Parameters.Add(New SqlParameter("@idServico", SqlDbType.Int))
				.Parameters("@idServico").Value = idServico				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr("returnCode").toString()
			end if
			dr.close()
		catch ex as exception

		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try	
	
	End Function
	
	
	Private Sub _bindFacebook()

	
	End Sub		
					  		
End Class


'publish review business

Public Class _reView

	public idInternauta as string
	public idCidade as Integer
	public idModulo as Integer
	public idConteudo as Integer
	public comentario as string
	public acao as string
	public postid as String
	public review as integer
	public mes as integer	
	public ano as integer
	public sNota1 as String
	public sNota2 as String	
	public sNota3 as String
	public sNota4 as String	
	
	Sub execute()
		_excluirReview
		newReview()
		
	End Sub
	
	Public property oldPost() as String
		Get
			return _verifyReview()
		end Get
        Set(ByVal value As String)
        	
		End Set
	End property
	
	Private function _verifyReview() as String
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
	
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select md5 from tbl_comentarios_netportal where idModulo = @idModulo and idConteudo = @idConteudo and idInternauta = @IdInternauta", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
				.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))
				.Parameters("@idInternauta").Value = idInternauta
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr(0).toString()
			end if
			dr.close()
		catch ex as exception
			return ex.message()
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()			
		end try
	
	End Function
	
	Private sub _excluirReview()
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("delete tbl_comentarios_netportal where idModulo = @idModulo and idConteudo = @idConteudo and idInternauta = @IdInternauta", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
				.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))
				.Parameters("@idInternauta").Value = idInternauta				
			end with
			cmd.ExecuteNonQuery()
			
			
			cmd = New SqlCommand("delete tbl_rating_netportal where id_Modulo = @idModulo and id_Conteudo = @idConteudo and id_Internauta = @IdInternauta", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
				.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))
				.Parameters("@idInternauta").Value = idInternauta				
			end with
			cmd.ExecuteNonQuery()
			
		catch ex as exception
			HttpContext.Current.response.write(ex.message().toString() )
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		end try			
	
	End Sub
	
	Private Sub newReview()
		Dim num1, num2, num3, num4 As Integer
	    If Not Integer.TryParse(sNota1, num1) Then num1 = -1
	    If Not Integer.TryParse(sNota2, num2) Then num2 = -1
	    If Not Integer.TryParse(sNota3, num3) Then num3 = -1
	    If Not Integer.TryParse(sNota4, num4) Then num4 = -1
		
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		
		try
			conn.Open()
			cmd = New SqlCommand( "insert into tbl_comentarios_netportal (idinternauta, idcidade, idmodulo, idconteudo, comentario, acao, data, md5, review, tipo) values (@idInternauta, @idcidade, @idmodulo, @idconteudo, @comentario, @acao, getdate(), @md5, @review, 1)", conn )		
			with cmd
				.Parameters.Add(New SqlParameter("@idInternauta", SqlDbType.VarChar))
				.Parameters("@idInternauta").Value = idInternauta
				.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
				.Parameters("@idcidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
				.Parameters.Add(New SqlParameter("@comentario", SqlDbType.VarChar))
				.Parameters("@comentario").Value = comentario.toString()
				.Parameters.Add(New SqlParameter("@acao", SqlDbType.VarChar))
				.Parameters("@acao").Value = acao.toString()
				.Parameters.Add(New SqlParameter("@md5", SqlDbType.VarChar))
				.Parameters("@md5").Value = postid.toString()				
				.Parameters.Add(New SqlParameter("@review", SqlDbType.Int))
				.Parameters("@review").Value = review
				.ExecuteNonQuery()
			end with
			conn.close()
			
			if review > 0 then
				conn.Open()
				cmd = New SqlCommand( "insert into tbl_rating_netportal (id_internauta, id_comments, id_cidade, id_modulo, id_conteudo, nota, V_Nota1, V_Nota2, V_Nota3, V_Nota4, Mes, Ano, data ) values (@userid, @postid, @idcidade, @idmodulo, @idconteudo, @review, @vNota1, @vNota2, @vNota3, @vNota4, @mes, @ano, getdate() )", conn )
				with cmd
					.Parameters.Add(New SqlParameter("@userid", SqlDbType.VarChar))
					.Parameters("@userid").Value = idInternauta
					.Parameters.Add(New SqlParameter("@postid", SqlDbType.VarChar))
					.Parameters("@postid").Value = postid.toString()			
					.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
					.Parameters("@idcidade").Value = idCidade
					.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
					.Parameters("@idModulo").Value = idModulo
					.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
					.Parameters("@idConteudo").Value = idConteudo
					.Parameters.Add(New SqlParameter("@review", SqlDbType.Int))
					.Parameters("@review").Value = review
					.Parameters.Add(New SqlParameter("@vNota1", SqlDbType.Int))
					.Parameters("@vNota1").Value = num1
					.Parameters.Add(New SqlParameter("@vNota2", SqlDbType.Int))
					.Parameters("@vNota2").Value = num2
					.Parameters.Add(New SqlParameter("@vNota3", SqlDbType.Int))
					.Parameters("@vNota3").Value = num3
					.Parameters.Add(New SqlParameter("@vNota4", SqlDbType.Int))
					.Parameters("@vNota4").Value = num4	
					.Parameters.Add(New SqlParameter("@mes", SqlDbType.Int))
					.Parameters("@mes").Value = mes
					.Parameters.Add(New SqlParameter("@ano", SqlDbType.Int))
					.Parameters("@ano").Value = ano																									
					.ExecuteNonQuery()
				end with
				conn.close()
			end if			
					
		catch ex as exception
		
		
		End try
		
	End Sub
		
End Class


