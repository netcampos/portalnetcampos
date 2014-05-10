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

Public Module _Comentarios

	public function CheckComentariosConteudo(idModulo as integer, idconteudo as integer) as string
		Dim total as integer = 0
		Dim html as string = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from view_tlb_comentarios_cidades where idcidade = @id_cidade and idmodulo = @id_modulo and idconteudo = @id_conteudo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@id_modulo", SqlDbType.Int))
				.Parameters("@id_modulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@id_conteudo", SqlDbType.Int))
				.Parameters("@id_conteudo").Value = idconteudo
			end with
			total = cmd.ExecuteScalar()
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
		End Try
		
		if idCidade > 0 then
			if total = 0 then
				html = "<div id=""postLists"" style=""display:none""><script type=""text/javascript"">var initComents=0; var mComments=" & idModulo & "; var cComments=" & idConteudo & ";</script><div id=""viewPostsComments""><span id=""totalComments""><div id=""headComments""><p class=""totalComments""><span><span class=""nComments"">0</span> Comentário(s) </span></p><div class=""btnNewComment""><button class=""btnAction"" title=""escrever comentário"">escrever comentário</button></div></div></span><div id=""posts-events""><ul class=""statuses""></ul></div></div></div>"
				html = html & "<div class=""headNoComments""><p>"
				html = html & "Não há comentários postados até o momento. "
				html = html & "<a href=""#"" class=""btnActionFirst"">Seja o primeiro!</a>" 
				html = html & "</p></div>"
			else
				html = "<div id=""postLists""><script type=""text/javascript"">var initComents=1; var mComments=" & idModulo & "; var cComments=" & idConteudo & ";</script><div id=""viewPostsComments""></div></div>"
			end if		
		 else
		 	html = "Nao foi possível carregar as informações, clique para tentar recarregar..."
		 end if
		return html

	end function
	
	
	Public Function CheckResenhasEmpresa(ByVal idCidade as Integer, ByVal idModulo as integer, ByVal idEmpresa as integer) as string
		Dim total as integer = 0
		Dim html as string = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from view_tlb_comentarios_cidades where idcidade = @id_cidade and idmodulo = @id_modulo and idconteudo = @id_conteudo and tipo = 1", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@id_modulo", SqlDbType.Int))
				.Parameters("@id_modulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@id_conteudo", SqlDbType.Int))
				.Parameters("@id_conteudo").Value = idEmpresa
			end with
			total = cmd.ExecuteScalar()
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
		End Try
		
		if idCidade > 0 then
			if total = 0 then
				html = "<div id=""postLists"" style=""display:none""><script type=""text/javascript"">var initComents=0; var mComments=" & idModulo & "; var cComments=" & idConteudo & ";</script><div id=""viewPostsComments""><span id=""totalComments""><div id=""headComments""><p class=""totalComments""><span><span class=""nComments"">0</span> Resenha(s) </span></p><div class=""btnNewComment""><button class=""btnAction"" title=""escrever comentário"">escrever resenha</button></div></div></span><div id=""posts-events""><ul class=""statuses""></ul></div></div></div>"
				html = html & "<div class=""headNoComments""><p>"
				html = html & "Não existem Resenhas até o momento. "
				html = html & "<a href=""#"" class=""btnActionFirst"">Seja o primeiro!</a>" 
				html = html & "</p></div>"
			else
				html = "<div id=""postLists""><script type=""text/javascript"">var initComents=1; var mComments=" & idModulo & "; var cComments=" & idConteudo & ";</script><div id=""viewPostsComments""></div></div>"
			end if		
		 else
		 	html = "Nao foi possível carregar as informações, clique para tentar recarregar..."
		 end if
		return html

	end function
	
	Public Function _getReviewConteudo(ByVal idCidade as Integer, ByVal idModulo as integer, ByVal idConteudo as integer, ByVal coluna as String) as String
	
		Dim valor as Double
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec np_getRating_Conteudo @idCidade, @idModulo, @idConteudo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				valor =	CDbl(dr(coluna)).toString()
				return valor
				Exit Function
			end if			
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()			  
		End Try
	
	End Function

end Module
