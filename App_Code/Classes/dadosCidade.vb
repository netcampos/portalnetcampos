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
Public Class city

	Public shared function dados(idcidade as integer, coluna as String) as String
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_cidades_netportal @id_cidade", conn)
			cmd.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
			cmd.Parameters("@id_cidade").Value = idcidade
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try	

   	End Function
	
	Public shared property logoCidade() as String
        Get
			return iif_template(webLogoCidade, dados(idcidade, "pasta_padrao") & "templates/" & dados(idcidade, "nTemplate") & "/" & dados(idcidade, "template") & "/images/logo-guia.png.ashx" )
        End Get	
        Set(ByVal value As String)
        End Set		
	End Property
	
	Public shared property logoCidadeSearch() as String
        Get
			return iif_template(webLogoCidade, dados(idcidade, "pasta_padrao") & "templates/" & dados(idcidade, "nTemplate") & "/" & dados(idcidade, "template") & "/images/logo-guia-search.png.ashx" )
        End Get	
        Set(ByVal value As String)
        End Set		
	End Property	
	
	Public shared property nTemplate() as String
        Get
			return iif_template(webNTemplate, dados(idcidade, "ntemplate"))
        End Get	
        Set(ByVal value As String)
        End Set		
	End Property
	
	Public shared property Template() as String
        Get
			return iif_template(webTemplate, dados(idcidade, "template"))
        End Get	
        Set(ByVal value As String)
        End Set		
	End Property
	
	Public Shared Property nomeCidade() as String
		Get
			return dados(idcidade, "cidade")
		End Get
		Set(ByVal value as String)
		
		End Set
	End Property	
	
	Public Shared Property urlCidade() as String
		Get
			return iif_template(webUrlCidade, lcase(dados(idcidade, "urlcidade")) )
		End Get
		Set(ByVal value as String)
		
		End Set
	End Property
	
	Public Shared Property nomeAmigavel() as String
		Get
			return iif_template(webNomeAmigavel, dados(idcidade, "nome_amigavel"))
		End Get
		Set(ByVal value as String)
		
		End Set
	End Property
	
	
	Public Shared Property urlCMS() as String
		Get
			return iif_template(webUrlCidadeOriginal, dados(idcidade, "urlCidade"))
		End Get
		Set(ByVal value as String)
		
		End Set
	End Property	
	
	
	Public Shared Property breadCrumb() as String
		Get
			return "<a class=""root"" href=""" & dados(idcidade, "urlCidade") & """ title=""Ir para: Página Inicial"" rel=""v:url"" property=""v:title"">Home</a>"
		End Get
		Set(ByVal value as String)
		
		End Set
	End Property
	
	Public Shared Property liberarVendaPub() as String
		Get
			return dados(idcidade, "liberar_venda_pub")
		End Get
		Set(ByVal value as String)
		
		End Set
	End Property
	
	Public Shared Property templateFull() as String
		Get
			return dados(idcidade, "templateFull")
		End Get
		Set(ByVal value as String)
		
		End Set
	End Property
	
	Public Shared Property idCity() as String
		Get
			return dados(idcidade, "id_Cidade")
		End Get
		Set(ByVal value as String)
		
		End Set
	End Property
	
	'######################## Inicio Retorna Total Membros ##################################################'
	Public shared Function getTotalMembers(ByVal idCidade as integer) as Integer
		Dim iTotal as Integer = 0
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec np_InternautasSearchTotal @idCidade", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
			end with
			iTotal = cmd.ExecuteScalar()
			httpContext.Current.session("totalMembers") = iTotal
		Catch ex As Exception
			
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try				
			
		return iTotal
	End Function
	
	'######################## Inicio Retorna Configuração do Módulo Da Cidade ##################################################'	

	Public shared function configModulo(ByVal idcidade as integer, ByVal idModulo as Integer, ByVal coluna as String) as String
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_config_pg_modulo @idCidade, @idModulo", conn)
			cmd.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
			cmd.Parameters("@idCidade").Value = idCidade
			cmd.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
			cmd.Parameters("@idModulo").Value = idModulo			
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try	

   	End Function
	'######################## Fim Retorna Configuração do Módulo Da Cidade ##################################################'
	
	'######################## Inicio Retorna Configuração do Bairro da Cidade ###############################################'
	
	Public function _dadosBairroByName(ByVal idcidade as integer, ByVal Bairro as String, ByVal Coluna as String) as String
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec np_bairros_dados_byName @idCidade, @bairro", conn)
			cmd.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
			cmd.Parameters("@idCidade").Value = idCidade
			cmd.Parameters.Add(New SqlParameter("@bairro", SqlDbType.VarChar))
			cmd.Parameters("@bairro").Value = Bairro			
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try	

	End Function
		
	'######################## Fim Retorna Configuração do Bairro da Cidade ###############################################'		
	
	
	'######################## Inicio Retorna Dados Configuração APP For City ###############################################'
	' Param: idCidade, idServico, Coluna
	' Return: Coluna as String
	'#######################################################################################################################'			
	Public shared Function _OAuth(byVal idcidade as integer, ByVal idServico as Integer, byVal coluna as String) as string
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
				return dr(coluna).toString()
				exit function
			end if
				
			dr.close()
		catch ex as exception
	
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try
		
		return ""
	End Function			
	'######################## Fim Retorna Configuração do Módulo Da Cidade ##################################################'
	
	'######################## Fim Class City Retorna Configuração do Módulo Da Cidade ##################################################'		
			
End Class

'################ Retorna Informações básicas referente a um determinado Conteudo Associado ao um Módulo #####################'
Public Class _getInfoPgConteudo

	Public idCidade as Integer
	Public idModulo as Integer
	Public idConteudo as Integer
		
	Public m_titulo as String
	Public m_Descricao as String
	Public m_Imagem as String
	Public m_url as String
	
	Dim tabela, campo_idCidade, campo_index, campo_title, campo_descricao, campo_imagem, campo_url_conteudo, sqlConteudo as String	
	
	Public Sub _setInfoConteudo()
			if _setInfoModulo() then
				_setInfo()
			End if
	End Sub
	
	Private Function _setInfoModulo() as Boolean
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select * from tbl_modulos_netportal where id = @pgModulo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@pgModulo", SqlDbType.Int))
				.Parameters("@pgModulo").Value = idModulo
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				tabela = dr("tabela").tostring()
				campo_idCidade = dr("campo_idcidade_conteudo").tostring()
				campo_index = dr("campo_index").tostring()
				campo_title = dr("campo_title").tostring()
				campo_descricao = dr("campo_descricao").tostring()
				campo_imagem = dr("campo_imagem").tostring()
				campo_url_conteudo = dr("campo_url_conteudo").tostring()
				return true
			end if
			dr.close()
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()			
		End Try		
	
	End Function
	
	Private Sub _setInfo()
		sqlConteudo = "select " & campo_title & "," & campo_descricao & "," & campo_imagem & "," & campo_url_conteudo & " from " & tabela & " where " & campo_index & "=@idConteudo and " & campo_idCidade & "=@idCidade"
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand(sqlConteudo, conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				tituloPage = dr(0).toString()
				descricaoPage = dr(1).toString()
				imagePage = dr(2).toString()
				urlPage = dr(3).toString() 
			end if
			dr.close()
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose() 
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()			 
		End Try
	
	End Sub
	
	Public property tituloPage() as String
        Get
			return m_titulo
        End Get	
        Set(ByVal value As String)
			m_titulo = value
        End Set		
	End Property
	
	Public property descricaoPage() as String
        Get
			return m_Descricao
        End Get	
        Set(ByVal value As String)
			m_Descricao = value
        End Set		
	End Property	
		
	Public property imagePage() as String
        Get
			return m_Imagem
        End Get	
        Set(ByVal value As String)
			m_Imagem = value
        End Set		
	End Property
	
	Public property urlPage() as String
        Get
			return m_url
        End Get	
        Set(ByVal value As String)
			m_url = value
        End Set		
	End Property

End Class


Public Class MetaHeader

	Public shared function getMetaHeader(idcidade as integer, idmodulo as integer, coluna as string) as String
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_config_pg_modulo @id_cidade, @id_modulo", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idcidade
				.Parameters.Add(New SqlParameter("@id_modulo", SqlDbType.Int))
				.Parameters("@id_modulo").Value = idmodulo	
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				return dr(coluna).toString()
			end if
		Catch ex As Exception
			return ex.Message()
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try				

	End Function

End Class

