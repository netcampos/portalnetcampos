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
Public Class _noticias

	Public Function metaHeaderCategoria(byVal idConteudo as Integer, byVal Coluna as string) as string
		'parametros procedure
		'@idConteudo = id referente a categoria em questão
		'@Coluna = coluna a qual se deseja obter os dados
		Dim idModulo as Integer = 5 'Módulo tbl_modulos_netportal
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_catNoticias @idConteudo", conn )
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
	End function	
		
End Class