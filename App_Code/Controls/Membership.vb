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
Public Class _Membership

	Public idmembro as string

	Public function dados(coluna as string) as string
		dim retorno as string = rsDados(coluna)
		return retorno
	end function
	
	Public function rsDados(coluna as string) as string
			Dim conn as SqlConnection 
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)		
			
			try
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand("exec proc_tbl_login_clientes @id_membro", conn )
				with cmd
					.Parameters.Add(New SqlParameter("@id_membro", SqlDbType.varChar ))
					.Parameters("@id_membro").Value = idmembro
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
			
	end function	
	
End Class