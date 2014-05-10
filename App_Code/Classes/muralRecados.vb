Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Web.Configuration
public class _muralRecados
	
	Public function muralHome(qtd as integer) as String
		Dim html as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_comentarios_netportal @qtd, @idcidade, @idmodulo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd			
				.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
				.Parameters("@idcidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idmodulo", SqlDbType.Int))
				.Parameters("@idmodulo").Value = 13				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = "<ul id=""recados"">"
				while dr.Read()
					html = html & "<li><div>"
					html = html & "<strong>De: " & dr(0).toString() & "</strong><br />"
					html = html & "" & clearStringComments(cortaTexto(dr(2).toString(), 140)) & "<br />"					
					html = html & "</div></li>"
				end while
				html = html & "</ul>"
			end if
			dr.close
		Catch ex As Exception
			return ex.Message().toString()
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		
		return html
		
	end function 

end class