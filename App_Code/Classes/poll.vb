Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Web.Configuration
public class _poll
	
	Public idEnquete As Integer
	
	Public function tituloEnquete() as String
		Dim titulo as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select top 1 id, pergunta from enquetes_netportal where ativo = 1 and idcidade = @id_cidade order by newid()", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				titulo = dr(1).toString()
				idEnquete = dr(0)			
			end if
			dr.close
		Catch ex As Exception
			return titulo
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		
		return titulo
	end function 
	
	
	public function perguntasEnquete() as string
		Dim html as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select id, resp from perguntas_enquetes where idenquete = @idEnquete order by resp asc;", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idEnquete", SqlDbType.Int))
				.Parameters("@idEnquete").Value = idEnquete
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = "<form name=""frmEnquete"" id=""frmEnquete"" method=""post""> <ul>"
				while dr.Read()
					html = html & " <li> <p class=""respEnquete"">"
					html = html & " <input name=""enqueteResposta"" class=""resposta"" type=""radio"" value=""" & dr(0).toString & """ />"
					html = html & " <span>" & dr(1).toString & "</span></p>"
					html = html & " </li>"
				end while
				html = html & "</ul>"
				html = html & "<div class=""btns"">"
				html = html & "<a href=""#"" id=""btVotarEnquete"" class=""btVotar"" title=""Votar""> <span> votar </span> </a>"
				html = html & "<a href=""#"" id=""btResultadoEnquete"" class=""btResultado"" title=""Resultado""> <span> resultado </span> </a>"
				html = html & "</div> <input type=""hidden"" name=""idEnquete"" id=""idEnquete"" value=""" & idEnquete & """/> <input type=""hidden"" name=""tokenEnquete"" id=""tokenEnquete"" value=""" & idCidade & """/> </form>"
			end if
			dr.close
		Catch ex As Exception
			return "erro"
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try		
		
		return html	
	
	end function


end class