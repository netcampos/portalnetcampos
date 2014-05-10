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

Public Class _dadosInstitucionais

		Public masterPage as boolean = false 'defini se pagina mestra ou pagina conteudo
		Public idCidade as integer
		Public idModulo as Integer
		Public idMaster as Integer = 0
		Public idConteudo as Integer = 0
		
		public function dados(coluna as string) as string
			if masterPage then idMaster = 1
			Dim retorno as string = rsDados(coluna)
			return retorno
		end function
		
	public function rsDados(coluna as string) as string
		'parametros procedure
		'@idCidade = cidade do conteudo
		'@idModulo = modulo do conteudo utilizado com masterPage é true
		'@Master = define se o conteudo é uma Pagina mestra - utilizado com idMaster = 1
		'@idConteudo = retorna conteudo filho da pagina mestra em idModulo
		
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)		
		
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_dados_page_institucional @idCidade, @idModulo, @Master, @idConteudo", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int ))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int ))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@master", SqlDbType.Int ))
				.Parameters("@master").Value = idMaster
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

end class

'retorna conteudo relacionado páginas institucionais
public class _conteudoRelacionadoInstitucionais

	Public thumbWidth as integer = 220
	Public thumbHeight as integer = 160
	Public exibeBanner as boolean = false
	Public cssPubAds as boolean = false
	Public ads as string
	Public idCidade as integer
	Public idModulo as Integer
	Public idCategoria as Integer
	Public exibir as boolean = false
	Public htmlConteudo as string
	Public padding as integer
		
	public function dados() as string
		return rsDados()
	end function
	
	public function rsDados() as string
		Dim x as integer
		Dim totalrows as integer = qtdDados()
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_paginas_relacionadas_netportal @id_cidade, @id_relacao, @id_modulo", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@id_relacao", SqlDbType.Int))
				.Parameters("@id_relacao").Value = idCategoria
				.Parameters.Add(New SqlParameter("@id_modulo", SqlDbType.Int))
				.Parameters("@id_modulo").Value = idModulo												
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				exibir = true
				if cssPubAds then htmlConteudo = "<div class=""pubad-relacionado"">"			
				htmlConteudo = htmlConteudo & "<div id=""conteudo-relacionado"">"
				htmlConteudo = htmlConteudo & "<div class=""title-section-pub""><h2 class=""title-c-rel"">Conteudo Relacionado</h2></div><div class=""crelacionado""><ul>"
				while dr.Read()
					
					if (totalrows > padding and padding > 0 ) then
						htmlConteudo = htmlConteudo & noPadding(x, padding)
					else
						htmlConteudo = htmlConteudo & "<li>"
					end if
					
					htmlConteudo = htmlConteudo & "<div class=""conteudo-lista""><p class=""ttp""><a href=""" & iif_url(webUrlCidade, dr(5).toString()) & """>" & dr(6).toString() & "</a></p>"
					htmlConteudo = htmlConteudo & fotoLegenda(dr(3).toString(), dr(0).toString(), iif_url(webUrlCidade, dr(5).toString()), thumbWidth, thumbHeight )
					htmlConteudo = htmlConteudo & "<p class=""rp"">" & dr(2).toString()  & "</p></div>"
					htmlConteudo = htmlConteudo & "</li>"
					
					if (totalrows > padding and padding > 0 ) then htmlConteudo = htmlConteudo & listaClear(x, padding)
					
					x = x + 1
					
				end while
				htmlConteudo = htmlConteudo & "</ul></div>"
				if exibeBanner then htmlConteudo = htmlConteudo  & "<div class=""pub""><div>" & ads  & "</div></div>"
				htmlConteudo = htmlConteudo & "</div>"
				if cssPubAds then htmlConteudo = htmlConteudo & "</div>"
			else
				return rsNoticias()
			end if
		catch ex as exception
			error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Montagem do Sistema de Menus: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try		
		
		return htmlConteudo
	
	end function
	
	private function qtdDados() as integer
		Dim x as integer = 0
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		Dim dr As SqlDataReader	
		cmd = New SqlCommand("exec proc_paginas_relacionadas_netportal @id_cidade, @id_relacao, @id_modulo", conn )
		with cmd
			.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
			.Parameters("@id_cidade").Value = idCidade
			.Parameters.Add(New SqlParameter("@id_relacao", SqlDbType.Int))
			.Parameters("@id_relacao").Value = idCategoria
			.Parameters.Add(New SqlParameter("@id_modulo", SqlDbType.Int))
			.Parameters("@id_modulo").Value = idModulo												
		end with
		
		try		
			If conn.State <> ConnectionState.Open Then conn.Open()		
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				while dr.Read()
					x = x + 1
				End While
			end if
		catch ex as exception
			error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Montagem do Sistema de Menus: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try				
		 		
		return x
		
	end function
	
	
	public function rsNoticias() as String
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
		
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_destaques_noticias_home_portal @idCidade, 1, 3, 3", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int ))
				.Parameters("@idCidade").Value = idCidade											
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				exibir = true
				if cssPubAds then htmlConteudo = "<div class=""pubad-relacionado"">"			
				htmlConteudo = htmlConteudo & "<div id=""conteudo-relacionado"">"
				htmlConteudo = htmlConteudo & "<div class=""title-section-pub""><h2 class=""title-c-rel"">Conteudo Relacionado</h1></div><div class=""crelacionado""><ul>"				
				while dr.Read()
					htmlConteudo = htmlConteudo & "<li><div class=""conteudo-lista""><p class=""ttp""><a href=""" & iif_url(webUrlCidade, dr(5).toString() ) & """>" & dr(1).toString() & "</a></p><div class=""img_box""><a href=""/noticias/" & dr(9).toString() & """ class=""crop-foto""><img src=""http://images.guiadoturista.net/?img=" & dr(7).toString() & "&amp;l=" & thumbWidth & "&amp;a=" & thumbHeight & """ width=""" & thumbWidth & """ height=""" & thumbHeight & """ /> </a></div><p class=""rp"">" & dr(8).toString()  & "</p></div></li>"
				end while
				htmlConteudo = htmlConteudo & "</ul></div>"
				if exibeBanner then htmlConteudo = htmlConteudo  & "<div class=""pub""><div>" & ads  & "</div></div>"
				htmlConteudo = htmlConteudo & "</div>"
				if cssPubAds then htmlConteudo = htmlConteudo & "</div>"				
				return htmlConteudo
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
	
	
end class