Imports Microsoft.VisualBasic
Imports System
Imports System.Type
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

Public Module _Kernel

	Public idCidade as Integer
	Public debug as boolean

	Public Class kernel

		Public slug as string
		Public pagina as string
		Public pgcat as string
		Public pgsubcat as String
		Public strSearch as String
		Public pgRef as string
		Public target as string
		Public nomeTemplate as String
			
		public function iniciar()
			Dim conn as SqlConnection 
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			Dim strCon as string = ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
			
			'debug = false
			'HttpContext.Current.response.write(idCidade)
			if (IsNothing(idCidade) or idCidade = 0) then error_404("WebSite Não Encontrado")
			if (IsNothing(slug) or slug = "") then error_404("Página Não Encontrada")
			if not (isNothing(pgRef) ) then pgURLRef = pgRef			
			
			nomeTemplate = iif_template(webNTemplate, dados_cidade(idCidade, "ntemplate"))
			nome_template = iif_template(webNTemplate, dados_cidade(idCidade, "ntemplate"))
						
			dim existe_area as boolean = false
			dim idmodulo as integer
			dim area as string	
			conn = New SqlConnection(strCon)
			try
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand("exec proc_check_modulo_init_netportal @id_cidade, @slug", conn )
				with cmd
					.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
					.Parameters("@id_cidade").Value = idcidade
					.Parameters.Add(New SqlParameter("@slug", SqlDbType.VarChar))
					.Parameters("@slug").Value = slug	
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					idmodulo = dr("id_modulo")
					area = dr("area").toString
					existe_area = true
				end if
				dr.close()
			Catch SQLExp As SqlException
				 Dim errorMessages As String
				 Dim i As Integer
		
				 For i = 0 To SQLExp.Errors.Count - 1
					 errorMessages += "Index #" & i.ToString() & ControlChars.NewLine _
									& "Message: " & SQLExp.Errors(i).Message & ControlChars.NewLine _
									& "LineNumber: " & SQLExp.Errors(i).LineNumber & ControlChars.NewLine _
									& "Source: " & SQLExp.Errors(i).Source & ControlChars.NewLine _
									& "Procedure: " & SQLExp.Errors(i).Procedure & ControlChars.NewLine
				 Next i
				error_SQL("<div style=""width:600px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu com o Banco de Dados: <b>" & ucase(SQLExp.Message.toString()) & "</b>, verifique a conexão ao Banco de Dados.<br><br> <b>Memsagem:</b><hr/>" & errorMessages & "</div>" )
			catch ex as exception
				error_404("<div style=""width:600px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Inicialização do Aplicativo: <b>" & ucase(ex.Message.toString()) & "</b>, verifique a conexão ao Banco de Dados.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
			finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			end try	
						
			if existe_area then
				
				pagina = _getURLContentPage(pagina)
													
				if ( Not IsNothing(pagina) and pagina <> "" ) or (Not IsNothing(pgcat) and pgcat <> "") then
					verifica_tabela(area, idmodulo, pagina, pgcat, pgsubcat)
					Exit Function
				'elseif ( isNumeric(pagina) ) then
				
					'error_404("Página: <b>" & idmodulo & "</b>, Numérica")				
				
				else
					
					target = "/App_Themas/" & nomeTemplate & "/paginas/" & area & "/home.aspx?id=" & idcidade & "&idmodulo=" & idmodulo	
					
					if (Not IsNothing(strSearch) and strSearch <> "" and area = "busca" ) then					
						target = "/App_Themas/" & nomeTemplate & "/paginas/" & area & "/home.aspx?id=" & idcidade & "&idmodulo=" & idmodulo & "&" & strSearch
					End If
					
					execHomePage(target, area)
					Exit Function
					
				end if
			else
				error_404("Página: <b>" & slug & "</b>, Não Encontrada")	
			end if				
		
		end function
		
		
		'verifica tabela correspondente quando conteudo textual
		private function verifica_tabela(area, id_modulo, pagina, categoria, subcategoria)
		
			'HttpContext.Current.response.write(id_modulo)						
			Dim sql_conteudo, tabela, campo_index, campo_url_conteudo, campo_idcidade_conteudo as String
			Dim sql_categoria, tabela_categoria, campo_index_categoria, campo_url_categoria, campo_idcidade_categoria as String	
			Dim pageCategoria as String
			
			dim conn as SqlConnection 
			dim cmd as SqlCommand
			dim dr As SqlDataReader
			dim strCon as string = ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString			
			conn = New SqlConnection(strCon)
			
			try	
				'verificamos os dados do modulo existente para consulta e descobrir se é pagina de conteudo ou categoria
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand( "select id, nome, tabela, campo_index, campo_url_conteudo, campo_idcidade_conteudo, tabela_categorias, campo_index_categoria, campo_url_categoria, campo_idcidade_categoria  from tbl_modulos_netportal where id = @id_modulo", conn )
				with cmd
					.Parameters.Add(New SqlParameter("@id_modulo", SqlDbType.Int))
					.Parameters("@id_modulo").Value = id_modulo			
				end with
	
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					tabela = dr(2).tostring()
					campo_index = dr(3).tostring()
					campo_url_conteudo = dr(4).tostring()
					campo_idcidade_conteudo = dr(5).tostring()
					tabela_categoria = dr(6).tostring()
					campo_index_categoria = dr(7).tostring()
					campo_url_categoria = dr(8).tostring()
					campo_idcidade_categoria = dr(9).tostring()
				end if
				dr.close()	
				If dr IsNot Nothing Then dr.Dispose()
				If cmd IsNot Nothing Then cmd.Dispose()			
			Catch ex As Exception
				error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Montagem do Sistema de Menus: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()

				'monta sql
				sql_conteudo = "select " & campo_index & " from " & tabela & " where " & campo_url_conteudo & "=@url and " & campo_idcidade_conteudo & "=@idcidade"
								
				if isNumeric(pagina) and id_modulo <> 11 then
					sql_conteudo = "select " & campo_index & " from " & tabela & " where " & campo_index & "=@url and " & campo_idcidade_conteudo & "=@idcidade"
				end if

				sql_categoria = "select " & campo_index_categoria & " from " & tabela_categoria & " where " & campo_url_categoria & "=@url and " & campo_idcidade_categoria & "=@idcidade"			
				
				if (campo_index_categoria = "0") and pagina = "" then
					pageCategoria = categoria				
					pagina = ""
					categoria = ""
				end if
				
				if (campo_idcidade_categoria = "0") then
					sql_categoria = "select " & campo_index_categoria & " from " & tabela_categoria & " where " & campo_url_categoria & "=@url"
				end if	
						
				if (Not IsNothing(pagina) and pagina <> "") then verifica_pagina(sql_conteudo, pagina, area, id_modulo)	
				if (Not IsNothing(categoria) and categoria <> "") then verifica_categoria(sql_categoria, categoria, area, id_modulo, subcategoria)	
				if (Not IsNothing(pageCategoria) and pageCategoria <> "") then verificaPageCategoria(sql_conteudo, pageCategoria, area, id_modulo)						
				
			End Try		
		
		end function
		
		
		'verifica a pagina do conteudo
		private function verifica_pagina(sql_conteudo, pagina, area, id_modulo)
			Dim existe_area as boolean = false
			Dim J As Integer 
			Dim idconteudo as String
			Dim Str as string = "SqlDbType.VarChar"
			If Integer.TryParse(pagina, j) then 
				Str = "SqlDbType.Int"
				pagina = CType(pagina, Integer) 
			end if	
			
			dim conn as SqlConnection 
			dim cmd as SqlCommand
			dim dr As SqlDataReader
			dim strCon as string = ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
			conn = New SqlConnection(strCon)
			
			
			'error_404("Página: <b>" & pagina & "</b>, Numérica")			
			
			try	
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand(sql_conteudo, conn )
				with cmd
					.Parameters.Add(New SqlParameter("@url", Str))
					.Parameters("@url").Value = pagina
					.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
					.Parameters("@idcidade").Value = idcidade		
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					idconteudo = dr(0).toString()
					existe_area = true
				end if
				dr.close()	
				If dr IsNot Nothing Then dr.Dispose()
				If cmd IsNot Nothing Then cmd.Dispose()			
			Catch ex As Exception
				error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro a Requisição da Página: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()		

				if existe_area then 
					if IsNumeric(idconteudo) then 
						execPageConteudo(area,idconteudo,id_modulo) 
					else
						if ( IsNumeric(id_modulo) and idConteudo <> "" ) then loadContentPage(area,idconteudo,id_modulo) else error_404("Página:" & pagina & ", Não Encontrada")
					end if
				else 
					error_404("Página:" & pagina & ", Não Encontrada")
				end if

			End Try		
	
		end function		
		
		'verifica a pagina nivel categoria
		private function verifica_categoria(sql, categoria, area, id_modulo, subcategoria)
			Dim idconteudo as integer
			dim conn as SqlConnection 
			dim cmd as SqlCommand
			dim dr As SqlDataReader
			dim strCon as string = ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
			conn = New SqlConnection(strCon)
	
			try	
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand(sql, conn )
				with cmd
					.Parameters.Add(New SqlParameter("@url", SqlDbType.varChar))
					.Parameters("@url").Value = categoria
					.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
					.Parameters("@idcidade").Value = idcidade		
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					idconteudo = dr(0).toString()
				end if
				dr.close()	
				If dr IsNot Nothing Then dr.Dispose()
				If cmd IsNot Nothing Then cmd.Dispose()			
			Catch ex As Exception
				error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante a Montagem do Sistema de Menus: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()				
			End Try
	
			If (Not IsNothing(idconteudo) and idconteudo <> 0) and ( IsNothing(subcategoria) and subcategoria = "" ) then
				execPageCategoria(area,idconteudo,id_modulo)
				Exit Function
			Else
				if area = "empresas" and subcategoria <> "" then  
					verificasubcat(area,idconteudo,id_modulo,subcategoria)
					Exit Function
				End If
				error_404("Página <u><strong>" & categoria & "</strong></u> não encontrada")
			End IF
	
		end function
		
		'verifica a pagina de subcategoria
		private function verificasubcat(area,idconteudo,id_modulo,subcategoria)
				
			Dim idsubcategoria as integer
			Dim conn as SqlConnection 
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			Dim strCon as string = ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
			conn = New SqlConnection(strCon)

			try	
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand("select idsub from subcategoria where url_amigavel = @url and idcategoria = @idconteudo", conn )
				with cmd
					.Parameters.Add(New SqlParameter("@url", SqlDbType.varChar))
					.Parameters("@url").Value = subcategoria
					.Parameters.Add(New SqlParameter("@idconteudo", SqlDbType.Int))
					.Parameters("@idconteudo").Value = idconteudo
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					idsubcategoria = dr(0).toString()					
				end if
				dr.close()
			Catch ex As Exception
				error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro a Requisição da Página: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
			Finally		
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
							
				if IsNumeric(idsubcategoria) then 
					if (Not IsNothing(idsubcategoria) and idsubcategoria <> 0) then				
						execPageSubCategoria(area,idconteudo,idsubcategoria,id_modulo)
					else
						error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" >Página Não Encontrada</div>" )	
					end if
				Else
					error_404("Página:" & pagina & ", Não Encontrada")
				End If

			End Try
			
		end function
		
		private function verificaPageCategoria(sql_conteudo, pagina, area, id_modulo)
			Dim J As Integer 
			Dim idconteudo as String
			Dim Str as string = "SqlDbType.VarChar"
			If Integer.TryParse(pagina, j) then 
				Str = "SqlDbType.Int"
				pagina = CType(pagina, Integer) 
			end if
			
			'HttpContext.Current.response.write(sql_conteudo)
	
			dim conn as SqlConnection 
			dim cmd as SqlCommand
			dim dr As SqlDataReader
			dim strCon as string = ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
			conn = New SqlConnection(strCon)
			
			try	
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand(sql_conteudo, conn )
				with cmd
					.Parameters.Add(New SqlParameter("@url", Str))
					.Parameters("@url").Value = pagina
					.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
					.Parameters("@idcidade").Value = idcidade		
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					idconteudo = dr(0).toString()
				end if
				dr.close()	
				If dr IsNot Nothing Then dr.Dispose()
				If cmd IsNot Nothing Then cmd.Dispose()			
			Catch ex As Exception
				error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro a Requisição da Página: <b>" & ucase(ex.Message.toString()) & "</b>, verifique se o mesmo existe no contexto.<br><br> <b>Memsagem:</b><hr/>" & ex.StackTrace.tostring() & "</div>" )
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
	
				if IsNumeric(idconteudo) then 
					if (Not IsNothing(idconteudo) and idconteudo <> "") then
						execPageConteudo(area,idconteudo,id_modulo)
					else
						error_404("<div style=""width:500px;border: solid 1px #666666; padding:10px; margin:0 auto;"" >Página Não Encontrada</div>" )	
					end if
				else
				
					error_404("Página:" & pagina & ", Não Encontrada") 				
				
				End If
				
				
			End Try					
	
		end function								

		
	end Class
	
	
	'classes auxiliares construtoras
	Public Class _metaHeader
		
		Public Function meta(idcidade as integer, idmodulo as integer, meta_name as string) As String  
			dim retorno as string = metadados(idcidade, idmodulo, meta_name)
			Return retorno
		End Function
		
		public function metadados(idcidade as integer, idmodulo as integer, coluna as string) as string
				Dim strSQL as String
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
	
	'Monta o Menu do Topo do Layout Novo 
	Public Class _topNavMenu
		Dim contar as Integer = 1
		Dim extra, html as String
	
		'Monta Menu do Top de Navegação
		public function bind() as string
			Dim strSQL as String
			Dim conn as SqlConnection
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			try
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand( "select nome, titulo, url, id, ordem from tbl_menu_top_cidades where id_cidade = @id_cidade and ativo = 1 order by ordem asc;", conn )
				with cmd
					.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
					.Parameters("@id_cidade").Value = idcidade
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					html = "<ul id=""jsddm"">"
					While dr.Read()
						extra = ""
						if contar = 1 then
							extra = "class=""primeiro"""
						end if
						html = html & "<li " & extra & " ><a href=""" & iif_url(webUrlCidade, dr(2).toString() ) & """ accesskey=""" & contar & """ class=""h-nav"" title=""" & dr(1) & """>" & dr(0) & "</a>"  & subNav(dr(3)) &  "</li>"
						contar += 1
					End While					
					html = html & "</ul>"
				end if
				dr.close()
				cmd.dispose()		
			Catch ex as exception
				
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()		
			end try
			
			return html
			
		end function
		
		'Monta o Sub Menu do Topo relacionado ao Menu do Layout
		private function subNav(submenu as integer) as string
			Dim html as string
			Dim strSQL as String
			Dim conn as SqlConnection
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			try
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand( "select nome, titulo, url, id, destacar from tbl_sub_menu_top_cidades where id_menu = @menu and ativo = 1 order by ordem asc;", conn )
				with cmd
					.Parameters.Add(New SqlParameter("@menu", SqlDbType.Int))
					.Parameters("@menu").Value = submenu
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					html = "<ul class=""children"">"
					While dr.Read()
						html = html & "<li><a href=""" & iif_url(webUrlCidade, dr(2).toString() )  & """ title=""" & dr(1) & """>"
							if dr(4) = 1 then 
								html = html & "<strong>" & dr(0) & "</strong>"
							else 
								html = html & dr(0)
							end if
						html = html & "</a></li>"
					End While
					html = html & "</ul>"
				end if
			catch ex as exception
			
			finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			end try
			return html
			
		end function
			
			
	End Class
	
	'Monta Menu Lateral Layout Full Home Cidades
	Public Class _leftNavMenu
		Dim html as String
		'Monta Menu Lateral Home de Navegação
		public function bind() as string
			Dim strSQL as String
			Dim conn as SqlConnection
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			
			try
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand( "select nome, titulo, url, id, destaque, bloquear_robo, nova_janela from tbl_menu_home_lateral_cidades where id_cidade = @id_cidade and ativo = 1 order by ordem asc;", conn )
				with cmd
					.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
					.Parameters("@id_cidade").Value = idcidade
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					html = "<ul>"
					While dr.Read()
						html = html & "<li><a href=""" & iif_url(webUrlCidade, dr(2).toString() ) & """ title="""& dr(1)&""""
						if dr(5) = 1 and dr(6) = 0 then html = html &  " rel=""nofollow"""
						if dr(6) = 1 and dr(5) = 0 then html = html & " rel=""nova_janela"""
						if dr(6) = 1 and dr(5) = 1 then html = html & " rel=""nofollow novajanela"""
						if dr(4) = 1 then 
							html = html & " class=""item-destaque""><strong>" & dr(0) & "</strong></a></li>"
						else
							html = html + ">" & dr(0) & "</a></li>"
						end if
					End While					
					html = html & "</ul>"
				end if
				dr.close()
				cmd.dispose()		
			Catch ex as exception
				
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()		
			end try
			
			return html		
			
		end function
	
	end Class	
	
end Module