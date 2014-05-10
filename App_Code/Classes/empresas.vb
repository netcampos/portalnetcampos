Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.MySql.Data
Imports System.MySql.Data.MySqlClient
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
Imports System.Web.UI.UserControl

Public Class empresa

	Public shared function dados(ByVal idEmpresa as integer, ByVal coluna as String) as String
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
	
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec np_DadosEmpresa @idEmpresa", conn)
			cmd.Parameters.Add(New SqlParameter("@idEmpresa", SqlDbType.Int))
			cmd.Parameters("@idEmpresa").Value = idEmpresa
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
	
	'retorna thumb da empresa
	Public Shared Function thumbEmpresa(ByVal img as string, largura as integer, altura as integer) as string
		if (not IsNothing(img) and img <> "") then
			if img <> "" then
				return "http://images.guiadoturista.net/?img=" & img & "&amp;l=" & largura.toString() & "&amp;a=" & altura.toString() 
			else
				return "/App_Themas/images/empresa-sem-foto.gif"		
			end if
		else
			return "/App_Themas/images/empresa-sem-foto.gif"
		end if
					
	End Function
	
	'retorna dados ficha hospedagem
	Public Function dadosHospedagem(ByVal idEmpresa as Integer, coluna as String) as String
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
	
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select * from tbl_hospeda where idcliente = @idEmpresa", conn)
			cmd.Parameters.Add(New SqlParameter("@idEmpresa", SqlDbType.Int))
			cmd.Parameters("@idEmpresa").Value = idEmpresa
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
	
	Public Function verificaServico(ByVal idEmpresa as Integer, Byval servico as Integer, ByVal separador as Integer) as String
	
		'separador
		' 1 = Listas
		' 2 = ,(virgua)
		' 3 = div
		' 4 = class
	
		Dim html as String = String.Empty
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select s2.servico from tbl_servico_empresa as s1 inner join tbl_servicos as s2 on s1.idservico = s2.idservicos inner join empresas as e on s1.idcliente = e.idempresa where e.idempresa = @idempresa and s2.onde = @servico", conn)
			with cmd
				cmd.Parameters.Add(New SqlParameter("@idEmpresa", SqlDbType.Int))
				cmd.Parameters("@idEmpresa").Value = idEmpresa
				cmd.Parameters.Add(New SqlParameter("@servico", SqlDbType.Int))
				cmd.Parameters("@servico").Value = servico				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				if separador = 1 then
					html = "<ul class=""bullets"">"
					While dr.Read()
						html = html & "<li>" & dr(0).toString() & "</li>"
					End While
					html = html & "</ul>"
				elseif separador = 2 then
					While dr.Read()
						html = html & "" & dr(0).toString() & ", "
					End While
				elseif separador = 3 then
					While dr.Read()
						html = html & "<div>" & dr(0).toString() & "</div>"
					End While
				elseif separador = 4 then
					While dr.Read()
						html = html & "<div class=""" & tiraAscento(removeEspacao(lcase(dr(0).toString()))) & " cartoes""></div>"
					End While					
				end if
			end if
			dr.close
		Catch ex As Exception
			html = ex.Message().toString()
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try	
		
		return html		
	
	End Function
	
	Public Function aceitaCartao(ByVal idEmpresa as Integer, Byval servico as Integer) as Boolean

		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_servico_empresa as s1 inner join tbl_servicos as s2 on s1.idservico = s2.idservicos inner join empresas as e on s1.idcliente = e.idempresa where e.idempresa = @idempresa and s2.onde = @servico", conn)
			with cmd
				cmd.Parameters.Add(New SqlParameter("@idEmpresa", SqlDbType.Int))
				cmd.Parameters("@idEmpresa").Value = idEmpresa
				cmd.Parameters.Add(New SqlParameter("@servico", SqlDbType.Int))
				cmd.Parameters("@servico").Value = servico
			end with
			if cmd.ExecuteScalar() > 0 then
				return true
			end if
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try
	
	End Function	
	
	Public function _getFotos(ByVal idEmpresa as Integer) as Integer
		Dim totalFotos as Integer
		Dim cmd As SqlCommand
		Dim conn As SqlConnection
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(f.idfoto) from tbl_galeria_foto_empresas as f where f.id_empresa = @idempresa", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idempresa", SqlDbType.Int))
				.Parameters("@idempresa").Value = idEmpresa
			end with
			totalFotos = cmd.ExecuteScalar()
		Catch ex As Exception
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try

		return totalFotos

	end function
	
	Public Function _getTotalFotosInternautas(ByVal idEmpresa as Integer) as Integer
	
		Dim totalFotos as Integer
		Dim cmd As SqlCommand
		Dim conn As SqlConnection
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(f.id) from tbl_fotos_clientes_netportal as f inner join tbl_album_fotos_clientes as a on a.id = f.idAlbum where a.idConteudo = @idempresa and a.idModulo = 3", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idempresa", SqlDbType.Int))
				.Parameters("@idempresa").Value = idEmpresa
			end with
			totalFotos = cmd.ExecuteScalar()
		Catch ex As Exception
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try

		return totalFotos	
	
	End Function
	
	Public Function _isCliente(ByVal idEmpresa as Integer) as Boolean
	
		Dim cmd As SqlCommand
		Dim conn As SqlConnection
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from cliente as c inner join empresas as e on e.idcliente = c.idcliente where e.bloqueado = 0 and c.ativo = 1 and e.idempresa = @idempresa", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idempresa", SqlDbType.Int))
				.Parameters("@idempresa").Value = idEmpresa
			end with
			if cmd.ExecuteScalar() > 0 then
				return true
			end if
		Catch ex As Exception
		
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try
	
	End Function

End Class

Public Class _catEmpresas

	Public Shared Function metaCategoria(ByVal idCidade as integer, ByVal idCategoria as Integer, ByVal Coluna as String) as String
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
	
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec np_empresasMetaHeadCategoria @idCidade, @idCategoria", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idCategoria", SqlDbType.Int))
				.Parameters("@idCategoria").Value = idCategoria
			end with
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


	Public Shared Function metaSubCategoria(ByVal idCidade as integer, ByVal idCategoria as Integer, ByVal idSubCategoria as Integer, ByVal Coluna as String) as String

		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	
	
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec np_empresasMetaHeadSubCategoria @idCidade, @idCategoria, @idSubCategoria", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idCategoria", SqlDbType.Int))
				.Parameters("@idCategoria").Value = idCategoria
				.Parameters.Add(New SqlParameter("@idSubCategoria", SqlDbType.Int))
				.Parameters("@idSubCategoria").Value = idSubCategoria				
			end with
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
	
	'retorna dataTable SubCategoria Relacionada
	Public function _subCategorias(ByVal idCidade as Integer, ByVal idCategoria as Integer) as System.Data.DataTable						
		Dim cmd As SqlCommand
		Dim conn As SqlConnection
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand( "exec proc_lista_subcat_empresas @idcategoria, @idcidade", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idcategoria", SqlDbType.Int))
				.Parameters("@idcategoria").Value = idCategoria
				.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
				.Parameters("@idcidade").Value = idCidade
			end with
			
			Dim daSubCat As sqlDataAdapter = New sqlDataAdapter(cmd)
			Dim dtSubCat As DataTable = New DataTable()
			daSubCat.Fill(dtSubCat)
			return dtSubCat
			
		Catch ex As Exception
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try
	
	End Function
	
	'retorna dataTable c/ Bairros Relacionados com a Categoria
	Public function _Bairros(ByVal idCidade as Integer, ByVal idCategoria as Integer) as System.Data.DataTable						
		Dim cmd As SqlCommand
		Dim conn As SqlConnection
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand( "exec proc_empresas_bairro_netportal @idcidade, @idcategoria", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
				.Parameters("@idcidade").Value = idCidade			
				.Parameters.Add(New SqlParameter("@idcategoria", SqlDbType.Int))
				.Parameters("@idcategoria").Value = idCategoria
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

'### Tab + Ficha Empresa #####"
Public Class _TabEmpresa
	
	Public idEmpresa, _idCategoria, totalFotos as Integer
	Public loadFicha, loadPageControl, txtMapa, totalResenhas as String
	Public tourVirtual, mapaVirtual, temMail, temFotos, temFicha as Boolean
	
	Public Function _renderTabs() as String
		Dim Empresa as New Empresa()	
		'Verifica Modulos
		'verifica classificacao
		If _idCategoria = 20 or _idCategoria = 12 then
			txtMapa = empresa.dadosHospedagem(idEmpresa, "localizacao")
			temFicha = _temFicha("tbl_hospeda")
			loadPageControl = "~/App_Modules/empresas/controls/formMail/formMailHospedagem.ascx"
		elseif _idCategoria = 25 then 
			loadPageControl = "~/App_Modules/empresas/controls/formMail/formMailHospedagem.ascx"
		elseif _idCategoria = 13 then 
			loadPageControl = "~/App_Modules/empresas/controls/formMail/formMailPadrao.ascx"
		else
			loadPageControl = "~/App_Modules/empresas/controls/formMail/formMailPadrao.ascx"
		End If
		'Tour Virtual
		if (Not IsNothing( empresa.dados(idEmpresa, "embed_tour") ) And empresa.dados(idEmpresa, "embed_tour") <> "" ) then tourVirtual = true
		'Verifica se tem Mapa
		if (Not IsNothing( empresa.dados(idEmpresa, "latitude") ) And empresa.dados(idEmpresa, "latitude") <> "" ) And (Not IsNothing( empresa.dados(idEmpresa, "longitude") ) And empresa.dados(idEmpresa, "longitude") <> "" ) then mapaVirtual = true
		
		'Verifica Email
		'temMail = false
		temMail = _getMail()
		
		totalFotos = empresa._getFotos(idEmpresa)
		totalFotos = totalFotos + empresa._getTotalFotosInternautas(idEmpresa)
		if totalFotos > 0 then temfotos = true
		'Resenhas
		totalResenhas = _getReviewConteudo(idCidade, 3, idEmpresa, "avaliacoes")
		
		Dim html as String
		html = html & "<div id=""Tabs"" class=""gnc-ui-tabs"">"
		html = html & _getNavTabs()
		html = html & _getContentTabs()
		html = html & "</div>"

		return html
	
	End Function
	
	Private Function _getNavTabs() as String
		Dim html as String
		html = html & "<div id=""navTabs"" class=""gnc-ui-nav-tabs"">"
			html = html & "<ul>"
				'Ficha Empresa
				if temFicha then
					html = html & "<li><a class=""active gnc-ui-tab""  href=""#tabInfo"">Detalhes</a></li>"
					html = html & "<li><a class=""inative gnc-ui-tab tabReview"" href=""#tabReview"">Resenhas (" & totalResenhas & ")</a></li>"					
				else
					html = html & "<li><a class=""active gnc-ui-tab tabReview"" href=""#tabReview"">Resenhas (" & totalResenhas & ")</a></li>"				
				end if
				'Get Fotos Empresa
				if temFotos then
					html = html & "<li><a class=""inative gnc-ui-tab"" href=""#tabFotos"">Fotos (" & totalFotos & ")</a></li>"
				End If				
				'Get Tour Virtual
				if tourVirtual then
					html = html & "<li><a class=""inative gnc-ui-tab"" href=""#tabTour"">Tour Virtual</a></li>"
				End If
				'Get Mapa Virtual
				If mapaVirtual then
					html = html & "<li><a class=""inative gnc-ui-tab tabMap"" href=""#tabMapa"">Mapa de Localização</a></li>"
				End If
				'Get Form Mail
				If temMail then
					html = html & "<li><a class=""inative gnc-ui-tab"" href=""#tabMail"">Enviar Email</a></li>"
				End If
				
			html = html & "</ul>"
		html = html & "</div>"
								
		return html
	End Function
	
	
	Private Function _getContentTabs() as String
		Dim html as String
		html = html & "<div id=""ContentTabs"" class=""gnc-ui-content-tabs"">"
							
			If temFicha then
				loadFicha = "~/App_Modules/empresas/controls/pageEmpresa/fichaHospedagem.ascx"
				html = html & "<div id=""tabInfo"" class=""gnc-ui-content-tab on"">"
					html = html & "<h3 class=""gnc-ui-title-ficha-empresa"">Informações Sobre: " & empresa.dados(idEmpresa, "empresa") & "</h3>"
					html = html & "<div>" & Me.GenerateControlMarkup(loadFicha) & "</div>"
				html = html & "</div>"
				html = html & "<div id=""tabReview"" class=""gnc-ui-content-tab off"">"
					html = html & "<div class=""UIStorycomments"">"
						html = html & CheckResenhasEmpresa(idCidade, 3, idEmpresa)
						html = html & Me.GenerateControlMarkup("~/App_Modules/comments/controls/formResenha.ascx")
					html = html & "</div>"
				html = html & "</div>"
			Else
				html = html & "<div id=""tabReview"" class=""gnc-ui-content-tab on"">"
					html = html & "<div class=""UIStorycomments"">"
						html = html & CheckResenhasEmpresa(idCidade, 3, idEmpresa)
						html = html & Me.GenerateControlMarkup("~/App_Modules/comments/controls/formResenha.ascx")
					html = html & "</div>"
				html = html & "</div>"
			End If
			'Get Fotos Empresa
			if temFotos then
				Dim loadPageFoto as String = "~/App_Modules/empresas/controls/fotosEmpresa.ascx"
				html = html & "<div id=""tabFotos"" class=""gnc-ui-content-tab off"">"
				html = html & "<div class=""gnc-ui-title-ficha-empresa""><div class=""gnc-ui-div-left""><h3>Galeria de Fotos</h3></div><div class=""gnc-ui-div-right""><a href=""http://api.portalguiadoturista.com.br/swf/slide/empresas/3dwall/?id=" & idEmpresa & """ id=""actionCooliris"" class=""SlideShowButton"" title=""Super Slide Show""><span>Super Slide Show</span></a></div></div>"
				html = html & "<div class=""gnc-ui-fotos-empresa-containner"">"
				html = html & Me.GenerateControlMarkup(loadPageFoto)
				html = html & "</div>"
				html = html & "</div>"
			End If
			'Get Tour Virtual
			if tourVirtual then
				html = html & "<div id=""tabTour"" class=""gnc-ui-content-tab off"">"
				html = html & "<h3 class=""gnc-ui-title-ficha-empresa"">Tour Virtual</h3>"
				html = html & empresa.dados(idEmpresa, "embed_tour") & "</div>"
			End If
			'Get Mapa Virtual
			If mapaVirtual then			
				html = html & "<div id=""tabMapa"" class=""gnc-ui-content-tab off"">"
				html = html & "<div class=""gnc-ui-title-ficha-empresa""><div class=""gnc-ui-div-left""><h3>Mapa de Localização</h3></div><div class=""gnc-ui-div-right""><a href=""" & city.dados(idCidade, "urlCidade") & "/mapa/empresa/" & empresa.dados(idEmpresa, "urlEmpresa") & ".html#gdir"" id=""actionMapa"" class=""SlideShowButton"" title=""Ver Mapa Ampliado""><span>Mapa Ampliado</span></a></div></div>"
				html = html & "<div class=""gnc-ui-t14-p""><p>" & txtMapa & "</p></div>"
				html = html & "<iframe id=""mapa-empresa"" src=""" & city.dados(idCidade, "urlCidade") & "/mapa/empresa/" & empresa.dados(idEmpresa, "urlEmpresa") & ".html"" width=""100%"" style=""display:none"" height=""500"" scrolling=""no"" frameborder=""0""></iframe>"
				html = html & "</div>"
			End If
			'Get Form Mail
			If temMail then						
				html = html & "<div id=""tabMail"" class=""gnc-ui-content-tab off"">"
				html = html & "<h3 class=""gnc-ui-title-ficha-empresa"">Entre em contato com: " & empresa.dados(idEmpresa, "empresa") & "</h3>"				
				html = html & Me.GenerateControlMarkup(loadPageControl)
				html = html & "</div>"
			End If
		html = html & "</div>"
								
		return html
	End Function
	
	Private Function _getMail() as Boolean
		dim email as string = empresa.dados(idEmpresa, "email")
		if validaEmailSingle(email) then
			return true
		end if
	End Function
	
	Private Function GenerateControlMarkup(ByVal virtualPath As String) As [String]
		Dim page As New SacrificialMarkupPage()
		Dim ctl As UserControl = DirectCast(page.LoadControl(virtualPath), UserControl)
		page.Controls.Add(ctl)
		Dim sb As New StringBuilder()
		Dim writer As New StringWriter(sb)
	
		page.Server.Execute(page, writer, True)
		Return sb.ToString()
	End Function
	
	Private Function _temFicha(ByVal tabela as String) as boolean
	
		Dim cmd As SqlCommand
		Dim conn As SqlConnection
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from " & tabela & " where ativo = 1 and idcliente = @idempresa", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idempresa", SqlDbType.Int))
				.Parameters("@idempresa").Value = idEmpresa
			end with
			if cmd.ExecuteScalar() = 1 then return true
		Catch ex As Exception
		
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()
			If conn IsNot Nothing Then conn.Dispose()
		End Try
	
	End Function			

End Class

Public Class SacrificialMarkupPage
    Inherits Page
    Public Overloads Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
    End Sub
End Class
