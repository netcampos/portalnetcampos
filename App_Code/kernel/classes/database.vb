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
Public Class database

	'public members
	Public idEnquete as Integer
	'Private members
	Private idCidade as Integer = ConfigurationSettings.AppSettings("id_cidade")
	Private siteURL as string = "http://www.netcampos.com"

	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Sub New()
		MyBase.New()
		idEnquete = 0
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
	
	'Private Property
	Private Property sqlServer() As String
		Get
			Return ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property	
	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function infoUser(byVal id as integer, byval coluna as string) as String
		dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from usuarios where id_usuario = @id ", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose()    
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return retorno
   	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	public function topNavMenu() as string
		Dim contar as Integer = 1
		Dim extra, html as String
			
		Dim strSQL as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
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
					html = html & "<li " & extra & " ><a href=""" & dr(2).toString() & """ class=""h-nav"" title=""" & dr(1) & """>" & dr(0) & "</a>"  & subNav(dr(3)) &  "</li>"
					contar += 1
				End While					
				html = html & "</ul>"
			end if
			dr.close()
			cmd.dispose()		
		Catch ex as exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()		
		end try
		
		return html
		
	end function	

	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	private function subNav(submenu as integer) as string
		Dim html as string
		Dim strSQL as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
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
			html = ""
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function verticalNavHome() as string
		Dim html as string
		Dim strSQL as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		
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
					html = html & "<li><a href=""" & dr(2).toString() & """ title="""& dr(1)&""""
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
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()		
		end try		
		
		return html
	end Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function ads(idArea as integer, idLocal as integer) as string
		Dim html as string
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		conn = New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_adserver_banners_cidades_netportal where idcidade = @id_cidade and idarea = @id_area and idlocalentrega = @id_local and ativo = 1", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@id_area", SqlDbType.Int))
				.Parameters("@id_area").Value = idArea
				.Parameters.Add(New SqlParameter("@id_local", SqlDbType.Int))
				.Parameters("@id_local").Value = idLocal
			end with
			dim t as integer = cmd.ExecuteScalar()
			if t = 0 then		
				html = exibeBannerGoogle(idArea)
			else
				html = exibeBannerCliente(idArea, idLocal)
			end if
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
		End Try
		
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	private function exibeBannerGoogle(idArea as Integer) as string
		Dim googleads as integer
		Dim html as String
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select google_ads, code_google_ads, banner_padrao, txt_banner_padrao from tbl_adserver_areas_cidades_netportal where idcidade = @id_cidade and idarea = @id_area", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@id_area", SqlDbType.Int))
				.Parameters("@id_area").Value = idArea
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				googleads = dr(0)
				if googleads = 1 then														
					html  = dr(1).toString()
				end if				
			end if
			dr.close
		Catch ex As Exception
			html = ""		
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	private function exibeBannerCliente(idArea as Integer, idLocal as Integer) as string
		Dim contar as Integer = 0
		Dim googleads as integer
		Dim html as string
			
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select md5, googleads from tbl_adserver_banners_cidades_netportal where idcidade = @id_cidade and idarea = @id_area and idlocalentrega = @id_local and ativo = 1 order by newid()", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@id_area", SqlDbType.Int))
				.Parameters("@id_area").Value = idArea
				.Parameters.Add(New SqlParameter("@id_local", SqlDbType.Int))
				.Parameters("@id_local").Value = idLocal				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				googleads = dr(1)
				if googleads = 1 then
					html = exibeBannerGoogle(idArea)
				else
					html = showAd(dr(0).toString())
				end if			
			end if
			dr.close
		Catch ex As Exception
			return ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		
		return html
		
	end function		

	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Private function showAd(ByVal id as String) as string
		dim swf as string
		Dim extensao as string
		Dim arquivo as string
		Dim largura as string
		Dim altura as string
		Dim urlDestino as string
		Dim txtUrl as string
		Dim html as string
		
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())		
		
		try
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand( "select b.arquivo, b.tipomidia, a.largura, a.altura, b.urlDestino, b.txtUrl from tbl_adserver_banners_cidades_netportal as b inner join tbl_adserver_areas_netportal as a on a.id = b.idarea where b.md5 = @md5", conn )
			cmd.Parameters.Add(New SqlParameter("@md5", SqlDbType.VarChar))
			cmd.Parameters("@md5").Value = id
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				arquivo = dr(0).toString
				extensao = "." & dr(1).toString
				urlDestino = dr(4).toString
				txtUrl = dr(5).toString
				
				largura = dr(2).toString
				altura = dr(3).toString
				
				if dr(1).toString = "swf" then
					arquivo = replace(arquivo,extensao,"")
					arquivo = replace(arquivo,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
					arquivo = "http://images.netcampos.com/" & arquivo & extensao
					html = _lib.playerSWF( arquivo, largura, altura, "clickTAG=http://adserver.guiadoturista.net/?catads=2%26ads=" & id, id.toString)
				end if
				
				if dr(1).toString = "gif" then
					arquivo = replace(arquivo,extensao,"")
					arquivo = replace(arquivo,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
					arquivo = "http://images.netcampos.com/" & arquivo & extensao
					html = "<a href=""http://adserver.guiadoturista.net/?catads=2&amp;ads=" & id & """ title=""" & txtUrl & """ target=""_blank""><img src=""" & arquivo & """ alt=""" & txtUrl & """ width=""" & largura & """ height=""" & altura & """ /></a>"	
				end if				
				
			end if
			dr.close()
		catch ex as exception
			html = ""
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try				
	
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function NetPortalAdWords() as string	
		Dim html as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())		
	
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select top 3 idempresa, empresa, txtdescritivo, url, chave_md5 from empresas where idcidade = @id_cidade and idcliente > 0 and bloqueado = 0 and plano in(5,6) order by newid()", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = "<div id=""links-patrocinados"">"
				html = html & "<p>Links Patrocinados</p>"
				html = html & "<div id=""links-patrocinados-ads"">"
				html = html & "<ul class=""ads-links"">"
				While dr.Read()
					html = html & "<li>"
					html = html & "<div class=""ad-link""><a href=""http://adserver.guiadoturista.net/?catads=1&amp;ads=" & dr(4).tostring() & """ target=""_blank""><span>" & dr(1).toString() & "</span></a></div>"
					html = html & "<div class=""adb-link"">" & mid(dr(2).tostring(),1,90) & "...</div>"
					html = html & "</li>"
				end while
				html = html & "</ul></div></div>"
			end if
			dr.close
		Catch ex As Exception
			html = html & ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		
		return html
	
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function tituloEnquete() as String
		Dim titulo as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())

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
			return ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		
		return titulo
	end function 	

	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function perguntasEnquete() as string
		Dim html as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())

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
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try		
		
		return html	
	
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function noticiasHome(qtd as integer) as String
		dim uri as string
		Dim contar as Integer = 1
		dim html as string = string.empty
		Dim strSQL, titulo as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()		
			cmd = New SqlCommand("exec np_noticiasHome @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				While dr.Read()
					uri = "/" & dr("slugCat").toString() & "/" & year(dr("dataPublicacao").toString()) & "/" & right( "0" & month(dr("dataPublicacao")) ,2) & "/" & dr("slugPage").toString() & ".html"
					
					html = html & "<div class=""n-lista"">"
						html = html & "<p class=""n-Data""><a href=""/" & dr("slugCat").toString() & "/" & dr("slugPageCat").toString() & "/"">"&dr("cat").toString()& "</a> - " & databr(dr("dataPublicacao").toString()) & "</p>"
						html = html & "<h3><a title="""& trim(replace(dr("titulo").toString(),"&","&amp;")) &""" href=""" & uri & """>" & replace(dr("titulo").toString(),"&","&amp;") & "</a></h3>"
					html = html & "</div>"
					if contar < qtd then html = html & "<div class=""linha""></div>"
					contar += 1
				End While
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function featuredNews(qtd as integer) as String
		dim html as string = string.empty
		Dim strSQL, titulo, uri as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()		
			cmd = New SqlCommand("exec np_destaques_campos_do_jordao @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<div id=""gnc-featured"">"
					html = html & "<div class=""gnc-inner"">"
						html = html & "<div id=""gnc-Slider"">"
						While dr.Read()
							uri = dr("link").toString()
							html = html & "<div class=""slider-inner"">"
								html = html & "<figure>"
									html = html & _lib.fotoLegenda( dr("imagem").toString(), dr("titulo").toString(), uri, 350, 255, dr("alias").toString(), "") & vbnewline
								html = html & "</figure>"
								html = html & "<p><a href=""" & uri & """ title=""" & dr("titulo").toString() & """><i>" & dr("titulo").toString() & ".</i></a></p>"								
							html = html & "</div>"
						End While
						html = html & "</div>"
					html = html & "</div>"
					
					html = html & "<div id=""gnc-buttons"">"
						html = html & "<div class=""bg"">"
							html = html & "<div class=""prev""><a href=""#"" title=""Anterior""><span>« Anterior</span></a></div>"
							html = html & "<div class=""next""><a href=""#"" title=""Próximo""><span>Próximo »</span></a></div>"
							html = html & "<ul id=""nav""></ul>"							
						html = html & "</div>"
					html = html & "</div>"
				html = html & "</div>"
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html			
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************		
	Public Function lastFeaturedNews(qtd as integer) as string
		Dim contar as Integer = 0
		Dim slugPage as String = dadosModulo(pageModule, "slug")
		Dim html as string = string.empty
		Dim strSQL, titulo, uri as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())		
		try	
			conn.Open()		
			cmd = New SqlCommand("exec np_noticias_cidades @idCidade, 1, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<div class=""not-destaques""><h2>Destaques de Campos do Jordão</h2>"
				html = html & "<ul>"
				While dr.Read()
				
					uri = "/" & slugPage & "/" & year(dr("data_publicacao")) & "/" &  right( "0" & month(dr("data_publicacao")) ,2)  & "/" & dr("slugConteudo").toString() & ".html"
				
					html = html & noPadding(contar, qtd)
						html = html & "<figure>"
							html = html & _lib.fotoLegenda( dr("picture").toString(), dr("titulo").toString(), uri, 168, 122, dr("categoria").toString(), "") & vbnewline
						html = html & "</figure>"
						html = html & "<h3><span class=""datanot"">" & databr(dr("data").toString()) & "</span> - <a href=""" & uri & """ title=""" & dr("titulo").toString() & """ >" & dr("titulo").toString() & "</a></h3>"						
					html = html & "</li>"
					contar = contar + 1
				End While
				html = html & "</ul>"
				html = html & "</div>"
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose() 
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function topNews(qtd as integer) as string
		Dim contar as Integer = 0
		Dim slugPage as String = dadosModulo(pageModule, "slug")
		Dim html as string = string.empty
		Dim strSQL, titulo, uri as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())		
		try	
			conn.Open()		
			cmd = New SqlCommand("exec np_noticias @idCidade,2,@qtd,4", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				
				html = html & "<div class=""widget-titulo"">"
					html = html & "<h2>Últimas Notícias</h2>"
					html = html & "<a class=""rss"" href=""http://www.netcampos.com/rss/""><img src=""http://s.glbimg.com/jo/g1/media/common/img/icones/icoRss.gif""></a>"
				html = html & "</div>"
				
				html = html & "<ul>"
				While dr.Read()
				
					uri = "/" & slugPage & "/" & year(dr("data_publicacao")) & "/" &  right( "0" & month(dr("data_publicacao")) ,2)  & "/" & dr("url_simples").toString() & ".html"
					
					html = html & noPadding(contar, qtd)
						html = html & "<figure>"
							html = html & _lib.fotoLegenda( dr("imagem_destaque").toString(), dr("titulo").toString(), uri, 115, 85, dr("cat").toString(), "") & vbnewline
						html = html & "</figure>"
						html = html & "<div class=""inner"">"
							html = html & "<p class=""datanot"">" & databr(dr("data").toString()) & "</p>"
							html = html & "<p class=""news""><a href=""" & uri & """>" & dr("titulo").toString() & "</a></p>"
						html = html & "</div>"
					html = html & "</li>"
					contar = contar + 1
				End While
				html = html & "</ul>"
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Function maisNoticias(qtd as integer) as string
		Dim contar as Integer = 0
		Dim slugPage as String = dadosModulo(pageModule, "slug")
		Dim html as string = string.empty
		Dim strSQL, titulo, uri as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())		
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()		
			cmd = New SqlCommand("exec np_noticias @idCidade,8,@qtd,4", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul>"
				While dr.Read()
										
					uri = "/" & slugPage & "/" & year(dr("data_publicacao")) & "/" &  right( "0" & month(dr("data_publicacao")) ,2)  & "/" & dr("url_simples").toString() & ".html"
					
					html = html & noPadding(contar, qtd)
						html = html & "<figure>"
							html = html & _lib.fotoLegenda( dr("imagem_destaque").toString(), dr("titulo").toString(), uri, 115, 85, dr("cat").toString(), "") & vbnewline
						html = html & "</figure>"
						html = html & "<div class=""inner"">"
							html = html & "<p class=""datanot"">" & databr(dr("data").toString()) & "</p>"
							html = html & "<p class=""news""><a href=""" & uri & """>" & dr("titulo").toString() & "</a></p>"
							html = html & "<a class=""read-more"" title=""Continue Lendo"" href=""" & uri & """>continue lendo...</a>"
						html = html & "</div>"
					html = html & "</li>"
					contar = contar + 1
				End While
				html = html & "</ul>"
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************		
	Public Function catNoticias() as string
		Dim contar as Integer = 0
		Dim slugPage as String = dadosModulo(pageModule, "slug")
		Dim html as string = string.empty
		Dim strSQL, titulo, uri as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())		
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()		
			cmd = New SqlCommand("exec cat_noticias_netportal @idCidade", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul>"
					While dr.Read()
						html = html & "<li>"
							html = html & "<a href=""" & dr("url") & """ title=""" & dr("cat") & """>" & dr("cat") & "</a>"
						html = html & "</li>"
					End While
				html = html & "</ul>"
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getNoticia(byVal id as integer, byval coluna as string) as String
		dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _netportal_noticias where id = @id ", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return retorno
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getNoticiaSlug(byVal slug as string, byval coluna as string) as String
		dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _netportal_noticias where slug = @slug", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@slug", SqlDbType.VarChar))
				.Parameters("@slug").Value = slug				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return retorno
   	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function listNoticias(byVal qtd as integer, byval notid as integer) as String
		Dim html, uri as string
		Dim data As DateTime
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top " & qtd & " * from _cms_view_cidades_noticias where idCidade = @idCidade and id <> (@id) and ativo = 1 order by id desc;", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade			
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = notid
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<div id=""noticias-relacionadas"">" & vbnewline
					html = html & "<ul class=""slides_container"">" & vbnewline
					while dr.Read()
						data = dr("dataPublicacao").toString()
						
						uri = "/noticias-campos-do-jordao/" & year(data) & "/" & Right("0" & month(data),2) & "/" & dr("seoSlug").toString() & ".html"
						
						html = html & "<li class=""list-container"">"
							html = html & "<div class=""gnc-header"">"
								html = html & "<figure>"
									html = html & _lib.fotoLegenda( dr("thumb").toString(), dr("titulo").toString(), uri, 175, 115, _lib.databr(data), "") & vbnewline
								html = html & "</figure>"							
							html = html & "</div>"
							
							html = html & "<div class=""gnc-containner"">"
								html = html & "<h3><a href=""" & uri & """>" & dr("titulo").toString() & "</a></h3>"
							html = html & "</div>"
							
						html = html & "</li>"
					end while
					html = html & "</ul>"
				html = html & "</div>"
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function listPaginasInstitucionais(byval notid as integer) as String
		Dim html, uri, categoria, slug as string
		dim tableViewCategoria as String = "_cms_view_cidades_paginas_categoria"
		dim id, idCategoria as Integer
		Dim data As DateTime
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select * from _cms_view_cidades_paginas where idCidade = @idCidade and parent in(0,2,3) and idcategoria not in(21,30) and id <> @id order by acessos desc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade			
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = notid
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<div id=""paginas-relacionadas"">" & vbnewline
					html = html & "<ul class=""slides_container"">" & vbnewline
					while dr.Read()
						data = dr("dataPublicacao").toString()
						id = dr("id")
						idCategoria = dr("idCategoria")
						categoria = getContentCategoria(tableViewCategoria,id,"nome")
						uri = _lib.baseURL("institucional") & dr("seoSlug") & ".html"
						
						html = html & "<li class=""list-container"">"
							html = html & "<div class=""gnc-header"">"
								html = html & "<figure>"
									html = html & _lib.fotoLegenda( dr("thumb").toString(), dr("titulo").toString(), uri, 175, 115, categoria, "") & vbnewline
								html = html & "</figure>"							
							html = html & "</div>"
							
							html = html & "<div class=""gnc-containner"">"
								html = html & "<h3><a href=""" & uri & """>" & _lib.shorten(dr("resumo").toString(),100,"...") & "</a></h3>"
							html = html & "</div>"
							
						html = html & "</li>"
					end while
					html = html & "</ul>"
				html = html & "</div>"
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function passeiosHome(qtd as integer) as String
		Dim contar as Integer = 1
		dim html as string = string.empty
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()		
			cmd = New SqlCommand("exec proc_passeios_home_cidades @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				While dr.Read()
					url = "/" & dr("slugCat").toString() & "/" & dr("slugPageCat").toString() & "/" & dr("slugPage").toString() & ".html"
					html = html & "<div class=""n-lista"">"
						html = html & "<p class=""n-Titulo""><a rel=""bookmark"" title="""&dr("titulo").toString()&""" href=""" & url & """>" & dr("titulo").toString() & "</a></p>"
						html = html & "<div class=""n-div"">"
							html = html & _lib.npThumb( dr("foto_representa").toString(), dr("titulo").toString(), url, 100, 73  )
							html = html & "<p class=""n-Texto"">" & cortaTexto(dr("resumo").toString(),95) & "</p>"
						html = html & "</div>"
					html = html & "</div>"
					if contar < qtd then html = html & "<div class=""linha""></div>"
					contar += 1
				End While
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function fotosHome(qtd as integer) as String
		Dim contar as Integer = 0
		dim html as string = string.empty
		Dim txt as string
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()		
			cmd = New SqlCommand("exec proc_fotos_home_netportal @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul id=""n-gfotos"" class=""carouselFotos"">"
				While dr.Read()
					txt = rtrim(cortaTexto(dr("txtgaleria").toString(),100) )
					url = "/" & dr("slugCatModulo").toString() & "/" & dr("slugCat").toString() & "/" & dr("slugPage").toString() & "/"
					html = html & noPadding(contar, qtd)
						html = html & _lib.npThumb( dr("foto").toString(), dr("titulo").toString(), url, 180, 130)
						html = html & "<p class=""n-Data"">" & databr(dr("datagaleria").toString()) & "</p>"
						html = html & "<p class=""n-Texto"">" & txt & "</p>"
					html = html & "</li>"
					contar += 1
				End While
				html = html & "</ul>"
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function thumbFotosTerra(qtd as integer) as String
		Dim contar as Integer = 1
		dim html as string = string.empty
		Dim txt as string
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()		
			cmd = New SqlCommand("exec proc_fotos_home_netportal @idCidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul class=""gnc-thumb-fotos"">"
				While dr.Read()
					url = "/" & dr("slugCatModulo").toString() & "/" & dr("slugCat").toString() & "/" & dr("slugPage").toString() & "/"
					html = html & "<li class=""itm-thumb" & contar & """>"
						html = html & "<a href=""" & url & """ title=""" & dr("titulo").toString() & """ class=""lnk-thumb"">"
							html = html & _lib.npThumbTerra(contar, dr("fotogrd").toString(), dr("titulo").toString())
						html = html & "</a>"
						
						html = html & "<a href=""" & url & """ class=""ctn-zoom"" style=""display:none;"">"
							html = html & _lib.npThumbImage(dr("fotogrd").toString(), dr("titulo").toString(), 250, 166)
							html = html & "GALERIA - " & databr(dr("datagaleria")) & "<br />"
							html = html & "<strong><span class=""photo"">Galeria de Fotos:</span>" & dr("titulo").toString() & "</strong>"
						html = html & "</a>"
						
					html = html & "</li>"
					
					contar += 1
				End While
				html = html & "</ul>"
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function getThumbGaleria(id as integer) as String
		dim retorno as string = string.empty
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			conn.Open()		
			cmd = New SqlCommand("select top 1 * from _cms_view_cidades_galeria_fotos_fotos where idgaleria = @id", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr("foto").toString()
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose() 
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return retorno
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function listFotos(byVal qtd as integer, byval notid as integer) as String
		Dim html, uri, thumb, folder, id as string
		Dim data As DateTime
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top " & qtd & " * from _cms_view_cidades_galeria_fotos where idCidade = @idCidade and id <> (@id) order by id desc;", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade			
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = notid
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<div id=""galerias-relacionadas"">" & vbnewline
					html = html & "<ul class=""slides_container"">" & vbnewline
					while dr.Read()
						id  = dr("id").toString()
						data = dr("dataPublicacao").toString()
						uri = "/fotos-campos-do-jordao/" & year(data) & "/" & Right("0" & month(data),2) & "/" & dr("seoSlug").toString() & ".html"
						folder = dr("folder").toString()
						thumb = folder & "/pq/" & getThumbGaleria(id)
						
						html = html & "<li class=""list-container"">"
							html = html & "<div class=""gnc-header"">"
								html = html & "<figure>"
									html = html & _lib.fotoLegenda( thumb, dr("titulo").toString(), uri, 175, 115, _lib.databr(data), "") & vbnewline
								html = html & "</figure>"							
							html = html & "</div>"
							
							html = html & "<div class=""gnc-containner"">"
								html = html & "<a href=""""><h3>" & dr("titulo").toString() & "</h3></a>"
							html = html & "</div>"
							
						html = html & "</li>"
					end while
					html = html & "</ul>"
				html = html & "</div>"
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function fotosGaleria(byVal id as integer, byval folder as String) as String
		Dim html, uri, thumb, foto, titulo, idFoto as string
		Dim data As DateTime
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select * from _cms_view_cidades_galeria_fotos_fotos where idGaleria = @id", conn)
			with cmd		
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul class=""fotos-galeria"">" & vbnewline
				while dr.Read()
					foto = folder & "/grd/" & dr("foto").toString()
					foto = _lib.getImageSite(foto)
					thumb =  folder & "/grd/" & dr("foto").toString()
					thumb = _lib.thumbSite(thumb,220,150)
					titulo = dr("titulo").toString()
					idFoto = dr("id").toString()
					
					html = html & "<li itemscope itemtype=""http://schema.org/ImageObject"" class=""foto-container"">"
						html = html & "<a rel=""lightbox[galeria]"" href=""" & foto & """ title=""" & titulo & """ >"
							html = html & "<img width=""220"" height=""150"" src=""" & thumb & """ itemprop=""contentURL"" title=""" & titulo & """ alt=""" & titulo & """ />"
						html = html & "</a>"
						html = html & "<span itemprop=""name"">" & titulo & " - Foto " & idFoto & "</span>"
					html = html & "</li>"
					
				end while
				html = html & "</ul>"
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function agendaHome(qtd as integer) as String
		Dim contar as Integer = 0
		dim html as string = string.empty
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()		
			cmd = New SqlCommand("exec proc_agenda_home_netportal @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				While dr.Read()
					html = html & "<div>"
						html = html & "<div class=""n-item"">"
							html = html & "<p class=""n-Data""><span>" & formatdatetime(dr("data_inicial"),1) & "</span></p>"
							html = html & "<p class=""n-Texto""><strong>" & dr("titulo") & "</strong></p>"
							html = html & "<p class=""n-Texto"">" & cortaTexto(dr("resumo"),150) & "</p>"
							if contar < qtd then html = html & "<div class=""linha""></div>"				
						html = html & "</div>"
					html = html & "</div>"
					contar += 1
				End While
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public Function widGetAgenda(byval qtd as integer) as string
		Dim contar as Integer = 1
		dim html as string = string.empty
		Dim strSQL, titulo, uri as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		
		html = html & "<div class=""n-colTitle""><h2>Agenda Campos do Jordão</h2></div>"
		html = html & "<div class=""n-content"">"
			html = html & "<div id=""agenda"">"
				
				html = html & "<div class=""n-list"">"
				try	
					If conn.State <> ConnectionState.Open Then conn.Open()		
					cmd = New SqlCommand("exec proc_agenda_home_netportal @id_cidade, @qtd", conn)
					with cmd
						.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
						.Parameters("@id_cidade").Value = idCidade
						.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
						.Parameters("@qtd").Value = qtd
					end with
					dr = cmd.ExecuteReader()
					if dr.HasRows then
						While dr.Read()
							html = html & "<div class=""inner"">"
								html = html & "<div class=""n-item"">"
									html = html & "<p class=""n-Data""><span>" & formatdatetime(dr("data_inicial"),1) & "</span></p>"
									
									html = html & "<figure>"
										html = html & _lib.fotoLegenda( dr("imagem_representa").toString(), dr("titulo").toString(), "/agenda/", 115, 85, databr(dr("data_inicial").toString()), "") & vbnewline
									html = html & "</figure>"
																		
									html = html & "<p class=""n-Texto""><strong>" & dr("titulo") & "</strong></p>"
									html = html & "<p class=""n-Texto"">" & cortaTexto(dr("resumo"),150) & "</p>"
									if contar < qtd then html = html & "<div class=""linha""></div>"				
								html = html & "</div>"
							html = html & "</div>"
							contar += 1
						End While
					end if
					dr.close()
					cmd.dispose()
				Catch ex As Exception
					html = ""
				Finally
					If conn.State = ConnectionState.Open Then conn.Close()
					If conn IsNot Nothing Then conn.Dispose()  
					If conn IsNot Nothing Then conn.Dispose()  
					If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
				End Try				
				html = html & "</div>"
				
			html = html & "</div>"
		html = html & "</div>"
		
		html = html & "<div class=""n-nav"">"
			html = html & "<div class=""n-link""><a href=""/agenda/"">ver agenda completa</a></div>"
			html = html & "<div class=""nav"">"
				html = html & "<span class=""n-prev""><a class=""prev disabled""></a></span>"
				html = html & "<span class=""n-next""><a class=""next""></a></span>"
			html = html & "</div>"
		html = html & "</div>"		
		
		return html
	end Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function institucionais(qtd as integer) as String
		Dim contar as Integer = 0
		dim html as string = string.empty
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()		
			cmd = New SqlCommand("exec procInstitucionaisHome @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul>"
				While dr.Read()
					html = html & noPadding(contar,qtd)
						html = html & _lib.npThumb( dr("img_representa").toString(), dr("titulopg").toString(), ( dr("urlPage").toString() ), 134, 98  )
						html = html & "<p class=""n-Data"">" & dr("nome") & "</p>"
						html = html & "<p class=""n-Texto"">" & cortaTexto(dr("resumo_destaque"),90 ) & "</p>"
					html = html & "</li>"
					contar += 1
				End While
				html = html & "</ul>"
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function widgetPaginas() as String
		dim html as string = string.empty
		Dim strSQL, url, urlBase as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()		
			cmd = New SqlCommand("select * from _cms_view_cidades_paginas where idCidade = @idCidade and pageMaster = 0 and parent in(2,3) order by acessos desc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				urlBase = _lib.baseURL("institucional")
				While dr.Read()
					
					url = urlBase & dr("seoSlug").toString & ".html"
					
					html = html & "<div class=""article"">"
						html = html & "<div class=""header"">"
							html = html & "<h3 class=""gncSubTitle""><a href=""" & url & """>" & dr("seoTitle") & "</a></h3>"
						html = html & "</div>"
						
						html = html & "<div class=""content"">"
						
							html = html & "<div class=""figure"">"
								html = html & _lib.fotoLegenda( dr("thumb").toString(), dr("titulo").toString(), url, 220, 150, "", "") & vbnewline
							html = html & "</div>"
							
							html = html & "<div class=""inner"">"
								html = html & "<p>" & dr("resumo") & "</p>"
							html = html & "</div>"
							
						html = html & "</div>"
						
						html = html & "<div class=""footer"">"
							html = html & "<div class=""gnc-left"">"
								html = html & "<a class=""btn-read-more"" rel=""nofollow"" href=""" & url & """>continue lendo</a>"
							html = html & "</div>"
						html = html & "</div>"
						
					html = html & "</div>"
				End While
			end if
			dr.close()
			cmd.dispose()
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
	End Function		
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function muralHome(qtd as integer) as String
		Dim html as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())

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
					html = html & "" & clearStringComments(cortaTexto(dr(2).toString(), 130)) & "<br />"					
					html = html & "</div></li>"
				end while
				html = html & "</ul>"
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		
		return html
		
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function metaHeader(idmodulo as integer, coluna as string) as string
		Dim strSQL, html as String
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())			
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
				html = dr(coluna).toString()
			end if						
		catch ex as exception
			html = ""
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try	
		
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function pageInfo(pageType as string,pageName as string,coluna as string) as string
		dim html as string
		if pageLevel = 3 then
			select case pageType
				case "institucionais"
				html = dataPageInstitucionais(pageName,coluna)
				case "bairros"
				html = dataPageBairros(pageName,coluna)
			end select
		end if
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function dataPageInstitucionais(byVal pageName as String, byval coluna as string) as String
		dim html as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select top 1 * from nv_page_institucionais where idCidade = @idCidade and name = @slug ", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idcidade
				.Parameters.Add(New SqlParameter("@slug", SqlDbType.VarChar))
				.Parameters("@slug").Value = pageName				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				html = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function infoPageInstitucionaisByID(byVal id as integer, byval coluna as string) as String
		dim html as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select top 1 * from nv_page_institucionais where id = @id ", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				html = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public Function PageInstitucionalRelation(ByVal idCat as Integer, ByVal pageModule as Integer)
		Dim html, uri as string
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
			html = html & "<div class=""gnc-chield"">"
			try	
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand("exec proc_paginas_relacionadas_netportal @idCidade, @idCat, @pageModule", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
					.Parameters("@idCidade").Value = idcidade
					.Parameters.Add(New SqlParameter("@idCat", SqlDbType.Int))
					.Parameters("@idCat").Value = idCat
					.Parameters.Add(New SqlParameter("@pageModule", SqlDbType.Int))
					.Parameters("@pageModule").Value = pageModule										
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					html = html & "<section class=""chield"">"
					while dr.Read()
						uri = "/" & dr("slugCategoria").toString() & "/" & dr("slugPage").toString() & ".html"
						
						html = html & "<article>"
							html = html & "<h1><a href=""" & uri & """>" & dr("titulopg").toString() & "</a></h1>"
							html = html & "<div class=""inner-chield"">"
								html = html & "<figure>"
									html = html & _lib.fotoLegenda( dr("img_representa").toString(), dr("titulopg").toString(), uri, 180, 130, dr("nome").toString(), "") & vbnewline
								html = html & "</figure>"
								html = html & "<p>" & _lib.shorten(dr("resumo_destaque").toString(), 180, "...") & "</p>"
								html = html & "<a class=""small_button"" href=""" & uri & """><span>Continue Lendo</span></a>"
							html = html & "</div>"
						html = html & "</article>"
					end while
					html = html & "</section>"
				end if
				dr.close
			Catch ex As Exception
				html = ""
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try							
			html = html & "</div>"
		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Info Bairros Cidade
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function dataPageBairros(byVal pageName as String, byval coluna as string) as String
		dim html as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select top 1 * from nv_page_bairros where idCidade = @idCidade and name = @slug ", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idcidade
				.Parameters.Add(New SqlParameter("@slug", SqlDbType.VarChar))
				.Parameters("@slug").Value = pageName
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				html = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Info Bairros Cidade
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function infoPageBairroID(byVal id as integer, byval coluna as string) as String
		dim html as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select top 1 * from nv_page_bairros where id = @id", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				html = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function		
	
	'******************************************************************************************************************
	'* Feature News Home
	'******************************************************************************************************************	
	Public Function featureLastNews(qtd)
		Dim html, uri, classCSS as string
		Dim i as integer = 1
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		
		html = html & "<div class=""gnc-feature-news"">" & vbnewline
			try	
				If conn.State <> ConnectionState.Open Then conn.Open()
				cmd = New SqlCommand("exec np_noticias @idCidade, @limita, @limitb, @acao", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
					.Parameters("@idCidade").Value = idcidade
					.Parameters.Add(New SqlParameter("@limita", SqlDbType.Int))
					.Parameters("@limita").Value = 1
					.Parameters.Add(New SqlParameter("@limitb", SqlDbType.Int))
					.Parameters("@limitb").Value = qtd
					.Parameters.Add(New SqlParameter("@acao", SqlDbType.Int))
					.Parameters("@acao").Value = 4														
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					html = html & "<ul>"
					while dr.Read()
						classCSS = ""					
						if i = 1 then classCSS = "class=""first"""
						if i = qtd then classCSS = "class=""last"""
						uri = dr("urlConteudoCategoria").toString()
						html = html & "<li " & classCSS & ">"
							html = html & _lib.fotoLegenda( dr("imagem_destaque").toString(), dr("titulo").toString(), uri, 95, 65, "", "" )
							html = html & "<p class=""news-cat"">" & dr("Cat").toString() & " - " & databr(dr("data_publicacao")) & "</p>" & vbnewline
							html = html & "<p class=""news""><a href=""" & uri & """>" & dr("titulo").toString() & "</a></p>" & vbnewline							
						html = html & "</li>"
						i = i + 1
					end while
					html = html & "</ul>"
				end if
				dr.close
			Catch ex As Exception
				html = ""
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				If conn IsNot Nothing Then conn.Dispose()  
				If conn IsNot Nothing Then conn.Dispose()  
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try
		html = html & "</div>"
			
		return html
	End Function					
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function dataCidade(byVal coluna as String) as String
		dim html as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())	
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_cidades_netportal @id_cidade", conn)
			cmd.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
			cmd.Parameters("@id_cidade").Value = idcidade
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				html = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function viewCidade(byVal id as integer, byVal coluna as String) as String
		dim html as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())	
		try	
			conn.Open()
			cmd = New SqlCommand("exec proc_cidades_netportal @idCidade", conn)
			cmd.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
			cmd.Parameters("@idCidade").Value = id
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				html = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose()    
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return html
   	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public function getOAuth(byVal idServico as Integer, byval coluna as String) as String
		Dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())	
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from cidades_OAuth where idCidade = @idCidade and idServico = @idServico", conn)
			cmd.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
			cmd.Parameters("@idCidade").Value = idcidade
			cmd.Parameters.Add(New SqlParameter("@idServico", SqlDbType.Int))
			cmd.Parameters("@idServico").Value = idServico			
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return retorno
   	End Function	
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Dados Modulo
	'**********************************************************************************************************
	Public function dadosModulo(byVal idModulo as Integer, ByVal coluna as String) as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as String = String.Empty	
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from nv_modulos_cidade where idCidade = @idCidade and idModulo = @idModulo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
				retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()	
		End Try	

		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Dados Modulo
	'**********************************************************************************************************
	Public function getModulo(byVal idModulo as Integer, ByVal coluna as String) as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as String = String.Empty	
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _cms_view_cidades_modulos_config where id_Cidade = @idCidade and id_Modulo = @idModulo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
				retorno = ""
		Finally
			conn.Close()
			conn.Dispose()    
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try	

		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check login UserName and password
	'**********************************************************************************************************
	Public function checkLoginUserName(byVal userName as String, byVal userPasswd as String) as Boolean
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as boolean	
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 id from tbl_login_clientes_netportal where usuario = @usuario and senha = @senha", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@usuario", SqlDbType.VarChar))
				.Parameters("@usuario").Value = userName
				.Parameters.Add(New SqlParameter("@senha", SqlDbType.VarChar))
				.Parameters("@senha").Value = userPasswd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				retorno = true
			end if
			dr.close
		Catch ex As Exception
			retorno = false
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try

		return retorno
   	End Function
		
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function checkUserName(byVal userName as String) as Boolean
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as boolean	
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 id from tbl_login_clientes_netportal where usuario = @userName", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@userName", SqlDbType.VarChar))
				.Parameters("@userName").Value = userName
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				retorno = true
			end if
			dr.close
		Catch ex As Exception
				retorno = false
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()	
		End Try	

		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function getUser(byVal id as Integer, byval coluna as string) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as string	
		if _lib.lenb(id) > 0 and _lib.lenb(coluna) > 0 then
			try
				conn.Open()
				cmd = New SqlCommand("select top 1 * from _cms_view_cidades_usuarios where id = @id", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
					.Parameters("@id").Value = id
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					If not IsDBNull(dr(coluna).toString()) then retorno = dr(coluna).toString()
				end if
				dr.close
			Catch ex As Exception
					retorno = ""
			Finally
				conn.Close()
				conn.Dispose()    
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try	
		end if
		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function getUserByLogin(byVal login as String, byval coluna as string) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as string	
		if _lib.lenb(login) > 0 and _lib.lenb(coluna) > 0 then
			try
				conn.Open()
				cmd = New SqlCommand("select top 1 * from _cms_view_cidades_usuarios where usuario = @login", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@login", SqlDbType.VarChar))
					.Parameters("@login").Value = login
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					If not IsDBNull(dr(coluna).toString()) then retorno = dr(coluna).toString()
				end if
				dr.close
			Catch ex As Exception
					retorno = ""
			Finally
				conn.Close()
				conn.Dispose()    
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try	
		end if
		return retorno
   	End Function	
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function getUserMeta(byVal idUser as String, byval metaKey as string) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as string	
		if _lib.lenb(idUser) > 0 and _lib.lenb(metaKey) > 0 then
			try
				conn.Open()
				cmd = New SqlCommand("select top 1 * from _cms_view_cidades_usuarios_meta where idUser = @idUser and metaKey = @metaKey", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@idUser", SqlDbType.VarChar))
					.Parameters("@idUser").Value = idUser
					.Parameters.Add(New SqlParameter("@metaKey", SqlDbType.VarChar))
					.Parameters("@metaKey").Value = metaKey					
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					dr.Read()
					retorno = dr("metaValue").toString()
				end if
				dr.close
			Catch ex As Exception
					retorno = ""
			Finally
				conn.Close()
				conn.Dispose()    
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try	
		end if
		return retorno
   	End Function	
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function getUserID(byVal userName as String) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as string	
		try
			conn.Open()
			cmd = New SqlCommand("select top 1 userID from tbl_login_clientes_netportal where usuario = @userName", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@userName", SqlDbType.VarChar))
				.Parameters("@userName").Value = userName
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr("userid").toString()
			end if
			dr.close
		Catch ex As Exception
				retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()	
		End Try	

		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function getUserMD5(byVal userName as String) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as string	
		try
			conn.Open()
			cmd = New SqlCommand("select top 1 md5 from tbl_login_clientes_netportal where usuario = @userName", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@userName", SqlDbType.VarChar))
				.Parameters("@userName").Value = userName
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr("md5").toString()
			end if
			dr.close
		Catch ex As Exception
				retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()	
		End Try	

		return retorno
   	End Function	
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function getInfoUserMD5(byval md5 as string, byVal coluna as String) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as string	
		try
			conn.Open()
			cmd = New SqlCommand("select top 1 * from tbl_login_clientes_netportal where md5 = @md5", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@md5", SqlDbType.VarChar))
				.Parameters("@md5").Value = md5
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
				retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()	
		End Try	

		return retorno
   	End Function	
	
	'******************************************************************************************************************
	'* @SDESCRIPTION: Update Content Single Access
	'******************************************************************************************************************	
	Public sub updateUltimoAcessoUsuario(byVal userName as String)
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		if (_lib.lenb(userName) > 0) then
			try	
				conn.Open()
				cmd = New SqlCommand("update tbl_login_clientes_netportal set dataUltimoLogin = getDate() where usuario = @userName", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@userName", SqlDbType.VarChar))
					.Parameters("@userName").Value = userName
					.executeNonQuery()
				end with
			Catch ex As Exception
				
			Finally
				conn.Close()
				conn.Dispose()
			End Try
		end if
	End Sub
	
	'******************************************************************************************************************
	'* @SDESCRIPTION: Update Content Single Access
	'******************************************************************************************************************	
	Public sub updateUltimaAtividadeUsuario(byVal userName as String)
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		if (_lib.lenb(userName) > 0) then
			try	
				conn.Open()
				cmd = New SqlCommand("update tbl_login_clientes_netportal set dataUltimaAtividade = '' where usuario = @userName", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@userName", SqlDbType.VarChar))
					.Parameters("@userName").Value = userName
					.executeNonQuery()
				end with
			Catch ex As Exception
				
			Finally
				conn.Close()
				conn.Dispose()
			End Try
		end if
	End Sub	
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function checkUserName_OAuth(byVal idServico as integer, byVal userName as String, byVal userKey as string) as Boolean
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as boolean	
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 id from tbl_login_clientes_netportal_OAuth where idServico = @idServico and userName = @userName and idMembro = @userKey", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idServico", SqlDbType.Int))
				.Parameters("@idServico").Value = idServico			
				.Parameters.Add(New SqlParameter("@userName", SqlDbType.VarChar))
				.Parameters("@userName").Value = userName
				.Parameters.Add(New SqlParameter("@userKey", SqlDbType.VarChar))
				.Parameters("@userKey").Value = userKey				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				retorno = true
			end if
			dr.close
		Catch ex As Exception
				retorno = false
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()	
		End Try	

		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function checkUserMeta(byval metaKey as string, byval metaValue as string) as Integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		Dim retorno as Integer = 0	
		try	
			conn.Open()
			cmd = New SqlCommand("select count(*) from _cms_view_cidades_usuarios_meta where metaKey = @metaKey and metaValue = @metaValue", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@metaKey", SqlDbType.VarChar))
				.Parameters("@metaKey").Value = metaKey	
				.Parameters.Add(New SqlParameter("@metaValue", SqlDbType.VarChar))
				.Parameters("@metaValue").Value = metaValue						
			end with
			retorno = cmd.ExecuteScalar()
		Catch ex As Exception
			retorno = -1
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function checkMetaUserKey(byVal UserKey as string,byval metaKey as string) as boolean
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		Dim retorno as boolean = false	
		try	
			conn.Open()
			cmd = New SqlCommand("select count(*) from _cms_view_cidades_usuarios_meta where iduser = @userKey and metaKey = @metaKey", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@userKey", SqlDbType.VarChar))
				.Parameters("@userKey").Value = userKey				
				.Parameters.Add(New SqlParameter("@metaKey", SqlDbType.VarChar))
				.Parameters("@metaKey").Value = metaKey			
			end with
			if  cmd.ExecuteScalar() > 0 then retorno = true
		Catch ex As Exception
			retorno = false
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		return retorno
   	End Function	
		
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function checkUserNameFBLoginMeta(fbLogin as string) as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		Dim retorno as integer	
		try	
			conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_login_clientes_netportal_meta where metaKey = 'fbLogin' and metaValue = @fbLogin", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@fbLogin", SqlDbType.VarChar))
				.Parameters("@fbLogin").Value = fbLogin				
			end with
			retorno = cmd.ExecuteScalar()
		Catch ex As Exception
			retorno = -1
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function checkUserNameFBLoginMetaMD5(fbLogin as string, md5 as String) as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		Dim retorno as integer	
		try	
			conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_login_clientes_netportal_meta where iduser = @idUser and metaKey = 'fbLogin' and metaValue = @fbLogin", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idUser", SqlDbType.VarChar))
				.Parameters("@idUser").Value = md5			
				.Parameters.Add(New SqlParameter("@fbLogin", SqlDbType.VarChar))
				.Parameters("@fbLogin").Value = fbLogin
			end with
			retorno = cmd.ExecuteScalar()
		Catch ex As Exception
			retorno = -1
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function checkUserNameFBLoginMetaMD52(md5 as String) as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		Dim retorno as integer	
		try	
			conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_login_clientes_netportal_meta where iduser = @idUser and metaKey = 'fbLogin'", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idUser", SqlDbType.VarChar))
				.Parameters("@idUser").Value = md5
			end with
			retorno = cmd.ExecuteScalar()
		Catch ex As Exception
			retorno = -1
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function returnMD5FBUser(fblogin as String) as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as String	
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 idUser from tbl_login_clientes_netportal_meta where metaValue = @metaValue and metaKey = 'fbLogin'", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@metaValue", SqlDbType.VarChar))
				.Parameters("@metaValue").Value = fblogin
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr("idUser").toString()
			end if
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		return retorno
   	End Function	
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function addUserFbLogin(fbID as String, fblogin as string,username as string,nome as string,sobrenome as string,dataNasc as string,sexo as string,fbPicture as string, fbtoken as string) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		Dim retorno as string
		
		if _lib.lenb(sexo) > 0 then
			if sexo = "male" then 
				sexo = "masculino"
			else
				sexo = "feminino"
			end if
		end if
		
		dim tam : tam = 8 'Define o tamanho da senha		
		dim md : md = _lib.getMD5(DateTime.Now.toString())
		dim aleat : aleat =Int(22-tam*Rnd)+1
		dim senha : senha = mid(md, aleat, tam)			
		dim md5 as string = get_md5(DateTime.Now.toString())
		dim senhaMD5 as string = get_md5(senha)
		
		if _lib.lenb(dataNasc) > 0 then
			dataNasc = _lib.stringToDate(dataNasc)			
		end if
		
		try
			conn.Open()
			cmd = New SqlCommand("insert tbl_login_clientes_netportal (usuario,senha,md5,nome,sobrenome,data_nasc,sexo,fbfoto,cidade_mae) values (@usuario,@senha,@md5,@nome,@sobrenome,@data_nasc,@sexo,@fbfoto,@cidade_mae)", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@usuario", SqlDbType.VarChar))	
				.Parameters("@usuario").Value = fblogin
				.Parameters.Add(New SqlParameter("@senha", SqlDbType.VarChar))	
				.Parameters("@senha").Value = senhaMD5
				.Parameters.Add(New SqlParameter("@md5", SqlDbType.VarChar))	
				.Parameters("@md5").Value = md5
				.Parameters.Add(New SqlParameter("@nome", SqlDbType.VarChar))	
				.Parameters("@nome").Value = nome
				.Parameters.Add(New SqlParameter("@sobrenome", SqlDbType.VarChar))	
				.Parameters("@sobrenome").Value = sobrenome				
				.Parameters.Add(New SqlParameter("@data_nasc", SqlDbType.DateTime))	
				.Parameters("@data_nasc").Value = dataNasc
				.Parameters.Add(New SqlParameter("@sexo", SqlDbType.VarChar))	
				.Parameters("@sexo").Value = lcase(sexo)
				.Parameters.Add(New SqlParameter("@fbfoto", SqlDbType.VarChar))	
				.Parameters("@fbfoto").Value = lcase(fbPicture)				
				.Parameters.Add(New SqlParameter("@cidade_mae", SqlDbType.Int))	
				.Parameters("@cidade_mae").Value = idCidade
				.executeNonQuery()
			end with
				addUserMeta(md5,"fbID",fbID)
				addUserMeta(md5,"fbLogin",fbLogin)
				addUserMeta(md5,"fbUserName",username)
				addUserMeta(md5,"fbToken",fbtoken)
			retorno = 1
		Catch ex As Exception
			retorno = -1
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		return retorno
   	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public Sub addUserMeta(userMD5 as string, metaKey as string, metaValue as String)
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		
		try
			conn.Open()
			cmd = New SqlCommand("delete tbl_login_clientes_netportal_meta where idUser = @userMD5 and metaKey = @metaKey", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@userMD5", SqlDbType.VarChar))	
				.Parameters("@userMD5").Value = userMD5
				.Parameters.Add(New SqlParameter("@metaKey", SqlDbType.VarChar))	
				.Parameters("@metaKey").Value = metaKey
				.executeNonQuery()
			end with
			
			cmd = New SqlCommand("insert into tbl_login_clientes_netportal_meta (idUser,metaKey,metaValue) values (@idUser,@metaKey,@metaValue)", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idUser", SqlDbType.VarChar))	
				.Parameters("@idUser").Value = userMD5
				.Parameters.Add(New SqlParameter("@metaKey", SqlDbType.VarChar))	
				.Parameters("@metaKey").Value = metaKey
				.Parameters.Add(New SqlParameter("@metaValue", SqlDbType.VarChar))	
				.Parameters("@metaValue").Value = metaValue				
				.executeNonQuery()
			end with
			
		Catch ex As Exception
		
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		
   	End Sub	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getIDConteudo(byVal tableName as String, byval slug as string) as Integer
		Dim retorno as Integer
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 id from " & tableName & " where idCidade = @idcidade and seoSlug = @slug", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
				.Parameters("@idcidade").Value = idcidade
				.Parameters.Add(New SqlParameter("@slug", SqlDbType.VarChar))
				.Parameters("@slug").Value = slug							
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr("id")
			end if
			dr.close
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose()
If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return retorno
   	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getContent(byVal tableName as String, byval id as integer, byval coluna as string) as String
		Dim retorno as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from " & tableName & " where idCidade = @idcidade and id = @id", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
				.Parameters("@idcidade").Value = idcidade
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id							
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose()    
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return retorno
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getContentCategoria(byVal tableName as String, byval id as integer, byval coluna as string) as String
		Dim retorno as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from " & tableName & " where idCidade = @idcidade and id = @id", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
				.Parameters("@idcidade").Value = idcidade
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id							
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			retorno = ""
		Finally
			conn.Close()
			conn.Dispose()    
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		return retorno
   	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************	
	Public function totalComentarios(idModulo as integer, id as integer) as Integer
		Dim retorno as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select count(*) from view_tlb_comentarios_cidades where idcidade = @idCidade and idmodulo = @idModulo and idconteudo = @id", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id
			end with
			retorno = cmd.ExecuteScalar()
		Catch ex As Exception
			retorno = 0
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		
		return retorno
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************	
	Public function getCommentsContent(byval idconteudo as integer, byval idModulo as integer, byval inicio as integer, byval qtd as integer) as String
		Dim html as String
		if _lib.lenb(idModulo) > 0 and _lib.lenb(idConteudo) > 0 and _lib.lenb(inicio) > 0 and _lib.lenb(qtd) > 0 then
			Dim user as new npUsers()
			Dim nomeUser, fotoUser, userIdCity, userIdState, userCity, userState, localUsuario, userKey, postKey, comments as String
			Dim idUser, review, x, i as Integer
			Dim dataPublicacao as DateTime
			Dim conn as SqlConnection
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(sqlServer())
			try
				conn.Open()
				cmd = New SqlCommand("exec _cms_proc_comentarios_cidades @idCidade, @idModulo, @idConteudo, @inicio, @qtd", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
					.Parameters("@idCidade").Value = idCidade
					.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
					.Parameters("@idModulo").Value = idModulo
					.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
					.Parameters("@idConteudo").Value = idConteudo
					.Parameters.Add(New SqlParameter("@inicio", SqlDbType.Int))
					.Parameters("@inicio").Value = inicio
					.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
					.Parameters("@qtd").Value = qtd
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					x = inicio
					i = 0					
					While dr.Read()
						idUser = dr("idUser")
						nomeUser = user.nomeUsuario(idUser)
						nomeUser = _lib.Capitalize(nomeUser)
						userKey = dr("userKey").toString()
						postKey = dr("postKey").toString()
						fotoUser = user.UserPicture(idUser,userKey,"square")
						review = dr("nota")
						
						userIdCity = getUser(idUser, "idCidade")
						userIdState = getUser(idUser, "idEstado")
						userCity = getUser(idUser, "cidade")
						userState = getUser(idUser, "uf")
						localUsuario = ""
						
						if userIdCity > 0 and userIdState > 0 then
							localUsuario = viewCidade(userIdCity,"Cidade") & ", " & ucase(viewCidade(userIdCity,"uf"))
						elseif _lib.lenb(userCity) > 0 and _lib.lenb(userState) > 0  then
							localUsuario = userCity & ", " & ucase(userState)						
						end if
						
						dataPublicacao = dr("dataPublicacao")
						comments = dr("comentario")
						
						html = html & "<li id=""post-" & postKey & """>"
							html = html & "<div class=""comments-wrapper"">"
								
								html = html & "<div id=""comments-header"" class=""gnc-left"">"
									html = html & "<div class=""picture"">"
										html = html & fotoUser
									html = html & "</div>"								
								html = html & "</div>"
								
								html = html & "<div id=""comments-content"" class=""gnc-right"">"
									html = html & "<div class=""comments-author"">"
										html = html & "<strong>" & nomeUser & "</strong>"
										if _lib.lenb(localUsuario) > 0 then html = html & "<span class=""user-local gnc-hidden"">" & localUsuario & "</span>"
										html = html & "<span class=""comments-momment gnc-hidden""><abbr class=""timestamp"" title=""" & dataPublishLogDate(dataPublicacao.toString()) & """>" & dataTweet(dataPublicacao.toString()) & "</abbr></span>"
										if user.userON() then 
											if userKey = user.getCurrentUserKey() then html = html & "<span class=""gnc-delete gnc-hidden""><span title=""Excluir Comentário"" data-key=""" & postKey & """ class=""comments-del gnc_tip""></span></span>"
										end if
										html = html & "<span class=""gnc-improprio gnc-hidden""><span title=""Marcar como Impróprio"" class=""comments-improprio gnc_tip""></span></span>"
									html = html & "</div>"
									html = html & "<div class=""comments"">"
										html = html & "<p>" & comments & "</p>"
									html = html & "</div>"									
								html = html & "</div>"
								
								html = html & "<div class=""commets-footer"">"
									
									html = html & "<div class=""gnc-left"">"
										html = html & "<div class=""gnc-nota""><strong>Nota: </strong><span class=""comment-rating""><span class=""rating review_" & review & """></span></span></div>"
									html = html & "</div>"
									
									html = html & "<div class=""gnc-right"">"
										html = html & "<div class=""barra-list-comments-share""><label>Compartilhar</label><button type=""button"" title=""Compartilhar no Facebook"" class=""comments-share-bt-facebook gnc_tip"">Compartilhar no Facebook</button><button type=""button"" title=""Compartilhar no Twitter"" class=""comments-share-bt-twitter gnc_tip"">Compartilhar no Twitter</button><button type=""button"" title=""Compartilhar no Google+"" class=""comments-share-bt-googleplus gnc_tip"">Compartilhar no Google+</button></div>"
									html = html & "</div>"									
									
								html = html & "</div>"
								
							html = html & "</div>"
						html = html & "</li>"
						x += 1
						i += 1						
					End While
					if i = qtd and moreComments(idconteudo,idModulo,x,qtd) then
						html = html & "<div id=""more""><a  id=""" & x & """ class=""load_more"" data-content=""" & idconteudo & """ data-module=""" & idModulo & """ data-more=""" & qtd & """ href=""#"">mais resultados</a></div>"
					end if
				end if
				dr.Close()
			Catch ex As Exception
				html = "<div class=""no-comments"">Não foi possível carregar os comentários. Por favor nos informe o erro <a href=""#"" rel=""comments"" class=""gnc-erro-notify"">clicando aqui.</a></div>"
			Finally
				conn.Close()
				conn.Dispose()
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try	
		else
			html = "<div class=""no-comments"">Não foi possível carregar os comentários. Por favor nos informe o erro <a href=""#"" rel=""comments"" class=""gnc-erro-notify"">clicando aqui.</a></div>"
		end if
		return html
	end function
	
	'******************************************************************************************************************
	'* @SDESCRIPTION nota conteudo rating
	'******************************************************************************************************************		
	Public function moreComments(byval idconteudo as integer, byval idModulo as integer, byval inicio as integer, byval qtd as integer) as boolean
		dim retorno as boolean
		if _lib.lenb(idModulo) > 0 and _lib.lenb(idConteudo) > 0 and _lib.lenb(inicio) > 0 and _lib.lenb(qtd) > 0 then
			Dim conn as SqlConnection
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(sqlServer())
			try
				conn.Open()
				cmd = New SqlCommand("exec _cms_proc_comentarios_cidades @idCidade, @idModulo, @idConteudo, @inicio, @qtd", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
					.Parameters("@idCidade").Value = idCidade
					.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
					.Parameters("@idModulo").Value = idModulo
					.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
					.Parameters("@idConteudo").Value = idConteudo
					.Parameters.Add(New SqlParameter("@inicio", SqlDbType.Int))
					.Parameters("@inicio").Value = inicio
					.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
					.Parameters("@qtd").Value = qtd
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					retorno = true
				end if
				dr.Close()
			Catch ex As Exception
				retorno = false
			Finally
				conn.Close()
				conn.Dispose()
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try									
		end if
		return retorno
	end Function
	
	'******************************************************************************************************************
	'* @SDESCRIPTION nota conteudo rating
	'******************************************************************************************************************	
	Public Function insertComments(userKey as String, postkey as string, idModulo as integer, idConteudo as integer, comentario as string, acao as string, review as integer, scrapeTitle as String, scrapeDescription as String, scrapeThumb as String, scrapeURI as String, scrapeType as String ) as boolean	
		dim retorno as boolean
		if _lib.lenb(userKey) > 0 and _lib.lenb(idModulo) > 0 and _lib.lenb(idConteudo) > 0 and _lib.lenb(comentario) > 0 and _lib.lenb(postKey) > 0 and _lib.lenb(review) > 0 then
			Dim conn as SqlConnection
			Dim cmd as SqlCommand
			conn = New SqlConnection(sqlServer())
			try
				conn.Open()
				cmd = New SqlCommand("insert into tbl_comentarios_netportal (idinternauta, idcidade, idmodulo, idconteudo, comentario, acao, data, md5, scrapeTitle, scrapeDescription, scrapeThumb, scrapeURI, scrapeType, review) values (@userKey, @idcidade, @idmodulo, @idconteudo, @comentario, @acao, getdate(), @postkey, @scrapeTitle, @scrapeDescription, @scrapeThumb, @scrapeURI, @scrapeType, @review)", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@userKey", SqlDbType.VarChar))
					.Parameters("@userKey").Value = userKey
					.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
					.Parameters("@idcidade").Value = idCidade
					.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
					.Parameters("@idModulo").Value = idModulo
					.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
					.Parameters("@idConteudo").Value = idConteudo
					.Parameters.Add(New SqlParameter("@comentario", SqlDbType.VarChar))
					.Parameters("@comentario").Value = comentario
					.Parameters.Add(New SqlParameter("@acao", SqlDbType.VarChar))
					.Parameters("@acao").Value = acao
					.Parameters.Add(New SqlParameter("@postKey", SqlDbType.VarChar))
					.Parameters("@postKey").Value = postKey
					.Parameters.Add(New SqlParameter("@scrapeTitle", SqlDbType.VarChar))
					.Parameters("@scrapeTitle").Value = scrapeTitle
					.Parameters.Add(New SqlParameter("@scrapeDescription", SqlDbType.VarChar))
					.Parameters("@scrapeDescription").Value = scrapeDescription
					.Parameters.Add(New SqlParameter("@scrapeThumb", SqlDbType.VarChar))
					.Parameters("@scrapeThumb").Value = scrapeThumb
					.Parameters.Add(New SqlParameter("@scrapeURI", SqlDbType.VarChar))
					.Parameters("@scrapeURI").Value = scrapeURI
					.Parameters.Add(New SqlParameter("@scrapeType", SqlDbType.VarChar))
					.Parameters("@scrapeType").Value = scrapeType				
					.Parameters.Add(New SqlParameter("@review", SqlDbType.Int))
					.Parameters("@review").Value = review
					.ExecuteNonQuery()
				end with
				conn.close()
				
				if review > 0 then
					conn.Open()
					cmd = New SqlCommand( "insert into tbl_rating_netportal (id_internauta, id_comments, id_cidade, id_modulo, id_conteudo, nota, data ) values (@userid, @postid, @idcidade, @idmodulo, @idconteudo, @review, getdate() )", conn )
					with cmd
						.Parameters.Add(New SqlParameter("@userid", SqlDbType.VarChar))
						.Parameters("@userid").Value = userKey
						.Parameters.Add(New SqlParameter("@postid", SqlDbType.VarChar))
						.Parameters("@postid").Value = postKey			
						.Parameters.Add(New SqlParameter("@idcidade", SqlDbType.Int))
						.Parameters("@idcidade").Value = idCidade
						.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
						.Parameters("@idModulo").Value = idModulo
						.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
						.Parameters("@idConteudo").Value = idConteudo
						.Parameters.Add(New SqlParameter("@review", SqlDbType.Int))
						.Parameters("@review").Value = review
						.ExecuteNonQuery()
					end with
					conn.close()
				end if				
				
				retorno = true
			Catch ex As Exception
				retorno = false
			Finally
				If conn.State = ConnectionState.Open Then conn.Close()
				conn.Dispose()
			End Try
		end if
		return retorno
	End Function
	'******************************************************************************************************************
	'* @SDESCRIPTION nota conteudo rating
	'******************************************************************************************************************	
	Public Function notaConteudo(idConteudo as integer, idModulo as integer, Nota as integer) as String
		Dim retorno as string
		Dim ipInternauta as String = _lib.getMD5(_lib.ipInternauta())
		Dim verifica as integer = verificaNota(idConteudo, idModulo)
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		
		if verifica = 0 then
			try
				conn.Open()
				cmd = New SqlCommand("insert into _netportal_nota_conteudo (idModulo, idConteudo, nota, ipInternauta) values (@idModulo, @idConteudo, @Nota, @ipInternauta) ", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
					.Parameters("@idModulo").Value = idModulo
					.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
					.Parameters("@idConteudo").Value = idConteudo
					.Parameters.Add(New SqlParameter("@Nota", SqlDbType.Int))
					.Parameters("@Nota").Value = Nota
					.Parameters.Add(New SqlParameter("@ipInternauta", SqlDbType.VarChar))
					.Parameters("@ipInternauta").Value = ipInternauta
					.ExecuteNonQuery()
				end with
				retorno = "true"
			Catch ex As Exception
				retorno = "false"
			Finally
				conn.Close()
				conn.Dispose()
			End Try	
		else
			retorno = "duplicado"
		end if
		
		return retorno	
	End Function
	
	'******************************************************************************************************************
	'* @SDESCRIPTION
	'******************************************************************************************************************	
	Private Function verificaNota(idConteudo as integer, idModulo as integer) as String
		Dim retorno as Integer = 0
		Dim ipInternauta as String = _lib.getMD5(_lib.ipInternauta())
		Dim exec as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		
		try
			conn.Open()
			cmd = New SqlCommand("select count(*) from _netportal_nota_conteudo where idModulo = @idModulo and idConteudo = @idConteudo and ipInternauta = @ipInternauta", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
				.Parameters.Add(New SqlParameter("@ipInternauta", SqlDbType.VarChar))
				.Parameters("@ipInternauta").Value = ipInternauta
				retorno = .ExecuteScalar()
			end with
		Catch ex As Exception
			retorno = 0
		Finally
			conn.Close()
			conn.Dispose()
		End Try	
		
		return retorno	
	End Function
		
	'******************************************************************************************************************
	'* @SDESCRIPTION
	'******************************************************************************************************************	
	Public Function getNotaConteudo(idConteudo as integer, idModulo as integer, coluna as String) as String
		Dim retorno as Integer = 0
		Dim ipInternauta as String = _lib.getMD5(_lib.ipInternauta())
		Dim exec as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		Dim dr As SqlDataReader
		
		try
			conn.Open()
			cmd = New SqlCommand("select avg(nota) as media, count(id) as total from _netportal_nota_conteudo where idConteudo = @idConteudo and idModulo = @idModulo", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
				.Parameters.Add(New SqlParameter("@idModulo", SqlDbType.Int))
				.Parameters("@idModulo").Value = idModulo				
			end with
			dr = cmd.ExecuteReader()
			dr.read()
			if dr.HasRows then
				retorno = dr(coluna).toString()
			end if
			dr.close
		Catch ex As Exception
			retorno = 0
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		
		return retorno	
	End Function	
		
	'******************************************************************************************************************
	'* @SDESCRIPTION: Update Content Single Access
	'******************************************************************************************************************	
	Public sub updateAcessosConteudo(id as integer, table as string)
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		if (_lib.lenb(id) > 0) and (_lib.lenb(table) > 0) then
			try	
				conn.Open()
				cmd = New SqlCommand("update " & table & " set acessos = acessos + 1, dataUltimoAcesso = getdate() where id = @id", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
					.Parameters("@id").Value = id
					.executeNonQuery()
				end with
			Catch ex As Exception
				
			Finally
				conn.Close()
				conn.Dispose()
			End Try
		end if
	End Sub
			
End Class