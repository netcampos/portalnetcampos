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

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Schema

Imports MySql.Data
Imports MySql.Data.MySqlClient

Public Class _database

	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public Sub New()
		MyBase.New()
	End Sub
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Private Property sqlServer() As String
		Get
			Return ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property domainName() As String
		Get
			Return "http://www.netcampos.com"
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Private Property mysqlServer() As String
		Get
			Return ConfigurationManager.ConnectionStrings("mysqlServer").ConnectionString
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Private Property idCidade() As Integer
		Get
			Return 4787
		End Get
		Set(ByVal value As Integer)
		
		End Set
	End Property
	
'/#########################################################################################################################
'* init Functions Tabelas Auxiliares
'/#########################################################################################################################

	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function infoModulo(byVal id as integer, byval coluna as string) as String
		dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _cms_view_cidades_modulos where id = @id", conn)
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
	Public function infoConfigModulo(byVal idModulo as integer, byval coluna as string) as String
		dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _cms_view_cidades_modulos_config where id_cidade = @idCidade and id_modulo = @idModulo ", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
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



'/#########################################################################################################################
'* end Functions Tabelas Auxiliares
'/#########################################################################################################################	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getMeta(byVal tableName as String, byVal idConteudo as integer, byval metaKey as string) as String
		Dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from " & tableName & " where idConteudo = @idConteudo and metaKey = @metaKey ", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idConteudo", SqlDbType.Int))
				.Parameters("@idConteudo").Value = idConteudo
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
		return retorno
   	End Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function viewCidades(byval coluna as string) as String
		dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _cms_view_cidades where id = @id ", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = idCidade()
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
	Public Function countUsers() as Integer
		Dim iTotal as Integer = 0
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		
		try	
			conn.Open()
			cmd = New SqlCommand("exec np_InternautasSearchTotal @idCidade", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
			end with
			iTotal = cmd.ExecuteScalar()
			httpContext.Current.session("totalUsers") = iTotal
		Catch ex As Exception
			
		Finally
			conn.Close()
			conn.Dispose()
		End Try				
			
		return iTotal
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	public function viewAds(idArea as integer, idLocal as integer) as string
		Dim html as string
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_adserver_banners_cidades_netportal where idcidade = @idCidade and idarea = @idArea and idlocalentrega = @idLocal and ativo = 1", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@idArea", SqlDbType.Int))
				.Parameters("@idArea").Value = idArea
				.Parameters.Add(New SqlParameter("@idLocal", SqlDbType.Int))
				.Parameters("@idLocal").Value = idLocal
			end with
			dim t as integer = cmd.ExecuteScalar()
			if t = 0 then		
				html = viewBannerGoogle(idArea)
			else
				html = viewBannerCliente(idArea, idLocal)
			end if
		Catch ex As Exception
			html = ""
		Finally
			conn.Close()
			conn.Dispose() 
		End Try
		
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	private function viewBannerGoogle(idArea as Integer) as string
		Dim googleads as integer
		Dim html as String
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		
		try	
			conn.Open()
			cmd = New SqlCommand("select google_ads, code_google_ads, banner_padrao, txt_banner_padrao from tbl_adserver_areas_cidades_netportal where idcidade = @id_cidade and idarea = @id_area", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade()
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
			conn.Close()
			conn.Dispose()    
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	private function viewBannerCliente(idArea as Integer, idLocal as Integer) as string
		Dim contar as Integer = 0
		Dim googleads as integer
		Dim html as string
			
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())

		try	
			conn.Open()
			cmd = New SqlCommand("select md5, googleads from tbl_adserver_banners_cidades_netportal where idcidade = @id_cidade and idarea = @id_area and idlocalentrega = @id_local and ativo = 1 order by newid()", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade()
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
					html = viewBannerGoogle(idArea)
				else
					html = showAd(dr(0).toString())
				end if			
			end if
			dr.close
		Catch ex As Exception
			return ""
		Finally
			conn.Close()
			conn.Dispose()  
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
			conn.Open()
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
					html = _lib2.playerSWF( arquivo, largura, altura, "clickTAG=http://adserver.guiadoturista.net/?catads=2%26ads=" & id, id.toString)
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
			conn.Close()
			conn.Dispose()   
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try				
	
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	noticias em destaque home
	'******************************************************************************************************************
	Public Function featuredNewsHome(qtd as integer) as string
		Dim slugPage as String = "http://www.netcampos.com/noticias-campos-do-jordao"
		Dim html, strSQL, titulo, uri, notId as String
		notId = "null"
		
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader		
		
		try
			conn.Open()
			
			cmd = New SqlCommand("Select top 1 * from nv_noticias_cidades where idCidade = @idCidade and destaque = 1 order by id desc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
			end with

			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				notId = dr("id").toString()
				
				uri = slugPage & "/" & year(dr("data_publicacao")) & "/" &  right( "0" & month(dr("data_publicacao")) ,2)  & "/" & dr("slugConteudo").toString() & ".html"
				
				html = html & "<div class=""destaque thumbnail"">"
					html = html & _lib2.fotoNoHref( dr("picture").toString(), dr("titulo").toString(), 422, 280)
					html = html & "<div class=""caption"">"
						html = html & "<span class=""gncCaption"">" & dr("categoria").toString() & "</span>"
						html = html & "<h3 class=""gncTexto""><a title=""" & dr("titulo").toString() & """ href=""" & uri & """>" & dr("titulo").toString() & "</a></h3>"
					html = html & "</div>"
				html = html & "</div>"
			end if
			dr.close			
			
			cmd = New SqlCommand("exec np_noticias_cidades @idCidade, 1, @qtd, null, null, null, null, null, null, @notid", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
				.Parameters.Add(New SqlParameter("@notid", SqlDbType.Varchar))
				.Parameters("@notid").Value = notId				
			end with
			
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul class=""ultimas_noticias"">"
				While dr.Read()
					
					uri = slugPage & "/" & year(dr("data_publicacao")) & "/" &  right( "0" & month(dr("data_publicacao")) ,2)  & "/" & dr("slugConteudo").toString() & ".html"
					
					html = html & "<li>"
						html = html & "<a href=""" & uri & """ class=""gncFoto"">"
							
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & _lib2.fotoNoHref( dr("picture").toString(), dr("titulo").toString(), 160, 110 )
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
								html = html & "<span class=""gncCaption"">" & dr("categoria").toString() & "</span>"
								html = html & "<h3 class=""gncTexto"">" & dr("titulo").toString() & "</h3>"
							html = html & "</span>"
							
						html = html & "</a>"
					html = html & "</li>"					
					
				End While
				html = html & "</ul>"
			end if
			dr.close()			
			
		Catch ex As Exception
		
		Finally
			conn.Close()
			conn.Dispose()
		End Try		

		return html
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	noticias em destaque home
	'******************************************************************************************************************
	Public Function featuredTrip(qtd as integer) as string
		Dim contar as Integer = 0
		Dim slugPage as String = "http://www.netcampos.com/passeios-campos-do-jordao"
		Dim html as string = string.empty
		Dim strSQL, titulo, uri as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())		
		try	
			conn.Open()		
			cmd = New SqlCommand("exec proc_passeios_home_cidades @idCidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul>"
				While dr.Read()
					
					uri = slugPage & "/" & dr("slugPage").toString() & ".html"
					
					html = html & "<li>"
						html = html & "<a href=""" & uri & """ class=""gncFoto"">"
							
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & _lib2.fotoNoHref( dr("foto_representa").toString(), dr("titulo").toString(), 160, 110)
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
								html = html & "<h3>" & dr("titulo").toString() & "</h3>"
								html = html & "<p>" & _lib2.shorten(dr("resumo").toString(),80,"...") & "</p>"
							html = html & "</span>"
							
						html = html & "</a>"
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

'/#########################################################################################################################
'* init Functions Videos
'/#########################################################################################################################
	
	'******************************************************************************************************************
	'* Feature Video ID Home
	'******************************************************************************************************************	
	Public Function getIDfeaturedVideo() as integer
		Dim id as Integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 id from _cms_view_cidades_videos where idCidade = @idCidade and destaque = 1 and ativo = 1 order by dataPublicacao desc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				id = dr("id")
			end if
			dr.close
		Catch ex As Exception
			id = 0
		Finally
			conn.Close()
			conn.Dispose()    
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		return id
	End Function
	
	'******************************************************************************************************************
	'* Feature Video
	'******************************************************************************************************************	
	Public Function featuredVideo(id as Integer, optional byval imgWidth as integer = 240, optional byval imgHeight as integer = 160) as String

		Dim idModulo as Integer = 10
		Dim tableMeta as String = _lib2.getTableMeta(idModulo)		
		Dim slugModulo as string = _lib2.getSlugModule(idModulo)		
		
		Dim mediaNota, totalNota as integer
		Dim dataPublicacao as Date
		Dim html, uri, conteudo, titulo, slug, picture, thumb, seoSlug as string

		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _cms_view_cidades_videos where id = @id", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = id
			end with
			
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				
				titulo = dr("titulo").toString()
				picture = dr("thumb").toString()
				
				dataPublicacao = dr("dataPublicacao")
				seoSlug = dr("seoSlug").toString()
							
				uri = "/" & slugModulo & "/" & year(dataPublicacao) & "/" &  right( "0" & month(dataPublicacao) ,2)  & "/" & seoSlug & "-" & id & ".html"				
				
				
				html = html & "<a href=""" & uri & """ class=""gncFeatured gncFoto"">"
					html = html & "<span class=""borda-foto""><span></span>"
						html = html & "<div class=""buttonPlay""></div>"
						html = html & _lib2.fotoNoHref(picture, titulo, imgWidth, imgHeight)
					html = html & "</span>"
					
					html = html & "<span class=""conteudo"">"
						html = html & "<span class=""data""><strong>Enviado:</strong> " & _lib2.dataTweetShort(dr("dataPublicacao").toString()) & "</span>"
						html = html & "<h2 class=""titulo"">" & titulo & "</h2>"						
					html = html & "</span>"
					
				html = html & "</a>"
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
	'* Feature Videos Home
	'******************************************************************************************************************	
	Public Function featuredVideos(qtd as integer, notid as Integer, optional byval imgWidth as integer = 160, optional byval imgHeight as integer = 110) as String

		Dim idModulo as Integer = 10
		Dim tableMeta as String = _lib2.getTableMeta(idModulo)		
		Dim slugModulo as string = _lib2.getSlugModule(idModulo)		
		
		Dim id, mediaNota, totalNota as integer
		Dim dataPublicacao as Date
		Dim html, uri, conteudo, titulo, slug, picture, thumb, seoSlug as string

		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top " & qtd & " * from _cms_view_cidades_videos where idCidade = @idCidade and id <> @id and destaque = 1 and ativo = 1 order by dataPublicacao desc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()			
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = notid
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul>"
				while dr.Read()
				
					id = dr("id").toString()
					dataPublicacao = dr("dataPublicacao")
					seoSlug = dr("seoSlug").toString()
								
					uri = "/" & slugModulo & "/" & year(dataPublicacao) & "/" &  right( "0" & month(dataPublicacao) ,2)  & "/" & seoSlug & "-" & id & ".html"				
				
					html = html & "<li>"
					
						html = html & "<a href=""" & uri & """ class=""gncFoto"">"
							
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & "<div class=""buttonPlay""></div>"
								html = html & _lib2.fotoNoHref( dr("thumb").toString(), dr("titulo").toString(), imgWidth, imgHeight)
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
							
							html = html & "</span>"
							
						html = html & "</a>"					
					
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
	'* Feature Videos Home
	'******************************************************************************************************************	
	Public Function featuredVideos2(qtd as integer, notid as Integer, optional byval imgWidth as integer = 160, optional byval imgHeight as integer = 110) as String
	
		Dim idModulo as Integer = 10
		Dim tableMeta as String = _lib2.getTableMeta(idModulo)		
		Dim slugModulo as string = _lib2.getSlugModule(idModulo)		
		
		Dim id, mediaNota, totalNota as integer
		Dim dataPublicacao as Date
		Dim html, uri, conteudo, titulo, slug, picture, thumb, seoSlug as string

		
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top " & qtd & " * from _cms_view_cidades_videos_categoria where idCidade = @idCidade and id <> @id and ativo = 1 order by dataPublicacao desc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()			
				.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
				.Parameters("@id").Value = notid
			end with
			
			idList.Clear()
			idList.Add(notid)
			
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul>"
				while dr.Read()
					id = dr("id").toString()
					idList.Add(id)
		
					dataPublicacao = dr("dataPublicacao")
					seoSlug = dr("seoSlug").toString()
								
					uri = "/" & slugModulo & "/" & year(dataPublicacao) & "/" &  right( "0" & month(dataPublicacao) ,2)  & "/" & seoSlug & "-" & id & ".html"
					
					mediaNota = getNotaConteudo(id,idModulo,"media")
					
					html = html & "<li>"
					
						html = html & "<a href=""" & uri & """ class=""gncFoto"">"
							
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & "<div class=""buttonPlay""></div>"
								html = html & _lib2.fotoNoHref( dr("thumb").toString(), dr("titulo").toString(), imgWidth, imgHeight)
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
								html = html & "<span class=""categoria"">" & dr("categoria").toString() & "</span>"
								html = html & "<h2 class=""titulo"">" &  dr("titulo").toString() & "</h2>"
								
								html = html & "<ul class=""estrelas " & mediaNota & """>"
									html = html & "<li class=""est1 " & IIf(mediaNota = 1, "ativo", "")  & """>1 estrela</li>"
									html = html & "<li class=""est2 " & IIf(mediaNota = 2, "ativo", "") & """>2 estrelas</li>"
									html = html & "<li class=""est3 " & IIf(mediaNota = 3, "ativo", "") & """>3 estrelas</li>"
									html = html & "<li class=""est4 " & IIf(mediaNota = 4, "ativo", "") & """>4 estrelas</li>"
									html = html & "<li class=""est5 " & IIf(mediaNota = 5, "ativo", "") & """>5 estrelas</li>"
								html = html & "</ul>"
								
								html = html & "<span class=""exibicoes"">" & dr("acessos").toString() & " exibições</span>"
								
							html = html & "</span>"
							
						html = html & "</a>"					
					
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
	'* Feature Videos Home
	'******************************************************************************************************************	
	Public Function featuredVideosList(qtd as integer, optional byval imgWidth as integer = 160, optional byval imgHeight as integer = 110) as String

		Dim idModulo as Integer = 10
		Dim tableMeta as String = _lib2.getTableMeta(idModulo)		
		Dim slugModulo as string = _lib2.getSlugModule(idModulo)
		
		
		Dim html, uri, conteudo, titulo, slug, picture, thumb as string
		Dim id as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		dim exclusion as string  = _lib2.joinArrayList(idList,",")
		
		try	
			conn.Open()
			cmd = New SqlCommand("select top " & qtd & " * from _cms_view_cidades_videos_categoria where idCidade = @idCidade and id not in (" & exclusion & ") and ativo = 1 order by dataPublicacao desc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
			end with
			
			idList.Clear()
			
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul class=""thumbnails"">"
				while dr.Read()
					html = html & "<li class=""span2"">"
						html = html & "<div class=""thumbnail"">"
							html = html & "<div class=""img"">"
								html = html & "<div class=""buttonPlay""></div>"
								html = html & _lib2.fotoNoHref( dr("thumb").toString(), dr("titulo").toString(), imgWidth, imgHeight)
							html = html & "</div>"
							
							html = html & "<div class=""caption"">"
								html = html & "<h3>Thumbnail label</h3>"
								
							html = html & "</div>"
						html = html & "</div>"
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
	
	'**********************************************************************************************************************
	'* List Videos Home
	'**********************************************************************************************************************
	Public Function getListVideos(totalRegistros as integer, maxLinkNav as integer, optional byval imgWidth as integer = 160, optional byval imgHeight as integer = 110) as String
		Dim idModulo as Integer = 10
		Dim tableMeta as String = _lib2.getTableMeta(idModulo)		
		Dim slugModulo as string = _lib2.getSlugModule(idModulo)		
		
		Dim html, titulo, resumo, categoria, thumb, seoSlug, uri as string
		Dim id, i, acessos, mediaNota, totalNota  as integer
		Dim dataPublicacao as Date
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		Dim dr As SqlDataReader		
		
		dim exclusion as string  = _lib2.joinArrayList(idList,",")
		
		Dim pAtual as String = _lib2.QS("page")
		If pAtual = "" Or pAtual <= 0 Then pAtual = 1
		Dim iTotalReg as Integer = totalSearchVideos(exclusion,0,"")
		Dim iTotalPaginas as Integer = 0
		
		If Cint(iTotalReg) > CInt(totalRegistros) Then
		   iTotalPaginas = Math.Ceiling(iTotalReg/totalRegistros)
		Else
		   iTotalPaginas = 1
		End If		
		
		if pAtual > iTotalPaginas then pAtual = iTotalPaginas
		
		if iTotalReg > 0 then
			try	
				conn.Open()
				cmd = New SqlCommand("exec _cms_proc_videos_cidades @idCidade, 1, @pAtual, @totalRegistros, @destaque, @categoria, @notIds", conn)
				with cmd
					.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
					.Parameters("@idCidade").Value = idCidade()
					.Parameters.Add(New SqlParameter("@pAtual", SqlDbType.Int))
					.Parameters("@pAtual").Value = pAtual
					.Parameters.Add(New SqlParameter("@totalRegistros", SqlDbType.Int))
					.Parameters("@totalRegistros").Value = totalRegistros
					.Parameters.Add(New SqlParameter("@destaque", SqlDbType.Int))
					.Parameters("@destaque").Value = DBNull.Value
					.Parameters.Add(New SqlParameter("@categoria", SqlDbType.VarChar))
					.Parameters("@categoria").Value = DBNull.Value
					.Parameters.Add(New SqlParameter("@notIds", SqlDbType.VarChar))
					.Parameters("@notIds").Value = IIf(_lib2.lenb(exclusion) > 0, exclusion, DBNull.Value)
				end with
				dr = cmd.ExecuteReader()
				if dr.HasRows then
					html = html & "<div class=""inner"">" & vbnewline
						html = html & "<ul class=""thumbnails"">" & vbnewline
					while dr.Read()
						id = dr("id")
						titulo = dr("titulo").toString()
						categoria = dr("categoria")						
						thumb = dr("thumb").toString()
						resumo = dr("resumo").toString()
						dataPublicacao = dr("dataPublicacao")
						acessos = dr("acessos")
						seoSlug = dr("seoSlug").toString()
						
						uri = "/" & slugModulo & "/" & year(dataPublicacao) & "/" &  right( "0" & month(dataPublicacao) ,2)  & "/" & seoSlug & "-" & id & ".html"
						
						
						mediaNota = getNotaConteudo(id,idModulo,"media")
						
						html = html & "<li class=""span2"">"
							html = html & "<div class=""thumbnail"">"
								html = html & "<div class=""img"">"
									html = html & "<div class=""buttonPlay""></div>"
									html = html & "<a href=""" & uri & """>"
										html = html & _lib2.fotoNoHref( thumb, titulo, imgWidth, imgHeight)
									html = html & "</a>"
								html = html & "</div>"
								
								html = html & "<div class=""caption"">"
									html = html & "<span class=""categoria"">" & categoria & "</span>"
									html = html & "<span class=""titulo"">"
									html = html & "<a href=""" & uri & """>" & titulo
									html = html & "</a></span>"
									
									html = html & "<span class=""exibicoes""><strong>" & acessos & "</strong> exibições</span>"
									html = html & "<span class=""estrela"">"
										html = html & "<ul class=""estrelas"">"
											html = html & "<li class=""est1 " & IIf(mediaNota = 1, "ativo", "")  & """>1 estrela</li>"
											html = html & "<li class=""est2 " & IIf(mediaNota = 2, "ativo", "") & """>2 estrelas</li>"
											html = html & "<li class=""est3 " & IIf(mediaNota = 3, "ativo", "") & """>3 estrelas</li>"
											html = html & "<li class=""est4 " & IIf(mediaNota = 4, "ativo", "") & """>4 estrelas</li>"
											html = html & "<li class=""est5 " & IIf(mediaNota = 5, "ativo", "") & """>5 estrelas</li>"
										html = html & "</ul>"																		
									html = html & "</span>"
								html = html & "</div>"
								
							html = html & "</div>"
							
							
						html = html & "</li>"						
						
					end while
						html = html & "</ul>"
					html = html & "</div>"
					
					if (iTotalPaginas > 1) then
						html = html & "<div class=""gnc-nav"">"
							html = html & _lib2.npNav(pAtual,iTotalPaginas,maxLinkNav,_lib2.getURL(0))
						html = html & "</div>"
					end if
					
				end if
				dr.close
			Catch ex As Exception
				html = ""
			Finally
				conn.Close()
				conn.Dispose()
				dr.close
			End Try
			
		end if
			
		return html
	End Function
	
	'******************************************************************************************************************
	'* @SDESCRIPTION
	'******************************************************************************************************************	
	Public Function totalSearchVideos(notIds as String, destaque as integer, slugcategoria as string) as Integer
		Dim retorno as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("exec _cms_proc_videos_cidades_total @idCidade, 1, @destaque, @categoria, @notIds", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@destaque", SqlDbType.Int))
				.Parameters("@destaque").Value = IIf(destaque > 0, destaque, DBNull.Value) 
				.Parameters.Add(New SqlParameter("@categoria", SqlDbType.VarChar))
				.Parameters("@categoria").Value = IIf(_lib2.lenb(slugcategoria) > 0, slugcategoria, DBNull.Value)				
				.Parameters.Add(New SqlParameter("@notIds", SqlDbType.VarChar))
				.Parameters("@notIds").Value = IIf(_lib2.lenb(notIds) > 0, notIds, DBNull.Value)
			end with
			retorno = cmd.ExecuteScalar()
		Catch ex As Exception
			retorno = 0
		Finally
			conn.Close()
			conn.Dispose()
		End Try				
		return retorno
	End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function getVideo(byVal id as integer, byval coluna as string) as String
		dim retorno as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _cms_view_cidades_videos where id = @id ", conn)
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
		
	
'#########################################################################################################################
'* end Functions Videos
'#########################################################################################################################
	
	'******************************************************************************************************************
	'* getLista Empresas
	'******************************************************************************************************************	
	Public Function getListaEmpresas(idCategoria as integer, qtd as Integer) as String
		Dim slugPage as String = "http://www.netcampos.com/empresas/"
		Dim html, uri, conteudo, classCSS as string
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		try	
			conn.Open()
			cmd = New SqlCommand("select top " & qtd & " * from view_empresas_anunciantes where idCidade = @idCidade and idCategoria = @idCategoria order by newid()", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@idCategoria", SqlDbType.Int))
				.Parameters("@idCategoria").Value = idCategoria				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul>" & vbnewline
				while dr.Read()
					
					uri = slugPage & dr("urlAmigavel").toString() & ".html"

					html = html & "<li" & classCSS & ">" & vbnewline
						
						html = html & "<a href=""" & uri & """ class=""gncFoto"">"
						
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & _lib2.fotoNoHref( dr("picture").toString(), dr("titulo").toString(), 160, 110)
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
								html = html & "<h4>" & dr("titulo").toString() & "</h4>"
								html = html & "<p>" & _lib2.shorten(dr("conteudo").toString(),100,"...") & "</p>"
							html = html & "</span>"							
												
						html = html & "</a>"
						
					html = html & "</li>"
					
				end while
				html = html & "</ul>"& vbnewline				
				
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
	Public function widgetFotosHome(qtd as integer) as String
		Dim slugPage as String = "http://www.netcampos.com/fotos-campos-do-jordao"
		Dim contar as Integer = 0
		dim html, cssLinha as string
		Dim data As DateTime
		Dim txt as string
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			conn.Open()		
			cmd = New SqlCommand("exec proc_fotos_home_netportal @idCidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				
				While dr.Read()
					
					data = dr("dataGaleria").toString()
					url = slugPage & "/" & year(data) & "/" & Right("0" & month(data),2) & "/" & dr("slugPage").toString() & ".html"
					
					cssLinha = "l1 c" & contar + 1
					
					if contar / 4 >= 1 then cssLinha = "l2 c" & (contar-4) + 1
					
					html = html & "<div class=""" & cssLinha & " " & dr("slugCat").toString() & """>"
						
						html = html & "<a href=""" & url & """>"
						
							html = html & "<div class=""foto"">"
								html = html & _lib2.npThumbImage(dr("fotogrd").toString(), dr("titulo").toString(), 204, 145)
							html = html & "</div>"
							
							html = html & "<span class=""legenda"">"
								html = html & dr("Categoria").toString()
							html = html & "</span>"
							
						html = html & "</a>"
						
						html = html & "<a href=""" & url & """ class=""titulo"">"
							html = html & dr("titulo").toString()
						html = html & "</a>"
						
					html = html & "</div>"
					
					contar = contar + 1
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
	Public function widgetInstitucionaisHome(qtd as integer) as String
		Dim slugPage as String = "http://www.netcampos.com/sobre-campos-do-jordao/"
		dim html as string = string.empty
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			conn.Open()		
			cmd = New SqlCommand("exec procInstitucionaisHome @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul class=""thumbnails"">"
				While dr.Read()
				
					url = slugPage & dr("slugPage").toString() & ".html"
					
					html = html & "<li class=""thumbnail"">"
						
						html = html & _lib2.fotoNoHref( dr("img_representa").toString(), dr("titulopg").toString(), 165, 110)
						
						html = html & "<div class=""caption"">"
						
							html = html & "<h3>"
								html = html & "<a href=""" & url & """ title=""" & dr("titulopg").toString() & """>"
									html = html & dr("titulopg").toString()
								html = html & "</a>"
							html = html &  "</h3>"
							
							html = html & "<p>" & _lib2.shorten(dr("resumo_destaque").toString(),75,"...") & "</p>"
						html = html & "</div>"
					html = html & "</li>"
					
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
	Public function widgetInstitucionaisHome2(qtd as integer) as String
		Dim contar as Integer = 0
		dim html as string = string.empty
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			conn.Open()		
			cmd = New SqlCommand("exec procInstitucionaisHome @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul class=""gncHorizontal"">"
				While dr.Read()

					html = html & "<li>"

						html = html & "<a href=""#"" class=""gncFoto"">"
						
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & "<div class=""buttonPlay""></div>"
								html = html & _lib2.fotoNoHref( dr("img_representa").toString(), dr("titulopg").toString(), 160, 110)
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
								html = html & "<h3>" & dr("titulopg").toString() & "</h3>"
								html = html & "<p>" & _lib2.shorten(dr("resumo_destaque").toString(),60,"...") & "</p>"
							html = html & "</span>"							
												
						html = html & "</a>"

					html = html & "</li>"
					
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
	Public function widgetAgendaHome(qtd as integer) as String
		Dim contar as Integer = 0
		dim html as string = string.empty
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			conn.Open()		
			cmd = New SqlCommand("exec proc_agenda_home_netportal @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul class=""gncHorizontal"">"
				While dr.Read()

					html = html & "<li>"

						html = html & "<a href=""http://www.netcampos.com/agenda/"" class=""gncFoto"">"
						
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & _lib2.fotoNoHref( dr("imagem_representa").toString(), dr("titulo").toString(), 160, 110)
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
								html = html & "<p class=""gncCaption"">" & _lib2.databr(dr("data_inicial").toString()) & "</p>"
								html = html & "<h3 class=""gncTexto"">" & dr("titulo").toString() & "</h3>"
							html = html & "</span>"							
												
						html = html & "</a>"

					html = html & "</li>"
					
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
	Public function dropDownCatEmpresas() as String
		Dim slugPage as String = "http://www.netcampos.com/empresas/"
		Dim contar as Integer = 0
		dim html, urlMasterPg as string
		Dim strSQL, url as String
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn =  New SqlConnection(sqlServer())
		try	
			conn.Open()		
			cmd = New SqlCommand("SELECT distinct c.idcategoria, c.nomecategoria, c.url_amigavel FROM dbo.categorias as c inner join empresas as e on c.idcategoria = e.idcategoria where e.idcidade = @id_cidade ORDER BY c.nomecategoria ASC", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade()
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<select name=""catEmpresas"" class=""listcatempresas""><option value="""">O que você procura?</option>"
				While dr.Read()
					html = html & "<option value=""" & slugPage & dr(2).toString() & "/"">" & dr(1).toString() & "</option>"
				End While
				html = html & "</select>"
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
		
	
	'##################################### INIT USER QUERYS '##################################################
		
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************	
	Public function getUser(byVal id as Integer, byval coluna as string) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as string	
		if _lib2.lenb(id) > 0 and _lib2.lenb(coluna) > 0 then
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
		
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function getUserMeta(byVal idUser as String, byval metaKey as string) as string
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(sqlServer())
		Dim retorno as string	
		if _lib2.lenb(idUser) > 0 and _lib2.lenb(metaKey) > 0 then
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
			
	'##################################### END USER QUERYS '##################################################
	
	'******************************************************************************************************************
	'* @SDESCRIPTION: Nota Conteudo
	'******************************************************************************************************************	
	Public Function getNotaConteudo(idConteudo as integer, idModulo as integer, coluna as String) as String
		Dim retorno as Integer = 0
		Dim ipInternauta as String = _lib2.getMD5(_lib2.ipInternauta())
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
	
End Class







