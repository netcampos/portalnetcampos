Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports MySql.Data
Imports MySql.Data.MySqlClient

Imports System.IO
Imports System.Xml

Public Class _npDbViews

	Private _lib as new _npLibrary()
	
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
	Public Property domainName() As String
		Get
			Return System.Configuration.ConfigurationManager.AppSettings.Get("domainName")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property domainAssets() As String
		Get
			Return System.Configuration.ConfigurationManager.AppSettings.Get("domainAssets")
		End Get
		Set(ByVal value As String)
		
		End Set
	End Property	
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************
	Public Property domainImages() As String
		Get
			Return System.Configuration.ConfigurationManager.AppSettings("domainImages")
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
	'* init Views
	'/#########################################################################################################################
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Function getTableMeta(ByVal id As Integer) As String
		Dim retorno As String = ""
		if id > 0 then
			select case id
				case 1,16
					retorno = ""				
				case 3
					retorno = ""			
				case 5
					retorno = ""
				case 6
					retorno = ""				
				case 7
					retorno = ""
				case 10
					retorno = "_cms_view_cidades_videos_meta"
				case 27
					retorno = ""					
			end select
		End If
		Return retorno
	End Function
	
	'***********************************************************************************************************
	'* description: 
	'***********************************************************************************************************	
	Public Function getSlugModule(ByVal id As Integer) As String
		Dim retorno As String = ""
		if id > 0 then
			select case id
				case 3
					retorno = ""			
				case 5
					retorno = ""
				case 6
					retorno = ""				
				case 7
					retorno = ""
				case 10
					retorno = "videos-campos-do-jordao"
				case 27
					retorno = ""					
			end select
		End If
		Return retorno
	End Function		
		
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function infoEstado(byval id as integer, byval coluna as string) as String
		dim retorno as string = ""

		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _cms_view_estados where id = @id ", conn)
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
	Public function infoCidade(byval id as integer, byval coluna as string) as String
		dim retorno as string = ""

		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		
		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from _cms_view_cidades where id = @id ", conn)
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
	Public function getCidade(byval coluna as string) as String
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
	

	'**********************************************************************************************************************
	'' @SDESCRIPTION: get total members
	'**********************************************************************************************************************
	Public Function getTotalMembers() as Integer
		Dim retorno as Integer = 0
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(sqlServer())
		
		try	
			conn.Open()
			cmd = New SqlCommand("exec np_InternautasSearchTotal @idCidade", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value =  idCidade()
			end with
			retorno = cmd.ExecuteScalar()
			httpContext.Current.session("totalMembers") = retorno
		Catch ex As Exception
			retorno = 0
		Finally
			conn.Close()
			conn.Dispose()
		End Try				
			
		return retorno
	End Function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: Check UserName
	'**********************************************************************************************************
	Public function getInfoUserMD5(byval md5 as string, byVal coluna as String) as string
		Dim retorno as string	
		
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		
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
			retorno = "<div class=""alert alert-error""><button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & ex.message() & "</div>"
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
		
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		
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
				retorno = "<div class=""alert alert-error""><button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & ex.message() & "</div>"
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
		Dim retorno as string
		
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader		
		
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
				retorno = "<div class=""alert alert-error""><button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & ex.message() & "</div>"
			Finally
				conn.Close()
				conn.Dispose()    
				If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
			End Try	
		end if
		
		return retorno
   	End Function	

	'******************************************************************************************************************
	'* @SDESCRIPTION: Update Content Single Access
	'******************************************************************************************************************	
	Public sub userLogout(byVal userName as String)
		
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		
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
	
	'***********************************************************************************************************
	'* description 
	'***********************************************************************************************************
	Public Function getPageConfig(byval slug as string, byval coluna as string) as String
		Dim i as integer = 0
		Dim retorno as string
		Dim siteConfig as string = lcase("/App_Config/config.xml")
		Dim xml as XmlDocument = new XmlDocument()
		Dim m_nodelist As XmlNodeList
		Dim m_node As XmlNode
		
		If not System.IO.File.Exists( _lib.mapPath(siteConfig) ) Then 
			_lib.pageNotFound()
		else
			try
				xml.Load( _lib.mapPath(siteConfig))
				m_nodelist = xml.SelectNodes("/config/item")
				if m_nodelist.Count > 0 then
					For Each m_node In m_nodelist
						if lcase(m_node.ChildNodes.Item(0).InnerText) = lcase(slug) then
							For i = 0 To m_node.ChildNodes.Count - 1
								if lcase(m_node.ChildNodes(i).name) = lcase(coluna) then 
									retorno = m_node.ChildNodes(i).innerText
									exit for
								end if
							Next i
						end if
					next
				end if
			catch ex as exception
				retorno = ""
			end try				
		end if
		
		return retorno
	End Function	
		
		
	'**********************************************************************************************************************
	'' @SDESCRIPTION: get total members
	'**********************************************************************************************************************
	Public Function widgetFeaturedHome(byval qtd as integer, byval domClass as string) as String
		Dim html, thumb, target, itemClass as string
		Dim i as integer = 1
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader				
		
		try
			conn.Open()
			cmd = New SqlCommand( "select top 5 * from _cms_view_destaques_home where idCidade = @idCidade and ativo = 1 order by newid()", conn )
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value =  idCidade()
				dr = .ExecuteReader()
			end with
			if dr.HasRows then			
				html = html & "<div id=""featuredHome"" class=""" & domClass & """>"
					
					html = html & "<div class=""carousel-inner"">"
					
					While dr.Read()
						
						thumb = _lib.npThumb( dr("thumb") )
						target = ""
						itemClass = "class=""item"""
						
						if i = 1 then itemClass = "class=""item active"""
						if dr("target") <> "" then target = "target=""_blank"""
						
						html = html & "<div id=""featuredHome_item_" & i & """ " & itemClass & ">"
							html = html & "<a href=""" & dr("url") & """ " & target & " rel=""nofollow"">"
								html = html & "<img width=""530"" height=""350"" title=""" & dr("titulo") & """ alt=""" & dr("titulo") & """ src=""" & thumb & """>"
							html = html & "</a>"
							
							html = html & "<div class=""carousel-caption"">"
								html = html & "<span class=""gncCaption"">" & dr("caption") & "</span>"
								html = html & "<p>" & dr("titulo") & "</p>"
							html = html & "</div>"
						html = html & "</div>"
						
						i = i + 1
					end while
					
					html = html & "</div>"
					
					html = html & "<a class=""carousel-control left"" href=""#featuredHome"" data-slide=""prev"">&lsaquo;</a>"
					html = html & "<a class=""carousel-control right"" href=""#featuredHome"" data-slide=""next"">&rsaquo;</a>"
					
				html = html & "</div>"
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
	End Function		
	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************	
	Public function dropDownCatEmpresas() as String
		Dim slugPage as String = domainName & "/empresas/"
		Dim contar as Integer = 0
		dim html, urlMasterPg as string
		Dim strSQL, url as String

		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader	
		
		try	
			conn.Open()		
			cmd = New SqlCommand("SELECT distinct c.idcategoria, c.nomecategoria, c.url_amigavel FROM dbo.categorias as c inner join empresas as e on c.idcategoria = e.idcategoria where e.idcidade = @id_cidade ORDER BY c.nomecategoria ASC", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade()
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<select name=""catEmpresas"" class=""listcatempresas""><option value="""">Selecione uma Categoria</option>"
				While dr.Read()
					html = html & "<option value=""" & dr(0).toString() & """ data-url=""" & slugPage & dr(2).toString() & "/"">" & dr(1).toString() & "</option>"
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
		
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function sobreACidadeHome(qtd as integer) as String
		Dim slugPage as String = "http://www.netcampos.com/sobre-campos-do-jordao/"
		dim html as string = string.empty
		Dim strSQL, url as String
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		
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
					html = html & "<li class=""thumbnail"">"
						
						url = slugPage & dr("slugPage") & ".html"
						
						html = html & "<a href=""" & url & """ title=""" & dr("titulopg").toString() & """>"
							html = html & _lib.fotoNoHref( dr("img_representa").toString(), dr("titulopg").toString(), 165, 110)
						html = html & "</a>"
						
						html = html & "<div class=""caption"">"
						
							html = html & "<h3>"
								html = html & "<a href=""" & url & """ title=""" & dr("titulopg").toString() & """>"
									html = html & dr("titulopg").toString()
								html = html & "</a>"
							html = html &  "</h3>"						
						
							html = html & "<p>" & _lib.shorten(dr("resumo_destaque").toString(),75,"...") & "</p>"
							
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
	'' @SDESCRIPTION:	noticias em destaque home
	'******************************************************************************************************************
	Public Function featuredNewsHome(qtd as integer) as string
		Dim slugPage as String = domainName & "/noticias-campos-do-jordao"
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
					html = html & _lib.fotoNoHref( dr("picture").toString(), dr("titulo").toString(), 422, 280)
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
				html = html & "<ul id=""ultimas_noticias"" class=""groupListView"">"
				While dr.Read()
					
					uri = slugPage & "/" & year(dr("data_publicacao")) & "/" &  right( "0" & month(dr("data_publicacao")) ,2)  & "/" & dr("slugConteudo").toString() & ".html"
					
					html = html & "<li>"
						html = html & "<a href=""" & uri & """ class=""gncFotoLegenda"">"
							
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & _lib.fotoNoHref( dr("picture").toString(), dr("titulo").toString(), 160, 110 )
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
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try		

		return html
	End Function
			
	'******************************************************************************************************************
	'' @SDESCRIPTION:	noticias em destaque home
	'******************************************************************************************************************
	Public Function featuredTrip(qtd as integer) as string
		Dim contar as Integer = 0
		Dim slugPage as String = domainName & "/passeios-campos-do-jordao"
		Dim html as string = string.empty
		Dim strSQL, titulo, uri as String
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader			

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
				html = html & "<ul class=""groupListView"">"
				While dr.Read()
					
					uri = slugPage & "/" & dr("slugPage").toString() & ".html"
					
					html = html & "<li>"
						html = html & "<a href=""" & uri & """ class=""gncFotoLegenda"">"
							
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & _lib.fotoNoHref( dr("foto_representa").toString(), dr("titulo").toString(), 160, 110)
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
								html = html & "<h3 class=""gncCaption"">" & dr("titulo").toString() & "</h3>"
								html = html & "<span class=""gncTexto"">" & _lib.shorten(dr("resumo").toString(),120,"...") & "</span>"
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
	
	'******************************************************************************************************************
	'* Feature Video
	'******************************************************************************************************************	
	Public Function featuredVideo(id as Integer, optional byval imgWidth as integer = 225, optional byval imgHeight as integer = 140) as String

		Dim idModulo as Integer 	= 10
		Dim tableMeta as String 	= getTableMeta(idModulo)		
		Dim slugModulo as string 	= domainName & "/" & getSlugModule(idModulo)		
		
		Dim mediaNota, totalNota as integer
		Dim dataPublicacao as Date
		Dim html, uri, conteudo, titulo, slug, picture, thumb, seoSlug as string

		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		
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
				
				uri = slugModulo & "/" & year(dataPublicacao) & "/" &  right( "0" & month(dataPublicacao) ,2)  & "/" & seoSlug & "-" & id & ".html"	
				
				html = html & "<a href=""" & uri & """ class=""gncFeatured gncFotoLegenda"">"
				
					html = html & "<span class=""borda-foto""><span></span>"
						html = html & "<div class=""buttonPlay""></div>"
						html = html & _lib.fotoNoHref(picture, titulo, imgWidth, imgHeight)
					html = html & "</span>"
					
					html = html & "<span class=""conteudo"">"
						html = html & "<span class=""gncCaption""><strong>Enviado:</strong> " & _lib.dataTweetShort(dr("dataPublicacao").toString()) & "</span>"
						html = html & "<h2 class=""gncTexto"">" & titulo & "</h2>"
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
	'* Featured Videos Home
	'******************************************************************************************************************	
	Public Function featuredVideos(qtd as integer, notid as Integer, optional byval imgWidth as integer = 160, optional byval imgHeight as integer = 110) as String

		Dim idModulo as Integer 	= 10
		Dim tableMeta as String 	= getTableMeta(idModulo)		
		Dim slugModulo as string 	= domainName & "/" & getSlugModule(idModulo)	
		
		Dim id, mediaNota, totalNota as integer
		Dim dataPublicacao as Date
		Dim html, uri, conteudo, titulo, slug, picture, thumb, seoSlug as string

		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		
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
								
					uri = slugModulo & "/" & year(dataPublicacao) & "/" &  right( "0" & month(dataPublicacao) ,2)  & "/" & seoSlug & "-" & id & ".html"				
				
					html = html & "<li>"
					
						html = html & "<a href=""" & uri & """ class=""gncFotoLegenda"">"
							
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & "<div class=""buttonPlay""></div>"
								html = html & _lib.fotoNoHref( dr("thumb").toString(), dr("titulo").toString(), imgWidth, imgHeight)
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
	'* Feature Video ID Home
	'******************************************************************************************************************	
	Public Function getIDfeaturedVideo() as integer
		Dim id as Integer
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader	
		
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
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function agendaAcordeon(byval qtd as integer) as string
		Dim html, uri as String
		Dim imagem, mes as String

		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		
		try	
			conn.Open()
			cmd = New SqlCommand("exec np_agenda_cidades_listview_novo @id_cidade, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value = qtd				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				
				uri = "http://www.netcampos.com/agenda/"
				
				html = html & "<ul class=""gnc-accordion1"">"
					while dr.Read()
						mes = lcase(MonthName(dr("mes"), false))
						html = html & "<li>"
								html = html & "<div class=""handle1"">"
									html = html & "<div class=""mes " & mes & """>"
										html = html & "<p class=""dia"">" & day(dr("data_inicial")).toString() & "</p>"
									html = html & "</div>"
								html = html & "</div>"
								
								html = html & "<div class=""content"">"
									html = html & "<div class=""gnc-left"">"
										html = html & "<a href=""" & uri & """ class=""link-evento"">" & dr("titulo").toString() & "</a>"
										html = html & "<p class=""resumo"">" & dr("resumo").toString() & "</p>"
										html = html & "<p class=""texto""></p>"
									html = html & "</div>"
									
									html = html & "<div class=""gnc-right"">"
										html = html & "<div class=""figure"">"
											'html = html & _lib.npThumbImage(dr("fotogrd").toString(), dr("titulo").toString(), 204, 145)
											html = html & _lib.fotoLegenda( dr("thumb").toString(), dr("titulo").toString(), uri, 252, 184, "<strong>LOCAL:&nbsp;&nbsp;&nbsp;</strong><br/>" & dr("local").toString(), "", "" )
										html = html & "</div>"
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
	end function
		
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
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader		
		
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
								html = html & _lib.npThumbImage(dr("fotogrd").toString(), dr("titulo").toString(), 204, 145)
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
	'* getLista Empresas
	'******************************************************************************************************************	
	Public Function getListaEmpresas(byval idCategoria as integer, byval qtd as Integer, optional byval shorten as integer = 100, optional byval head as string = "h4") as String
		Dim slugPage as String = "http://www.netcampos.com/empresas/"
		Dim html, uri, conteudo, classCSS as string
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader	
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
				html = html & "<ul class=""groupListView"">" & vbnewline
				while dr.Read()
					
					uri = slugPage & dr("urlAmigavel").toString() & ".html"

					html = html & "<li" & classCSS & ">" & vbnewline
						
						html = html & "<a href=""" & uri & """ class=""gncFotoLegenda"">"
						
							html = html & "<span class=""borda-foto"">"
								html = html & "<span></span>"
								html = html & _lib.fotoNoHref( dr("picture").toString(), dr("titulo").toString(), 160, 110)
							html = html & "</span>"
							
							html = html & "<span class=""conteudo"">"
								html = html & "<" & head & " class=""gncCaption"">" & dr("titulo").toString() & "</" & head & ">"
								html = html & "<p class=""gncTexto"">" & _lib.shorten(dr("conteudo").toString(),shorten,"...") & "</p>"
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
	'* @SDESCRIPTION
	'******************************************************************************************************************	
	Public Function noticiasMaisLidas(qtd as Integer) as String
		Dim html, classCSS, uri, titulo, slug as string
		Dim id as integer
		Dim i as integer = 1
		
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader	
		
		try	
			conn.Open()
			cmd = New SqlCommand("select top " & qtd & " id, titulo, url_amigavel from nv_last_noticias where idcidade = @idcidade and year(data) = year(getdate())-1 and month(data) > month(getdate())-2 order by hits desc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ol>" & vbnewline
				while dr.Read()
					classCSS = ""
					
					id = dr("id")
					titulo = dr("titulo").toString()
					slug = dr("url_amigavel").toString()					
					uri = "/noticias-campos-do-jordao/" & slug & "-" & dr("id").toString() & ".html"
					
					if i = 1 then classCSS = "first"
					if i = qtd then classCSS = "last"
					
					classCSS = "class=""position-" & i & " " & classCSS & """"
									
					html = html & "<li " & classCSS & ">"
						html = html & "<a href=""" & uri & """ title=""" & titulo & """ >" & titulo & "</a>"
					html = html & "</li>"
					
					i += 1
				end while
				html = html & "</ol>"
			end if
			dr.close()			
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
	Public function getLastEnquete() as String
		Dim html, titulo as String
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader

		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 id, pergunta from enquetes_netportal where ativo = 1 and idcidade = @idCidade order by id desc;", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				html = html & "<h2>" & dr("pergunta").toString() & "</h2>"
				html = html & "<div class=""gncWidgetEnquete"">"
					html = html & perguntasEnquete(dr("id").toString())
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
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function perguntasEnquete(byval id as integer) as string
		Dim html as String = String.Empty
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader

		try	
			conn.Open()
			cmd = New SqlCommand("select id, resp from perguntas_enquetes where idenquete = @idEnquete order by resp asc;", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idEnquete", SqlDbType.Int))
				.Parameters("@idEnquete").Value = id
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = "<form name=""frmEnquete"" id=""frmEnquete"" method=""post""> <ul>"
				while dr.Read()
					html = html & " <li><p class=""respostas_enquete"">"
					html = html & " <input name=""enquete_resposta"" class=""enquete_resposta"" type=""radio"" value=""" & dr(0).toString & """ />"
					html = html & " <span>" & dr(1).toString & "</span></p>"
					html = html & " </li>"
				end while
				html = html & "</ul>"
				html = html & "<div class=""btns"">"
				html = html & "<button id=""btVotarEnquete"" class=""btn btn-info btn-small bt_votar sprite-portal"" type=""button""><span>votar</span></button>"
				html = html & "<button id=""btResultadoEnquete"" class=""btn btn-small bt_resultado sprite-portal"" type=""button""><span>resultado</span></button>"
				html = html & "</div> <input type=""hidden"" name=""idEnquete"" id=""idEnquete"" value=""" & id & """/></form>"
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
	
	'**********************************************************************************************************************
	'' @SDESCRIPTION: get total members
	'**********************************************************************************************************************
	Public Function widgetLastUsers(byval qtd as integer) as string
		Dim _user as new _npUsers()
		Dim html as string
		
		Dim conn as SqlConnection = New SqlConnection(sqlServer())
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		
		Dim idUser, nome, userUri, cidade, estado, userHomeTown as string
		Dim id, idCity, idEstado as integer
		
		try	
			conn.Open()
			cmd = New SqlCommand("exec _cms_proc_usuarios_cidades @idCidade, 1, @qtd", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value =  idCidade()
				.Parameters.Add(New SqlParameter("@qtd", SqlDbType.Int))
				.Parameters("@qtd").Value =  qtd				
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<div class=""gncFacePile"">"
					html = html & "<ul>"
						While dr.Read()
						
							id = dr("id")
							nome = dr("nome") & " " & dr("sobrenome")
							nome = _lib.trataNome(nome)
							
							idUser = dr("idUser")
							userUri = "/membros/" & id & "-" & GenerateSlug(tiraAscento(dr("nome")),255) & ".html"
							
							estado = IIf(IsDBNull(dr("uf")), "", dr("uf"))
							cidade = IIf(IsDBNull(dr("cidade")), "", dr("cidade"))
							
							idEstado = dr("idEstado")						
							idCity = dr("idCidade")
							
							if idEstado > 0 then estado = infoEstado(idEstado,"uf")	
							if idCity > 0 then cidade = infoCidade(idCity,"cidade")
							
							userHomeTown = cidade
							
							if _lib.lenb(userHomeTown) > 0 and _lib.lenb(estado) > 0 then userHomeTown = cidade & ", " & estado
													
							html = html & "<li>"
								html = html & "<a target=""_blank"" href=""" & userUri & """ title=""" & nome & """>"
									html = html & _user.userThumb(id,"square")
								html = html & "</a>"
							html = html & "</li>"
						end while
					html = html & "</ul>"
				html = html & "</div>"
			end if
			dr.close
		Catch ex As Exception
			html = "<div class=""alert alert-error""><button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & ex.message() & "</div>"
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try				
			
		return html
	End Function						
	
	'**********************************************************************************************************************
	'' @SDESCRIPTION:
	'**********************************************************************************************************************
	public function ads(idArea as integer, idLocal as integer) as string
		Dim total as integer 
		Dim html as string
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		
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
			total = cmd.ExecuteScalar()
			html = iif(total=0,exibeBannerGoogle(idArea),exibeBannerCliente(idArea, idLocal))
			
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
	private function exibeBannerGoogle(idArea as Integer) as string
		'Dim googleads as integer
		'Dim html as String
		
		'Dim conn as SqlConnection = New SqlConnection(sqlServer())
		'Dim cmd as SqlCommand
		'Dim dr As SqlDataReader	
				
		return "google"
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	private function exibeBannerCliente(idArea as Integer, idLocal as Integer) as string
		Dim contar as Integer = 0
		Dim googleads as integer
		Dim html as string
			
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select md5, googleads from tbl_adserver_banners_cidades_netportal where idcidade = @idCidade and idarea = @idArea and idlocalentrega = @idLocal and ativo = 1 order by newid()", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@idArea", SqlDbType.Int))
				.Parameters("@idArea").Value = idArea
				.Parameters.Add(New SqlParameter("@idLocal", SqlDbType.Int))
				.Parameters("@idLocal").Value = idLocal				
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
			html = ex.message.toString()
		Finally
			conn.Close()
			conn.Dispose()
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		
		return html
		
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'' @PARAM:
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
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		
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
					arquivo = domainImages & "/" & arquivo & extensao
					html = _lib.playerSWF( arquivo, largura, altura, "clickTAG=http://adserver.guiadoturista.net/?catads=2%26ads=" & id, id.toString)
				end if
				
				if dr(1).toString = "gif" then
					arquivo = replace(arquivo,extensao,"")
					arquivo = replace(arquivo,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
					arquivo = domainImages & "/" & arquivo & extensao
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
	
	'**********************************************************************************************************************
	'' @SDESCRIPTION: get total members
	'**********************************************************************************************************************
	public function adsGroup(idGroup as integer, idLocal as integer) as string

		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		
		Dim swf as string
		Dim extensao as string
		Dim arquivo as string
		Dim largura as string
		Dim altura as string
		Dim urlDestino as string
		Dim txtUrl as string
		Dim html, element, md5 as string		
		
		try	
			conn.Open()
			cmd = New SqlCommand("select ad1.id, ad1.urlDestino, ad1.txtURL, ad1.arquivo, ad1.tipomidia, ad2.largura, ad2.altura, ad1.md5 from tbl_adserver_banners_cidades_netportal as ad1 inner join tbl_adserver_areas_netportal as ad2 on ad2.id = ad1.idArea where ad1.idCidade = @idCidade and ad2.idGroup = @idGroup and ad1.idlocalentrega = @idLocal and ad1.ativo = 1 order by ad1.idArea asc", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()
				.Parameters.Add(New SqlParameter("@idGroup", SqlDbType.Int))
				.Parameters("@idGroup").Value = idGroup
				.Parameters.Add(New SqlParameter("@idLocal", SqlDbType.Int))
				.Parameters("@idLocal").Value = idLocal
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = "<div class=""content""><ul>"
				While dr.Read()
				
					arquivo = dr(3).toString
					extensao = "." & dr(4).toString
					urlDestino = dr(1).toString
					txtUrl = dr(2).toString
					
					largura = dr(5).toString
					altura = dr(6).toString	
					
					md5 = dr(7).toString
					
					if dr(4).toString = "swf" then
						arquivo = replace(arquivo,extensao,"")
						arquivo = replace(arquivo,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
						arquivo = domainImages & "/" & arquivo & extensao
						element = _lib.playerSWF( arquivo, largura, altura, "clickTAG=http://adserver.guiadoturista.net/?catads=2%26ads=" & md5, md5.toString)
					end if
					
					if dr(4).toString = "gif" then
						arquivo = replace(arquivo,extensao,"")
						arquivo = replace(arquivo,"/cidades/cms/netgallery/media/saopaulo/camposdojordao/","")
						arquivo = domainImages & "/" & arquivo & extensao
						element = "<a href=""http://adserver.guiadoturista.net/?catads=2&amp;ads=" & md5 & """ title=""" & txtUrl & """ target=""_blank""><img src=""" & arquivo & """ alt=""" & txtUrl & """ width=""" & largura & """ height=""" & altura & """ /></a>"	
					end if					
				
					html = html & "<li>"
						html = html & element
					html = html & "</li>"
					
				End While					
				html = html & "</ul></div>"
			end if
			dr.close()
			
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
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function totalRespostasEnquete(byval id as integer) as integer
		Dim retorno as integer
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand

		try	
			conn.Open()
			cmd = New SqlCommand("select sum(p.qtdresp) as Total from perguntas_enquetes as p inner join enquetes_netportal as e on e.id = p.idEnquete where p.idenquete = @idEnquete and e.idCidade = @idCidade", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idEnquete", SqlDbType.Int))
				.Parameters("@idEnquete").Value = id
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()			
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
	Public function getEnquete(byval id as integer, byval coluna as string) as String
		Dim retorno as string
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader

		try	
			conn.Open()
			cmd = New SqlCommand("select top 1 * from enquetes_netportal where id = @id", conn)
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
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function resultadoEnquete(byval id as integer, byval totalVotos as integer) as string
		Dim html as string
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		
		try	
			conn.Open()
			cmd = New SqlCommand("select p.id, p.resp, p.qtdresp from perguntas_enquetes as p inner join enquetes_netportal as e on e.id = p.idEnquete where p.idenquete = @idEnquete and e.idCidade = @idCidade order by resp asc;", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idEnquete", SqlDbType.Int))
				.Parameters("@idEnquete").Value = id
				.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
				.Parameters("@idCidade").Value = idCidade()			
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				while dr.Read()
					html = html & "<li>"
					html = html & "<div class=""total""><span class=""numero"">" & formata(Porcentagem(dr(2),totalVotos)) & "</span><span class=""percentual"">%</span></div>"
						html = html & "<div class=""resposta"">"
							html = html & "<span class=""opcao"">" & dr(1).toString() & "</span>"
							html = html & "<div class=""grafico"">"
								html = html & "<span style=""width:" & formata(Porcentagem(dr(2),totalVotos)) & "%""></span>"
							html = html & "</div>"
					html = html & "</div>"
					html = html & "</li>"
				end while				
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
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public function podeVotarEnquete(byval id as integer) as boolean
		Dim retorno as boolean
		Dim ipInternauta as String = _lib.getSessionID()	
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand

		try	
			conn.Open()
			cmd = New SqlCommand("select count(r.id) from respostas_enquetes as r inner join enquetes_netportal as e on e.id = r.idEnquete where r.idenquete = @idEnquete and r.ip = @ipInternauta", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idEnquete", SqlDbType.Int))
				.Parameters("@idEnquete").Value = id
				.Parameters.Add(New SqlParameter("@ipInternauta", SqlDbType.VarChar))
				.Parameters("@ipInternauta").Value = ipInternauta
			end with
			if cmd.ExecuteScalar() = 0 then retorno = true
			
		Catch ex As Exception
			retorno = false
		Finally
			conn.Close()
			conn.Dispose()
		End Try		
		
		return retorno
	end function	
	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public sub VotarEnquete(byval idEnquete as integer, byval idResposta as integer)
		Dim retorno as boolean
		Dim ipInternauta as String = _lib.getSessionID()	
		
		Dim conn As SqlConnection = New SqlConnection(sqlServer())
		Dim cmd As SqlCommand

		try	
			conn.Open()
			cmd = New SqlCommand("update perguntas_enquetes set QtdResp = QtdResp+1 where id = @idResposta and idEnquete = @idEnquete", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idResposta", SqlDbType.Int))
				.Parameters("@idResposta").Value = idResposta
				.Parameters.Add(New SqlParameter("@idEnquete", SqlDbType.Int))
				.Parameters("@idEnquete").Value = idEnquete
				.ExecuteNonQuery()
			end with
		
			cmd = New SqlCommand("insert into respostas_enquetes (idEnquete, resp, ip) values (@idEnquete, @idResposta, @ipInternauta)", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@idEnquete", SqlDbType.Int))
				.Parameters("@idEnquete").Value = idEnquete
				.Parameters.Add(New SqlParameter("@idResposta", SqlDbType.Int))
				.Parameters("@idResposta").Value = idResposta			
				.Parameters.Add(New SqlParameter("@ipInternauta", SqlDbType.VarChar))
				.Parameters("@ipInternauta").Value = ipInternauta				
				.ExecuteNonQuery()
			end with	
		Catch ex As Exception
		
		Finally
			conn.Close()
			conn.Dispose()
		End Try

	end sub		
	
	Private function formata(valor as String) as String
		valor = cDec(replace(valor,"%",""))
		Dim retorno as String
		Dim numero As Double = valor
		numero = System.Math.round(numero, 0)
		retorno = numero.toString()
		if numero < 10 then retorno = "0" & numero
		return retorno.toString
	end function
	
	Private function Porcentagem(valor as Integer, total as Integer) as String
		Dim totalResp as String
		
		if IsNothing(Total) OR Total = 0 then
			Total = 0
		end if
		
		if Total < 1 then
			TotalResp = FormatPercent(valor/1)
		else
			TotalResp = FormatPercent(valor/Total,2)
		end if
		
		return TotalResp
		
	end function		
																
End Class