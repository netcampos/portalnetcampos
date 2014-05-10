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
Imports MySql.Data
Imports MySql.Data.MySqlClient
Imports System.Net
Imports System.Net.Mail

Public Class ads
	Public cidade, estado, pais, latitude, longitude, HostRefer as string	
	public shared _idCityADS as integer

	public shared function adPub(idArea as integer, idLocal as integer) as string
	
		if (not IsNothing(_idCityADS) and _idCityADS <> 0 ) then
			idCidade = _idCityADS
		End if
			

		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
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
				return exibeBannerGoogle(idArea)
			else
				return exibeBannerCliente(idArea, idLocal)
			end if
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
		End Try	
	end function
	
	public shared function exibeBannerGoogle(idArea as Integer) as string
		Dim googleads as integer
		Dim html as String
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		
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
					return (dr(1).toString())
				end if				
			end if
			dr.close
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		return "<a href=""" & city.urlCidade() & """ target=""_blank""><img src=""/App_Themas/" & city.nTemplate() & "/styles/" & city.Template() & "/images/top_banners/top_banner_reserve.gif"" width=""165"" width=""115"" /></a>"
	end function
	
	public shared function exibeBannerCliente(idArea as Integer, idLocal as Integer) as string
		Dim contar as Integer = 0
		Dim googleads as integer
		Dim html as string
			
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)

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
					return exibeBannerGoogle(idArea)
				else
					'html = html & "<script type=""text/javascript"">var _OAS_AD(""" & dr(0).toString() & """);</"
					'html = html & "script>"							
					'html = html & "<script type=""text/javascript"" src=""http://adserver.guiadoturista.net/ad/" & dr(0).toString() & ".js""></"
					'html = html & "script>"
					html = showAd(dr(0).toString())
					return html
				end if			
			end if
			dr.close
		Catch ex As Exception
			return "cliente"
		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try			
		
		return "cliente"
		
	end function
	
	Public shared function showAd(ByVal id as String) as string
		dim swf as string
		Dim extensao as string
		Dim arquivo as string
		Dim largura as string
		Dim altura as string
		Dim urlDestino as string
		Dim txtUrl as string		
		
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)		
		
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
				arquivo = "http://adserver.guiadoturista.net" & replace(arquivo,extensao,"")
				urlDestino = dr(4).toString
				txtUrl = dr(5).toString
				
				largura = dr(2).toString
				altura = dr(3).toString
				
				if dr(1).toString = "swf" then
					return embedPlayer ( arquivo & ".swf", largura, altura, "clickTAG=http://adserver.guiadoturista.net/?catads=2%26ads=" & id, id.toString)
				end if
				
				if dr(1).toString = "gif" then
					arquivo = arquivo & extensao & ".ashx"
					return "<a href=""http://adserver.guiadoturista.net/?catads=2&amp;ads=" & id & """ title=""" & txtUrl & """ target=""_blank""><img src=""" & arquivo & """ alt=""" & txtUrl & """ width=""" & largura & """ height=""" & altura & """ /></a>"	
				end if				
				
			end if
			dr.close()
		catch ex as exception
			return "cliente2"
		finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		end try				
	
	end function
	
	
	Public shared function LinksAds() as string	
		Dim html as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)		
	
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

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		
		return html
	
	End Function
	
	Public shared function LinksAds_topSearch() as String
		Dim html as String = String.Empty
		Dim conn As SqlConnection
		Dim cmd As SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)		
	
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select top 2 idempresa, empresa, txtdescritivo, url, chave_md5 from empresas where idcidade = @id_cidade and idcliente > 0 and bloqueado = 0 and plano in(5,6) order by newid()", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
				.Parameters("@id_cidade").Value = idCidade
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				html = html & "<ul class=""ads-links"">"
				While dr.Read()
					html = html & "<li>"
					html = html & "<div class=""ad-link""><a href=""http://adserver.guiadoturista.net/?catads=1&amp;ads=" & dr("chave_md5").tostring() & """ target=""_blank""><span>" & dr(1).toString() & "</span></a></div>"
					html = html & "<div class=""adb-link"">" & mid(dr(2).tostring(),1,250) & "...</div>"					
					html = html & "<div class=""ad-link-r""><a href=""http://adserver.guiadoturista.net/?catads=1&amp;ads=" & dr("chave_md5").tostring() & """ target=""_blank"">" & replace(dr("url").tostring(),"http://","") & "</a></div>"			
					html = html & "</li>"
				end while
				html = html & "</ul>"
			end if
			dr.close
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try
		
		return html
	end function	


	public sub logLink(idCidade, idCliente, nomeProduto, nomeCliente, urlCliente, localProduto, idAds)
	
		Dim geo as boolean
		
		
		Dim ipInternauta as String = HttpContext.Current.request.ServerVariables("REMOTE_ADDR")
		Dim urlReferencia as String = HttpContext.Current.Request.ServerVariables("HTTP_REFERER")
		if urlReferencia = "" then
			urlReferencia = "Nao Identificado"	
		End if	
		
		HostRefer = "Não Identificado"
		pais = "desconhecido"
		estado = "desconhecido"
		cidade = "desconhecido" 
		latitude = "desconhecido" 
		longitude = "desconhecido"
		Dim objX As System.Net.IPHostEntry
		Try
			objX = System.Net.Dns.GetHostByAddress(ipInternauta)
			HostRefer = objX.HostName
		Catch ex As Exception
		 
		End Try 
				
		geo = geoip(ipInternauta)			
		
		Dim mysql_stats As String = ConfigurationManager.ConnectionStrings("mysqlStats").ConnectionString
		Dim conn_mysql as New MySqlConnection(mysql_stats)
		Dim cMysql as MySqlCommand

		try
			conn_mysql.open()
			dim cMysqlup = New MySqlCommand()
			with cMysqlup
				.commandText = "insert into tbl_logAnuncio (idCidade,idCliente,nomeProduto,nomeCliente,urlCliente,urlReferencia,ipInternauta,cidade,estado,pais,latitude,longitude,view,clique,localproduto,idADS,idCatADS,host) values (@idCidade,@idCliente,@nomeProduto,@nomeCliente,@urlCliente,@urlReferencia,@ipInternauta,@cidade,@estado,@pais,@latitude,@longitude,@view,@clique,@localproduto,@idADS,@idCatADS,@HostRefer)"
				.connection = conn_mysql			
				.Parameters.Add(New MySqlParameter("@idCidade", MySqlDbType.Int32))
				.Parameters("@idCidade").Value = idCidade
				.Parameters.Add(New MySqlParameter("@idCliente", MySqlDbType.Int32))
				.Parameters("@idCliente").Value = idCliente
				.Parameters.Add(New MySqlParameter("@nomeProduto", MySqlDbType.String))
				.Parameters("@nomeProduto").Value = nomeProduto
				.Parameters.Add(New MySqlParameter("@nomeCliente", MySqlDbType.String))
				.Parameters("@nomeCliente").Value = nomeCliente
				.Parameters.Add(New MySqlParameter("@urlCliente", MySqlDbType.String))
				.Parameters("@urlCliente").Value = urlCliente
				.Parameters.Add(New MySqlParameter("@urlReferencia", MySqlDbType.String))
				.Parameters("@urlReferencia").Value = urlReferencia
				.Parameters.Add(New MySqlParameter("@ipInternauta", MySqlDbType.String))
				.Parameters("@ipInternauta").Value = ipInternauta
				.Parameters.Add(New MySqlParameter("@cidade", MySqlDbType.String))
				.Parameters("@cidade").Value = cidade	
				.Parameters.Add(New MySqlParameter("@estado", MySqlDbType.String))
				.Parameters("@estado").Value = estado	
				.Parameters.Add(New MySqlParameter("@pais", MySqlDbType.String))
				.Parameters("@pais").Value = pais	
				.Parameters.Add(New MySqlParameter("@latitude", MySqlDbType.String))
				.Parameters("@latitude").Value = latitude	
				.Parameters.Add(New MySqlParameter("@longitude", MySqlDbType.String))
				.Parameters("@longitude").Value = longitude
				.Parameters.Add(New MySqlParameter("@view", MySqlDbType.Int32))
				.Parameters("@view").Value = 0
				.Parameters.Add(New MySqlParameter("@clique", MySqlDbType.Int32))
				.Parameters("@clique").Value = 1
				.Parameters.Add(New MySqlParameter("@localproduto", MySqlDbType.Int32))
				.Parameters("@localproduto").Value = localproduto
				.Parameters.Add(New MySqlParameter("@idADS", MySqlDbType.String))
				.Parameters("@idADS").Value = idADS	
				.Parameters.Add(New MySqlParameter("@idCatADS", MySqlDbType.Int32))
				.Parameters("@idCatADS").Value = 1
				.Parameters.Add(New MySqlParameter("@HostRefer", MySqlDbType.String))
				.Parameters("@HostRefer").Value = HostRefer
				.ExecuteNonQuery()
			end with
		catch ex as Exception
			
		Finally
			conn_mysql.close() : conn_mysql.dispose()
		End try	

	
	End Sub
			
			
	Private Function geoip(ip) as boolean
	
		Dim servidor1 = "8QOJEQFllRUH"
		Dim servidor2 = "V69Wmy5jELay"
		Dim servidor3 = "V69Wmy5jELayi"		
		
		Dim url = "http://maxmind.com:8010/b?l=" & servidor1 & "&i=" & ip
		try
			Dim req As WebRequest = WebRequest.Create(url)
			Dim res As WebResponse = req.GetResponse()
			Dim reader As StreamReader = New StreamReader(res.GetResponseStream(), System.Text.Encoding.GetEncoding("ISO-8859-1") )
			
			Dim str As String = reader.ReadLine()
			dim meuarray = Split(str, ",", -1, 1)
			pais = meuarray(0)
			estado = meuarray(1)
			cidade = meuarray(2)
			if cidade = "(null)" then
				cidade = "Cidade não identificada"
			end if
			latitude = meuarray(3)
			longitude = meuarray(4)		
		catch ex as exception
		
		end try
		
		try
			if latitude = "" then  
				url = "http://maxmind.com:8010/b?l=" & servidor2 & "&i=" & ip
				Dim req As WebRequest = WebRequest.Create(url)			
				Dim res As WebResponse = req.GetResponse()
				Dim reader As StreamReader = New StreamReader(res.GetResponseStream(), System.Text.Encoding.GetEncoding("ISO-8859-1") )
				Dim str As String = reader.ReadLine()
				dim meuarray = Split(str, ",", -1, 1)
				pais = meuarray(0)
				estado = meuarray(1)
				cidade = meuarray(2)
				if cidade = "(null)" then
					cidade = "Cidade não identificada"
				end if
				latitude = meuarray(3)
				longitude = meuarray(4)	
				
			end if
		catch ex as exception
		
		end try
		
		
		if estado <> "(null)" and lcase(pais) = lcase("br") then	
		
			select case estado
				case "01"
					estado = "Acre"
				case "02"
					estado = "Alagoas"
				case "03"
					estado = "Amapa"
				case "04"
					estado = "Amazonas"
				case "05"
					estado = "Bahia"
				case "06"
					estado = "Ceara"
				case "07"
					estado = "Distrito Federal"
				case "08"
					estado = "Espirito Santo"
				case "11"
					estado = "Mato Grosso do Sul"
				case "13"
					estado = "Maranhao"
				case "14"
					estado = "Mato Grosso"
				case "15"
					estado = "Minas Gerais"	
				case "16"
					estado = "Para"
				case "17"
					estado = "Paraiba"
				case "18"
					estado = "Parana"	
				case "20"
					estado = "Piaui"
				case "21"
					estado = "Rio de Janeiro"
				case "22"
					estado = "Rio Grande do Norte"
				case "23"
					estado = "Rio Grande do Sul"
				case "24"
					estado = "Rondonia"
				case "25"
					estado = "Roraima"
				case "26"
					estado = "Santa Catarina"
				case "27"
					estado = "Sao Paulo"
				case "28"
					estado = "Sergipe"
				case "29"
					estado = "Goias"
				case "30"
					estado = "Pernambuco"
				case "31"
					estado = "Tocantins"																																						
			end select
		else
				estado = "Desconhecido"
		end if
	
		return true
	End Function			
			
End Class