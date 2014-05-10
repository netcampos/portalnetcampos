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

Public Module DadosCidade

	Public readonly webUrlCidadeOriginal as string = ConfigurationSettings.AppSettings("urlOriginal")
	Public readonly webUrlCidade as string = ConfigurationSettings.AppSettings("urlCidade")
	Public readonly webNomeAmigavel as string = ConfigurationSettings.AppSettings("nomeAmigavel")
	Public readonly webLogoCidade as string = ConfigurationSettings.AppSettings("logoTopCidade")	
	Public webNTemplate as string = ConfigurationSettings.AppSettings("ntemplate")
	Public webTemplate as string = ConfigurationSettings.AppSettings("template")
	Public readonly webSubDominio as boolean = ConfigurationSettings.AppSettings("useSubDominio")
	
	'membros
	Public userLogado as boolean = false
	Public userMembro, nomeMembro, idMembro as String
	Public customerID as Integer 'ID do Membro
	Public pgURLRef as string 'pega url referencia do internauta antes de logar
	'frmcadastro
	Public nome, email as string
	
	'extras
	Public existe_area as boolean = false 'verifica se existe area em questao
	Public exibe_title_hd as boolean = false 'define se deseja exibir o titulo principal da sessão na pagina
	Public pubTopAds as boolean = true 'exibe anuncios top categoria
	Public LocalBanner as Integer	
	
	'############# variaveis globais ##################'
	Public idmodulo, idconteudo, idCategoria, idSubCategoria, pgMenu, idusuario, pagefull, SlideShow as integer
	Public colunaPub, area, target, masterPg, urlMasterPg, catPg, urlCatPg, nomePg, urlPg, pagina, titulopg, urlCanonical, urlCurta, searchYoutube, searchFlickr as String
	
	'data
	Public readonly dataAtual As DateTime = DateTime.Now
	Public dataPub, dataRev as String	
	
	'variaveis barra de ferramentas e social bookmarks
	Public SocialBookMarks as integer
	Public tituloSocialBookMarks as string 'texto com o resumo do que vai ser indicado para o amigo
	Public liberacomentarios as Integer = 0 'libera comentarios na pagina
	Public PrintVersion as boolean = false
	Public assuntoFormAmigo as string 'define o assunto referente ao envio de indicação para um amigo
	Public resumoFormAmigo as string 'texto com o resumo do que vai ser indicado para o amigo	
	Public readonly ipInternauta as String = HttpContext.Current.request.ServerVariables("REMOTE_ADDR")
	
	Public metaTitle as String = dados_cidade(idcidade, "title_pagina")
	Public metaDescription as String = dados_cidade(idcidade, "meta_description")
	Public metaKeywords as String = dados_cidade(idcidade, "keywords")
	Public metaRobots as String = "all"
	Public metaAutor as String = "Alan C. Germano"
	Public verifycaptcha as string = String.Empty

	'###### variaveis informacoes da cidade corrente ###############
	Public cidade as string = dados_cidade(idcidade, "cidade") 'Migrado
	Public template as string = iif_template(webTemplate, dados_cidade(idcidade, "template")) 'Migrado
	Public TemplateFull as Integer
	Public cssPageDefaultModulo, cssPageModulo as String
	Public nome_template as string = iif_template(webNTemplate, dados_cidade(idCidade, "ntemplate")) 'Migrado

	Public google_cse as String = dados_cidade(idcidade, "google_cse")
	Public nome_amigavel as String = iif_template(webNomeAmigavel, dados_cidade(idcidade, "nome_amigavel"))
	Public urlCidade as string = iif_template(webUrlCidade, lcase(dados_cidade(idcidade, "urlcidade")) )
	Public urlCidadeWWW as string = lcase(replace(urlCidade,"http://",""))
	Public urlCidadeWWW2 as string = lcase(replace(urlCidadeWWW,"www.",""))
	Public urlEstado as String = dados_cidade(idcidade, "urlestado") 'Migrado
	Public emailContatoCidade as string = lcase(dados_cidade(idcidade, "emailcidade"))
	Public emailComercialCidade as string = lcase(dados_cidade(idcidade, "email_comercial"))
	Public emailFinanceiroCidade as string = lcase(dados_cidade(idcidade, "email_financeiro"))
	
	'Public urlfacebook as string = dados_cidade(idcidade, "urlfacebook")
	'Public urltwitter as string = dados_cidade(idcidade, "user_twitter")
	'Public urlOrkut as string = dados_cidade(idcidade, "urlorkut")
	
	Public liberarVendaPub as integer = dados_cidade(idcidade, "liberar_venda_pub")
	Public logoTopCidade as string = iif_template(webLogoCidade, dados_cidade(idcidade, "pasta_padrao") & "templates/" & nome_template & "/" & template & "/images/logo-guia.png.ashx" )
	Public breadcrumb as string = "<a class=""root"" href=""" & urlCidade & """ title=""Ir para: Página Inicial"" rel=""v:url"" property=""v:title"">Home</a>"
	
	'######## variaveis dados franquia cidade corrente #######################'
	Public nomeFranquia as string = dados_cidade(idcidade, "nFranquia")
	Public razaoFranquia as string = dados_cidade(idcidade, "razaosocial")
	Public cnpjFranquia as string = dados_cidade(idcidade, "cnpj")
	Public endFranquia as string = dados_cidade(idcidade, "endFranquia")
	Public bairroFranquia as String = dados_cidade(idcidade, "bairroFranquia")
	Public cidadeFranquia as String = dados_cidade(idcidade, "cidadeFranquia")
	Public estadoFranquia as String = dados_cidade(idcidade, "estadoFranquia")
	Public tipoPessoaFranquia as String = dados_cidade(idcidade, "tipoPessoa")
	'###### variaveis netmail cidade corrente ###########################'
	Public netmailAtivo as integer = dados_cidade(idcidade, "netmailAtivo")
	
	Public function dados_cidade(idcidade as integer, coluna as String) as String
	
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("exec proc_cidades_netportal @id_cidade", conn)
			cmd.Parameters.Add(New SqlParameter("@id_cidade", SqlDbType.Int))
			cmd.Parameters("@id_cidade").Value = idcidade
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
	
	Public function execHomePage(ByVal caminho as string, ByVal area as String) as string
		'create cache page
		'Dim urlCidade as string = replace(city.dados(idCidade, "urlcidade"),"http://","")
		'if area = "home" then
			'If IO.File.Exists( httpContext.Current.server.mappath("/App_Cache/" & urlCidade & "/home.aspx") ) Then	
				'httpContext.Current.server.execute("/App_Cache/" & urlCidade & "/home.aspx")
				'httpContext.Current.response.end()
				'Exit Function
			'End If
		'end if
						
		Dim pastaBase as String = httpContext.Current.server.mappath("/App_Themas/" & nome_template & "/paginas/base/")
		Dim pastaDestino as String = httpContext.Current.server.mappath("/App_Themas/" & nome_template & "/paginas/" & area & "/")
		
		Dim arquivo as string = caminho
		
		if inStr(arquivo, "?") > 0 then
			arquivo = left(arquivo, inStr(arquivo, "?") - 1)
		end if	
		
		arquivo = httpContext.Current.server.mappath(arquivo)
	 	
		If IO.File.Exists(arquivo) Then	
			execPagina(lcase(caminho))
			Exit Function
		else
			If System.IO.File.Exists(pastaBase & "home.aspx") Then
				
				If Not System.IO.Directory.Exists(pastaDestino) then
					System.IO.Directory.CreateDirectory(pastaDestino)
				End If
				
				System.IO.File.Copy(pastaBase & "home.aspx", arquivo)
				System.IO.File.Copy(pastaBase & "home.aspx.vb", replace(arquivo,"aspx","aspx.vb"))				
				
			End If
			error_404("Página: <b>" & area & "</b>, Não Encontrada foi criada com sucesso... recarregue para visualizar")
		end if
		
	End function
	
		
	Public function execPageConteudo(byVal area as String, ByVal idConteudo as Integer, ByVal idModulo as Integer) as string
		Dim pathFileBase as String = httpContext.Current.server.mappath("/App_Themas/" & nome_template & "/paginas/" & area & "/")
		Dim arquivo as string = ("/App_Themas/" & nome_template & "/paginas/" & area & "/pagina.aspx?idconteudo=" & idconteudo & "&idmodulo=" & idModulo)
		Dim caminho as String = arquivo		
		
		if inStr(arquivo, "?") > 0 then
			arquivo = left(arquivo, inStr(arquivo, "?") - 1)
		end if	
		
		arquivo = lcase(httpContext.Current.server.mappath(arquivo))
		
		If System.IO.File.Exists(arquivo) Then	
			execPagina(lcase(caminho))
			Exit Function
		Else
			If System.IO.File.Exists(pathFileBase & "home.aspx") Then
				System.IO.File.Copy(pathFileBase & "home.aspx", arquivo)
				System.IO.File.Copy(pathFileBase & "home.aspx.vb", replace(arquivo,"aspx","aspx.vb"))	
			else
				error_404("Arquivo Mestre Não Definido - Impossível Criar Página Filha")
			end if
			error_404("Página:" & caminho & ", Não Encontrada")		
		End If
		
	End function	
	
	Public function execPageCategoria(area as string,idconteudo as integer,id_modulo as integer) as string
		Dim arquivo as string = ("/App_Themas/" & nome_template & "/paginas/" & area & "/paginaCategoria.aspx?idconteudo=" & idconteudo & "&idmodulo=" & id_modulo)
		Dim caminho as String = arquivo		
		
		if inStr(arquivo, "?") > 0 then
			arquivo = left(arquivo, inStr(arquivo, "?") - 1)
		end if	
		
		arquivo = httpContext.Current.server.mappath(arquivo)
		
		If System.IO.File.Exists(arquivo) Then	
			execPagina(lcase(caminho))
			Exit Function			
		else
			error_404("Página: <u>" & caminho & "</u>, Não Encontrada")
		end if				
		
	end function
	
	Public Sub execPageSubCategoria(area as string,idconteudo as integer,idsubcategoria as integer,id_modulo as integer)				
		
		Dim pathFileBase as String = httpContext.Current.server.mappath("/App_Themas/" & nome_template & "/paginas/" & area & "/")
		Dim arquivo as string = ("/App_Themas/" & nome_template & "/paginas/" & area & "/paginaSubCategoria.aspx?idsubcategoria=" & idsubcategoria & "&idconteudo=" & idconteudo & "&idmodulo=" & id_modulo)
		Dim caminho as String = arquivo

		if inStr(arquivo, "?") > 0 then
			arquivo = left(arquivo, inStr(arquivo, "?") - 1)
		end if	
		
		arquivo = lcase(httpContext.Current.server.mappath(arquivo))
		
		If System.IO.File.Exists(arquivo) Then	
			execPagina(lcase(caminho))
			Exit Sub
		Else
			If System.IO.File.Exists(pathFileBase & "paginaCategoria.aspx") Then
				System.IO.File.Copy(pathFileBase & "paginaCategoria.aspx", arquivo)
				System.IO.File.Copy(pathFileBase & "paginaCategoria.aspx.vb", replace(arquivo,"aspx","aspx.vb"))	
			else
				error_404("Arquivo Mestre Não Definido - Impossível Criar Página Filha")
			end if
			error_404("Página:" & caminho & ", Não Encontrada")		
		End If
		
	End Sub	
	
	Public Sub execPagina( ByVal caminho as String)
		try
			HttpContext.Current.Server.Execute(caminho)
		catch ex as exception
			error_SQL("<div style=""width:auto;border: solid 1px #666666; padding:10px; margin:0 auto;"" ><h3 style=""margin:0;"">Erro de Execução:</h3><hr/>Ocorreu um erro Durante o Carregamento da Página, por favor tente novamente mais tarde.<br><br> <b>Atenção!!!</b><hr/>Se o erro persistir informe ao administrador do sistema.</div>" )
		end try
	End Sub
		
    Public Function iif_url(ByVal webUrlCidade As String, ByVal urlDataSet As String) As String
        Dim urlDataSetNova, urlTransform As String
        urlDataSet = replace(urlDataSet, "http://", "")

        If Not webSubDominio Then
            urlDataSet = replace(urlDataSet, "http://", "")
            If InStr(1, urlDataSet, ".") Then
                Dim verifica() As String
                verifica = urlDataSet.split(".")
                If verifica(0) <> "www" Then
                    Dim index As Integer = InStr(1, urlDataSet, "/")
                    If index > 0 Then
                        Dim np() As String
                        urlDataSetNova = replace(Left(urlDataSet, InStr(urlDataSet, "/") - 1), "http://", "")
                        np = urlDataSetNova.split(".")
                        If np(0) <> "www" Then
                            urlDataSet = replace(urlDataSet, np(0), "www")
                            urlDataSetNova = replace(urlDataSetNova, np(0), "www")
                            urlTransform = replace(urlDataSetNova, np(0), "www") & "/" & np(0) & "/"
                            urlDataSet = replace(replace(urlDataSet, urlDataSetNova, urlTransform), "//", "/")
                        End If
                    End If
                End If
            End If
        End If


        Dim urlRefCidade As String = webUrlCidadeOriginal
        urlRefCidade = replace(webUrlCidadeOriginal, "http://www.", "")
        Dim urlDestCidade As String = webUrlCidade
        urlDestCidade = replace(urlDestCidade, "http://www.", "")
        Dim urlDestino As String = urlDataSet
        If webUrlCidade <> "" Then urlDestino = replace(urlDestino, urlRefCidade, urlDestCidade)
        Return ("http://" & urlDestino)

    End Function

End Module
