Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.IO
Imports System.Xml
Imports System.Configuration
Imports System.Text
Imports System.Security.Cryptography
Imports System.Web
Imports System.Web.Configuration
Imports System.Web.Management
Imports System.Web.Security
Imports System.Web.Caching
Imports System.Web.UI.UserControl

Public Module netPage

	'public members Module
	Public pageLevel as integer = 0
	Public pageID as integer = 0
	Public pageExists as Boolean = true
	Public pageModule as Integer
	Public pageName as String

	Public Class pages
		
		'private members
		Private idCidade as Integer = ConfigurationSettings.AppSettings("id_cidade")
		Private siteURL as string = "http://www.netcampos.com"
		Private uri as string = _lib.parse(_lib.QS("uri"),"")
		Private REWRITE_URL as string = _lib.REWRITE_URL()
		Private modulo as integer = _lib.parse(_lib.QS("modulo"),"")
		Private adArea as Integer = _lib.parse(_lib.QS("ads"),1)
		Private pageSlug, moduleSlug, pageView, pageType, ogURL as String
		Private ogImage as String = "http://www.netcampos.com/assets/images/marca_og.gif"
		Private str as new StringOperations()
		Private user as new npUsers()
		Private db as new database()
		
		'private members meta headers
		Private p_seoTitle As String
		Private p_seoDescription As String
		Private p_seoKeyWords As String
	
		'Private Property
		Private property seoTitle() As String
			Get
				return p_seoTitle
			End Get
			Set(ByVal value As String)
				p_seoTitle = value
			End Set
		End Property
		
		Private property seoDescription() As String
			Get
				return p_seoDescription
			End Get
			Set(ByVal value As String)
				p_seoDescription = value
			End Set
		End Property
		
		Private property seoKeyWords() As String
			Get
				return p_seoKeyWords
			End Get
			Set(ByVal value As String)
				p_seoKeyWords = value
			End Set
		End Property	
		
		Private Property sqlServer() As String
			Get
				Return ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
			End Get
			Set(ByVal value As String)
			
			End Set
		End Property	
	
		'******************************************************************************************************************
		'' @SDESCRIPTION:	writes a string to the output in the same line
		'' @PARAM:			value [string]: output string
		'******************************************************************************************************************
		Public Sub New()
			MyBase.New()						
		End Sub
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:	writes a string to the output in the same line
		'' @PARAM:			value [string]: output string
		'******************************************************************************************************************	
		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub
		
		'**********************************************************************************************************
		'' @SDESCRIPTION: Draw MetaHeader Page
		'**********************************************************************************************************
		Public Sub metaHeaderModule()
			
		End Sub	
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:	writes a string to the output in the same line
		'' @PARAM:			value [string]: output string
		'******************************************************************************************************************
		Private sub setHttpHeader()
			HttpContext.Current.Response.Clear()
			HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
			Dim ts As New TimeSpan(0,60,0)
			HttpContext.Current.Response.Cache.SetMaxAge(ts)
			HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddSeconds(3600))
			HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
			HttpContext.Current.Response.Cache.SetValidUntilExpires(true)
			HttpContext.Current.Response.Cache.VaryByHeaders("Accept-Language") = true
			HttpContext.Current.Response.Cache.VaryByHeaders("User-Agent") = true
			HttpContext.Current.Response.AppendHeader("X-Powered-By","Grupo NetCampos Tecnologia")
		End sub	
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:	writes a string to the output in the same line
		'' @PARAM:			value [string]: output string
		'******************************************************************************************************************	
		Public function Render()
			setHttpHeader()
			pageModule = modulo
			moduleSlug = db.dadosModulo(modulo, "slugmodulo")			
			pageSlug = "/" & db.dadosModulo(modulo, "slug") & "/"
			pageType = db.dadosModulo(modulo, "categoria")
			pageView = "/App/web/views/" & pageType
			
			'if instr(uri,"?") > 0 then uri = replace(uri,"?","&")
			'_lib.write(uri)
			'_lib.responseEnd()
			
			if _lib.LenB(uri)=0 then
				pageLevel = 0
				p_seoTitle = db.metaHeader(modulo,"meta_title") & " | NetCampos"
				p_seoDescription = db.metaHeader(modulo,"meta_description")
				p_seoKeyWords = db.metaHeader(modulo,"meta_keywords")
				pageView = pageView & "/home.ascx"
				
				ogURL = siteURL & pageSlug
			else
				'REWRITE_URL = replace(REWRITE_URL,pageSlug,"")								
				'_lib.extractPageFromURI(REWRITE_URL)
				_lib.extractPageFromURI(uri)
				
				pageID = db.pageInfo(pageType,pageName,"id")				
				
				if(pageID=0) then 
					_lib.pageNotFound()
				else
					p_seoTitle = db.pageInfo(pageType,pageName,"seoTitle") & " | NetCampos"
					p_seoDescription = db.pageInfo(pageType,pageName,"seoDescription")
					p_seoKeyWords = db.pageInfo(pageType,pageName,"seoKeyWords")
				end if
								
				if pageLevel = 3 then
					dim thumbNail as string = db.pageInfo(pageType,pageName,"thumbNail")
					if _lib.lenB(thumbNail) > 0 then ogImage = replace(thumbNail,"cidades/cms/netgallery/media/saopaulo/camposdojordao/","imagens-campos-do-jordao/")
					ogURL = siteURL & pageSlug & pageName & ".html"
					pageView = pageView & "/page.ascx" 'page content
				end if
				
			end if
			drawHeader()
			drawContent()
			drawFooter()
			
		End Function
		
		'******************************************************************************************************************
		'' draw header
		'******************************************************************************************************************
		Private sub drawHeader()
			dim html as String = ""
			
			html = html & "<!DOCTYPE html>" & vbnewline
			html = html & "<html id=""netcampos"" dir=""ltr"" lang=""pt-BR"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:og=""http://opengraphprotocol.org/schema/"" xmlns:fb=""http://developers.facebook.com/schema/"">"& vbnewline
			html = html & "<head profile=""http://gmpg.org/xfn/11"">"& vbnewline
				html = html & "<meta charset=""UTF-8"">"& vbnewline
				html = html & "<title>"&seoTitle()&"</title>"& vbnewline
				html = html & "<meta http-equiv=""X-UA-Compatible"" content=""chrome=1"" />"& vbnewline
				if _lib.lenB(seoDescription()) > 0 then html = html & "<meta name=""description"" content="""&seoDescription()&""" />"& vbnewline
				if _lib.lenB(seoKeyWords()) > 0 then html = html & "<meta name=""keywords"" content="""&seoKeyWords()&""" />"& vbnewline
				html = html & "<meta name=""robots"" content=""index,follow"" />"& vbnewline
				html = html & "<meta name=""viewport"" content=""width=device-width"" />" & vbnewline
				html = html & "<link rel=""shortcut icon"" href=""http://www.netcampos.com/web_data/templates/netcampos/themes/verde/favicon.ico"" type=""image/x-icon"" />"& vbnewline
				html = html & "<link rel=""alternate"" type=""application/rss+xml"" title=""RSS 2.0"" href=""http://www.netcampos.com/feed/"" />"& vbnewline
				html = html & "<link href=""https://plus.google.com/100400043965768897941"" rel=""publisher"" />"& vbnewline
							
				'facebook
				html = html & "<meta property=""og:locale"" content=""pt_BR"" />"& vbnewline
				html = html & "<meta property=""fb:app_id"" content=""136304184521"" />"& vbnewline
				html = html & "<meta property=""og:title"" content="""&seoTitle()&""" />"& vbnewline
				html = html & "<meta property=""og:type"" content=""website"" />"& vbnewline
				html = html & "<meta property=""og:url"" content=""" & ogURL & """ />"& vbnewline
				html = html & "<meta property=""og:image"" content=""" & ogImage & """ />"& vbnewline
				html = html & "<meta property=""og:site_name"" content=""netcampos.com"" />"& vbnewline
				if _lib.lenB(seoDescription()) > 0 then html = html & "<meta property=""og:description"" content="""& seoDescription() &""" />"& vbnewline							
							
				html = html & loadCSS() & vbnewline
				html = html & loadJSHeader() & vbnewline
				html = html & "<!--[if lt IE 9]><script src=""http://html5shim.googlecode.com/svn/trunk/html5.js""></script><![endif]-->"& vbnewline
				
			html = html & "</head>"& vbnewline
			html = html & "<body id=""" & moduleSlug & """ class=""gnc-body"">"& vbnewline
				html = html & "<div id=""fb-root""></div><div id=""gnc-twitter""></div><div id=""gnc-google""></div>"& vbnewline
				html = html & "<div id=""wrapper-full"">"& vbnewline
					html = html & user.top_toolbar()
					html = html & contentHead()
			str.write(html)
		End Sub
		
		'******************************************************************************************************************
		'' draw footer
		'******************************************************************************************************************	
		Private function contentHead()
			dim html as string
			
				html = html & "<div id=""head"">"& vbnewline
					html = html & "<div class=""header"">"& vbnewline
						html = html & "<div class=""sombratopesq""></div>"& vbnewline
						html = html & "<div class=""sombratopdir""></div>"& vbnewline
						
						'logo
						html = html & "<div id=""logoPortal"" class=""logo"">"& vbnewline
							html = html & "<a class=""marca_cidade"" href=""" & siteURL & """ title=""Voltar para a Página Inicial"">"& vbnewline
								html = html & "<div class=""logotxt"">Campos do Jordão - Campos do Jordão com NetCampos</div>"& vbnewline
							html = html & "</a>"& vbnewline
						html = html & "</div>"& vbnewline
						
						'saudacao
						html = html & "<div class=""saudacao"">"& vbnewline
							html = html & "<p class=""line1""><strong>Campos do Jordão</strong>, seja bem vindo.</p>"& vbnewline
							html = html & "<p class=""line2"">"& FormatDateTime(dataAtual, DateFormat.LongDate)&".</p>"& vbnewline
						html = html & "</div>"& vbnewline
						
						'search bar
						html = html & "<div class=""top-search"">"& vbnewline
							html = html & "<div class=""bg-left"">"& vbnewline
								html = html & "<h3 class=""anuncie""><a title=""Publicidade - Campos do Jordão"" href=""/publicidade/""><span>Publicidade - Campos do Jordão</span></a></h3>"& vbnewline
								html = html & "<div class=""form-busca"">"& vbnewline
									html = html & "<form action=""/busca/"" name=""busca"" id=""frmbusca"">"& vbnewline
										html = html & "<fieldset>"& vbnewline
											html = html & "<input type=""hidden"" value=""pt-BR"" id=""hl"" name=""hl"" />" & vbnewline
											html = html & "<input type=""text"" value="""" id=""q"" class=""input-search"" name=""q"" />" & vbnewline
											html = html & "<input type=""hidden"" name=""ordenacao"" />" & vbnewline
											html = html & "<input type=""hidden"" name=""aba"" />" & vbnewline
											html = html & "<input type=""submit"" class=""btn-search"" value=""Buscar"" />"		& vbnewline							
										html = html & "</fieldset>"& vbnewline
									html = html & "</form>"& vbnewline
								html = html & "</div>"& vbnewline
							html = html & "</div>"& vbnewline
							html = html & "<div class=""bg-right""></div>"& vbnewline
						html = html & "</div>"& vbnewline
						
						'header menu
						html = html & "<div class=""nav"" id=""navTopMenu"">"& vbnewline
							html = html & db.topNavMenu()
						html = html & "</div>"& vbnewline																	
											
					html = html & "</div>"& vbnewline
				html = html & "</div>"& vbnewline			
			
			return html
		end function
			
		'***********************************************************************************************************
		'* draw Page Content 
		'***********************************************************************************************************
		private Sub drawContent()
			dim html as string
			html = html & "<div id=""wrapper"" class=""pagepub"">"& vbnewline
				'SkyBanner
				html = html & topADS(1,adArea)
				html = html & "<div id=""page"">"
					'conteudo
					html = html & Me.GenerateControlMarkup(pageView)
					'sidebar
					html = html & "<div id=""sidebar"">"
						html = html & "<div id=""wrapper-col-pub"">"
							html = html & "<div class=""title""><h3>publicidade</h3></div>"
							html = html & "<div class=""content"">"
								html = html & asideADS(13,adArea)
							html = html & "</div>"
							html = html & "<div class=""bot""></div>"
							
						html = html & "</div>"
					html = html & "</div>"
				
					html = html & "<div class=""clear""></div>"
				html = html & "</div>"
			html = html & "</div>"& vbnewline
			str.write(html)	
		end Sub	
		
		'******************************************************************************************************************
		'' draw footer
		'******************************************************************************************************************
		Private sub drawFooter()
			dim html as string
			'footer area
			html = html & "<div id=""footer"">"
				html = html & "<div class=""bg"">"
					html = html & "<div class=""sombraBotEsq""></div><div class=""sombraBotDir""></div>"
					html = html & navFooter()
				html = html & "</div>"
			html = html & "</div>"
			
			html = html & "</div>" & vbnewline 'end wrapper-full
			html = html & footerResources()
			html = html & "</body>"& vbnewline
			html = html & "</html>"& vbnewline
			str.write(html)	
		End Sub
		
		'***********************************************************************************************************
		'* draw footer Home 
		'***********************************************************************************************************	
		Private Function navFooter() as string
			dim html as String
				html = html & "<div class=""footer"">"
					'primary
					html = html & "<div class=""foot-primary"">"
						html = html & "<ul>"
							html = html & "<li><a href=""/"" rel=""nofollow"">Home</a></li>"
							html = html & "<li><a href=""/acidade/"" rel=""nofollow"">A Cidade</a></li>"
							html = html & "<li><a href=""/cadastro/"" rel=""nofollow"">Cadastro</a></li>"
							html = html & "<li><a href=""/fotos-campos-do-jordao/"" rel=""nofollow"">Fotos</a></li>"
							html = html & "<li><a href=""/noticias-campos-do-jordao/"" rel=""nofollow"">Notícias</a></li>"
							html = html & "<li><a href=""/passeios-campos-do-jordao/"" rel=""nofollow"">Passeios</a></li>"
							html = html & "<li><a href=""/empresas/"" rel=""nofollow"">Empresas</a></li>"
							html = html & "<li><a href=""/empresas/restaurantes/"" rel=""nofollow"">Restaurantes</a></li>"
							html = html & "<li><a href=""/acidade/dicas-de-hospedagem.html"" rel=""nofollow"">Hospedagem</a></li>"
							html = html & "<li><a href=""/empresas/utilidade-publica/"" rel=""nofollow"">Utilidade Pública</a></li>"
							html = html & "<li><a href=""/acidade/prefeitura-municipal.html"" rel=""nofollow"">Prefeitura</a></li>"
							html = html & "<li class=""final_lista""><a href=""/busca/"" rel=""nofollow"">Buscador</a></li>"
						html = html & "</ul>"
					html = html & "</div>"
					
					'secondary
					html = html & "<div class=""foot-secondary"">"
						html = html & "<ul>"
							'html = html & "<li><a href=""http://www.sobreeletrikusbrasiliensis.com.br/"">Eletrikus Brasiliensis</a></li>"
							'html = html & "<li><a href=""/parceiros/"" rel=""nofollow"">Parceiros</a></li>"
							'html = html & "<li><a href=""/press-release/expediente.html"" rel=""nofollow"">Expediente</a></li>"
							html = html & "<li><a href=""/press-release/aviso-legal.html"" rel=""nofollow"">Aviso Legal</a></li>"
							html = html & "<li><a href=""/press-release/direitos-autorais.html"" rel=""nofollow"">Direitos Autorais</a></li>"
							html = html & "<li><a href=""/press-release/politica-de-privacidade.html"" rel=""nofollow"">Politica de Privacidade</a></li>"
							html = html & "<li><a href=""/contato/"" rel=""nofollow"">Contato</a></li>"
							html = html & "<li><a href=""/publicidade/"" rel=""nofollow"">Publicidade</a></li>"
							html = html & "<li><a href=""/sitemap/"" title=""Mapa do Site"">Mapa do Site</a></li>"
							html = html & "<li class=""final_lista""><a href=""#"" rel=""nofollow"">Trabalhe Conosco</a></li>"
						html = html & "</ul>"						
					html = html & "</div>"
					'finaly
					html = html & "<div class=""foot-finaly"">"
						html = html & "<ul>"
							html = html & "<li class=""copy final_lista"">Copyright © 2004-" & year(dataAtual) & " - " & city.dados(idCidade,"Cidade") & " " & city.dados(idCidade,"UF") & " - <a href=""http://www.gruponetcampos.com.br"" target=""_blank"" style=""padding:0; color:#000000;"">Grupo NetCampos Tecnologia</a> -</li>"
							html = html & "<li class=""press""><a href=""/press-release/"" rel=""nofollow"">Press Release</a></li>"
							html = html & "<li class=""press""><a href=""/feed/"" rel=""nofollow"">RSS</a></li>"
							html = html & "<li class=""destinos""><a href=""http://negocios.guiadoturista.net/franquias/"" target=""_blank"">Franquias</a></li>"
							html = html & "<li class=""final_lista""><a href=""http://www.guiadoturista.net"" target=""_blank"">Guia do Turista</a></li>"
						html = html & "</ul>"				
					html = html & "</div>"
									
				html = html & "</div>"
			
			return html
		End Function		
		
		'***********************************************************************************************************
		'* draw Header 
		'***********************************************************************************************************	
		Private function loadCSS()
			dim html as string
			html = str.loadCSS("/assets/css/style.css,slimbox.css", string.empty)
			return html
		end function
		
		'***********************************************************************************************************
		'* draw Header 
		'***********************************************************************************************************	
		Private function loadJSHeader()
			dim html as string
			html = str.loadJS("/assets/js/jquery.min.js,jquery-ui.js")
			return html
		end function
		
		'***********************************************************************************************************
		'* footer Resources
		'***********************************************************************************************************
		Private Function footerResources() as string
			dim html as string
				html = html & str.loadJS("/assets/js/jquery.md5.js,jquery.metadata.js,jquery.form.js,jquery.validate.js") & vbnewline				
				html = html & str.loadJS("/assets/js/jquery.tools.min.js,submenu.js,slimbox.js,nyroModal.min.js,jquery.simplemodal.js,gnc.netUI.js") & vbnewline
				html = html & str.loadJS("/assets/js/jquery.maskedinput.js") & vbnewline				
				select case pageType
					case "noticias"
					html = html & str.loadJS("/assets/js/cycle.min.js,jquery.wslide.js,jquery.jcarousel.js,zoomimagethumb.js") & vbnewline
					if pageType = "noticias" then html = html & str.JS("$.pageNoticias({page:'home'});")
				end select
				html = html & str.loadJS("/assets/js/crossbrowser.js") & vbnewline
				html = html & city.dados(idCidade, "estatisticas")
			return html
		End Function
		
		'***********************************************************************************************************
		'* top ads sky banner
		'***********************************************************************************************************
		Private function topADS(idArea as integer, idLocal as integer) as string
			dim html as string
			
			html = html & "<div id=""advertising"">"
				
				html = html & "<div class=""sky"">"
					html = html & "<div class=""sky-bg-left""></div>"
					html = html & "<div class=""sky-banner"">"
						html = html & "<div id=""super_banner"">"
							html = html & db.ads(idArea,idLocal)
						html = html & "</div>"
					html = html & "</div>"
					html = html & "<div class=""sky-bg-right""></div>"
				html = html & "</div>"
				
				html = html & "<div class=""ad-weather"">"
					html = html & "<div class=""wea-bg-left""></div>"
					html = html & "<div id=""top-weather"">"
						html = html & "<div class=""tempo tempoHoje"">"
							html = html & weather_now()
						html = html & "</div>"
					html = html & "</div>"
					html = html & "<div class=""wea-bg-right""></div>"
				html = html & "</div>"
				
			html = html & "</div>"		
			
			return html	
		End Function
		
		'***********************************************************************************************************
		'* Top Banners
		'***********************************************************************************************************
		Private function asideADS(total as integer, idLocal as integer) as string
			dim area as integer = 1
			dim html as string
				html = html & "<ul class=""top-ads"">"
					if total > 0 then
						for i = 1 to total
							html = html & "<li>" & db.ads(area+i,idLocal) & "</li>"
						next
					end if
				html = html & "</ul>"
			return html
		End Function
		
		'***********************************************************************************************************
		'* Previsao do Tempo
		'***********************************************************************************************************		
		Private Function GenerateControlMarkup(ByVal virtualPath As String) As [String]
			Dim page As New NPMarkupPage()
			Dim ctl As UserControl = DirectCast(page.LoadControl(virtualPath), UserControl)
			page.Controls.Add(ctl)
			Dim sb As New StringBuilder()
			Dim writer As New StringWriter(sb)
		
			page.Server.Execute(page, writer, True)
			Return sb.ToString()
		End Function		
		
		'***********************************************************************************************************
		'* Previsao do Tempo
		'***********************************************************************************************************
		public function weather_now() as string
			Dim html, html2, data as String
			Dim dtDate As Date
			Dim erro as boolean = false
			Dim fileWeather as string = lcase("/ReferencesCache/") & DateTime.Now.ToString("dd") & "_" & DateTime.Now.ToString("MM") & "_" & DateTime.Now.ToString("yyyy") & ".xml"
			Dim xml as XmlDocument = new XmlDocument()
			Dim m_nodelist As XmlNodeList
			Dim m_node As XmlNode
			
			html2 = html2 & "<a class=""follow"" target=""_blank"" title=""Siga-nos no Twitter"" href=""http://twitter.com/netcampos/"">"
				html2 = html2 & "<span>Siga-nos no Twitter</span>"
			html2 = html2 & "</a>"		
			
			If not System.IO.File.Exists( _lib.mapPath(fileWeather) ) Then 
				dim remoteFile as string = "http://servicos.cptec.inpe.br/XML/cidade/1219/previsao.xml"
				try
					xml.Load(remoteFile)
					xml.Save(_lib.mapPath(fileWeather))
				catch ex as exception
					erro = true
				end try
			end if
			
			If not System.IO.File.Exists( _lib.mapPath(fileWeather) ) Then
				erro = true
			else
				html = ""
				try
					xml.Load( _lib.mapPath(fileWeather))
					m_nodelist = xml.SelectNodes("/cidade/previsao")
					if m_nodelist.Count > 0 then
						
						html = html & "<div id=""gnc_slider_weather"">"
							For Each m_node In m_nodelist
								data = m_node.ChildNodes.Item(0).InnerText
								dtDate = DateTime.Parse(data, Globalization.CultureInfo.CreateSpecificCulture("pt-BR"))
								data = dtDate.ToString("dd") & "/" & dtDate.ToString("MM")
								
								html = html & "<div class=""inner"">"
									html = html & "<p class=""txt-tempo"">Previsão do Tempo - dia " & data & "</p>"
									html = html & "<img src=""/assets/images/weather/cptec/" & m_node.ChildNodes.Item(1).InnerText & ".png"" alt=""" & m_node.ChildNodes.Item(1).InnerText & """ />"
									html = html & "<div class=""previsao"">"
										html = html & "<div class=""max""><div class=""ico_max""></div><div class=""txt_max"">Max: " & m_node.ChildNodes.Item(2).InnerText & "<span class=""txtgraus"">ºC</span></div> </div>"
										html = html & "<div class=""min""><div class=""ico_min""></div><div class=""txt_min"">Min: " & m_node.ChildNodes.Item(3).InnerText & "<span class=""txtgraus"">ºC</span></div> </div>"
									html = html & "</div>"
									html = html & "<p class=""weatherUpdate""><span class=""txt-update"">Fonte: Cptec</span></p>"
								html = html & "</div>"
							next
						html = html & "</div>"
						
						'html = html & "<div class=""nav"">"
							'html = html & "<div class=""pos-ant""><a id=""ant"" href=""#""><span>anterior</span></a></div>"
							'html = html & "<div class=""pos-prox""><a id=""prox"" href=""#""><span>próximo</span></a></div>"
						'html = html & "</div>"					
						
					end if
				catch ex as exception
					erro = true
				end try				
			end if
			
			if erro then html = html2
			
			return html
		end function		
			
	End Class
	
end Module

'***********************************************************************************************************
'* Previsao do Tempo
'***********************************************************************************************************
Public Class NPMarkupPage
    Inherits Page
    Public Overloads Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
    End Sub
End Class
