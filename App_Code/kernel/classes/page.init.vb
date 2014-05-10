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
Public Class pageInit

	Public idModulo, idConteudo, idCategoria, restrito as Integer
	Public adArea as integer = 1
	Public active as integer = 0
	Public gncRestrito as boolean = false
	Public pageName, tableView, pageCSS as String
	Public seoTitle, seoDescription, seoKeywords, extraHeader as String
	
	Private idCidade as Integer = ConfigurationSettings.AppSettings("id_cidade")
	Private siteURL as string = "http://www.netcampos.com"
	Private str as new StringOperations()
	Private user as new npUsers()
	Private db as new database()
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public Sub New()
		MyBase.New()
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public sub setHttpHeader()
		
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
		Dim ts As New TimeSpan(0,60,0)
		HttpContext.Current.Response.Cache.SetMaxAge(ts)
		HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddSeconds(3600))
		HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
		HttpContext.Current.Response.Cache.SetValidUntilExpires(true)		
		HttpContext.Current.Response.ContentType = "text/html"
		HttpContext.Current.Response.Cache.VaryByHeaders("Accept-Language") = true
		HttpContext.Current.Response.Cache.VaryByHeaders("User-Agent") = true
		HttpContext.Current.Response.AppendHeader("X-Powered-By","Grupo NetCampos Tecnologia")
		HttpContext.Current.Response.Cache.SetLastModified(DateTime.Now.AddHours(-48))
		'HttpContext.Current.Response.Cache.SetETag( _lib.getMD5( DateTime.Now.toString() ) )

	End sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public Function drawHeader() as String
		dim html as string
			html = html & "<!DOCTYPE html>" & vbnewline
			html = html & "<html id=""netcampos"" lang=""pt-br"">"& vbnewline
			html = html & "<head>"& vbnewline
				
				html = html & "<meta charset=""utf-8"">"& vbnewline
				html = html & "<title>" & seoTitle & " | Portal NetCampos </title>"
				if _lib.lenB(seoDescription) > 0 then html = html & "<meta name=""description"" content=""" & seoDescription & """>"& vbnewline
				if _lib.lenB(seoKeyWords) > 0 then html = html & "<meta name=""keywords"" content=""" & seoKeyWords & """>"& vbnewline
				if active = 1 then 
					html = html & "<meta name=""robots"" content=""index,follow"">"& vbnewline
				else
					html = html & "<meta name=""robots"" content=""noindex,nofollow"">"& vbnewline
				end if
				html = html & extraHeader
				html = html & "<link href=""https://plus.google.com/100400043965768897941"" rel=""publisher"">"& vbnewline
				html = html & "<link rel=""icon"" type=""image/png"" href=""/assets/images/favicon.png"">"& vbnewline
				html = html & loadCSS()& vbnewline
				
				html = html & "<!--[if lt IE 8]> <script src=""http://html5shim.googlecode.com/svn/trunk/html5.js""></script><![endif]-->"& vbnewline
			html = html & "</head>"& vbnewline
			
			html = html & "<body id=""" & pageName & """ class=""gnc-conteudo"" itemscope itemtype=""http://schema.org/WebPage"">"& vbnewline
			
				html = html & "<div id=""fb-root""></div><div id=""gnc-googleplus""></div><div id=""gnc-twitter""></div>" & vbnewline
				if not user.userON() then
					if gncRestrito then html = html & "<div id=""gnc-restrito""></div>"& vbnewline
				end if
				html = html & "<div id=""wrapper-full"">"& vbnewline
				html = html & "<div id=""gnc-fixed-top"">" & vbnewline
					html = html & user.gncUserTopBar()
					html = html & contentHead()
				html = html & "</div>" & vbnewline
					
				html = html & "<div id=""wrapper"" class=""pagepub page-fixed"">"& vbnewline
				html = html & topADS(1,adArea)
				html = html & "<div id=""page"">"& vbnewline
					html = html & "<div id=""content"">" & vbnewline
					html = html & "<div id=""rbtop""><div class=""left""></div><div class=""right""></div></div>"& vbnewline
					html = html & "<div id=""bd"">"& vbnewline
		return html
	End Function
	
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
						html = html & "<a class=""marca_cidade"" href=""" & siteURL & """ title=""Voltar para a Pgina Inicial"">"& vbnewline
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
							html = html & "<span class=""anuncie""><a title=""Publicidade - Campos do Jordão"" href=""/publicidade/""><span>Publicidade - Campos do Jordão</span></a></span>"& vbnewline
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
	private Function SiteBar()
		dim html as string
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
		return (html)	
	end Function		
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Public Function drawFooter() as String
		dim html as string
			'footer area
					html = html & "</div>" 'end bd
					html = html & "<div id=""rbbot""><div class=""left""></div><div class=""right""></div></div>"
					html = html & "</div>" 'end bd
				html = html & SiteBar()
				html = html & "<div class=""clear""></div>"
				html = html & "</div>" 'end page
			html = html & "</div>" 'end wrapper
			html = html & "<div id=""footer"">"
				html = html & "<div class=""bg"">"
					html = html & "<div class=""sombraBotEsq""></div><div class=""sombraBotDir""></div>"
					html = html & navFooter()
				html = html & "</div>"
			html = html & "</div>"
			
			html = html & "</div>" & vbnewline 'end wrapper-full
				html = html & getJS()
			html = html & "</body>"& vbnewline
			html = html & "</html>"& vbnewline		
		return html
	End Function
	
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
	'* load css
	'***********************************************************************************************************	
	Private function loadCSS()
		dim html as string
			if _lib.lenb(pageCSS) > 0 then
				html = str.loadCSS("/assets/css/style.css," & pageCSS, string.empty)
			else
				html = str.loadCSS("/assets/css/style.css", string.empty)
			end if
		return html
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified stylesheet-file
	'******************************************************************************************************************
	Private function loadJS(byVal url as String)
		dim html as string
		html = html & ("<script type=""text/javascript"" src=""" & url & """></script>")
		return html
	end Function	
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************
	Private Function getJS()
		dim html as string
		html = html & loadJS("//ajax.googleapis.com/ajax/libs/jquery/1.7.1/jquery.min.js")
		html = html & "<script>window.jQuery || document.write('<script src=""/assets/js/jquery.min.js""><\/script>')</script>"
		
		html = html & loadJS("/assets/js/jquery.autosize.js,jquery.cookie.js,jquery.maskedinput.js,jquery.md5.js,jquery.simplemodal.js,jquery.carousel.js,gncWindow.js,submenu.js")
		html = html & loadJS("/assets/js/jquery.tipsy.js,jquery.gnc.user.js,jquery.gnc.page.js")
		html = html & loadJS("/assets/js/jquery.rating.js")
					
		select case pageName
			case "noticias"				
				html = html & "<script>$." & pageName & "();</script>"
			case "passeios"				
				html = html & "<script>$." & pageName & "();</script>"				
			case "fotos"				
				html = html & "<script>$." & pageName & "();</script>"
			case "contato"		
				html = html & loadJS("/assets/js/jquery.gnc.forms.js")		
				html = html & "<script>$." & pageName & "();</script>"
			case "institucionais"				
				html = html & "<script>$." & pageName & "();</script>"				
									
		end select
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
									html = html & "<div class=""max""><div class=""ico_max""></div><div class=""txt_max"">Max: " & m_node.ChildNodes.Item(2).InnerText & "<span class=""txtgraus"">C</span></div> </div>"
									html = html & "<div class=""min""><div class=""ico_min""></div><div class=""txt_min"">Min: " & m_node.ChildNodes.Item(3).InnerText & "<span class=""txtgraus"">C</span></div> </div>"
								html = html & "</div>"
								html = html & "<p class=""weatherUpdate""><span class=""txt-update"">Fonte: Cptec</span></p>"
							html = html & "</div>"
						next
					html = html & "</div>"
					
				end if
			catch ex as exception
				erro = true
			end try				
		end if
		
		if erro then html = html2
		
		return html
	end function		
	
End Class