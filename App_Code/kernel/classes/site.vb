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

Public Class Site
	
	Private idCidade as Integer = ConfigurationSettings.AppSettings("id_cidade")
	Private pageFBID as string = "108204652584452"
	Private str as new StringOperations()
	Private user as new npUsers()
	Private db as new database()

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
		HttpContext.Current.Response.Cache.SetLastModified(DateTime.Now.AddHours(-1))
	End sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
	Public function draw_home()
	
		if not user.userON then HttpContext.Current.Response.redirect("http://camposdojordao.netcampos.com")
	
		setHttpHeader()
		drawHeader_Home()
		'drawPage_Home()
		drawContentHome()
		drawFooter_Home()
	End Function
	
	'***********************************************************************************************************
	'* draw Header 
	'***********************************************************************************************************				
	Private sub drawHeader_Home()
		dim html as String = ""
		dim idmodulo as integer = 18
		
		html = html & "<!DOCTYPE html>" & vbnewline
		html = html & "<html id=""netcampos"" dir=""ltr"" lang=""pt-BR"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:og=""http://opengraphprotocol.org/schema/"" xmlns:fb=""http://developers.facebook.com/schema/"">"& vbnewline
		html = html & "<head profile=""http://gmpg.org/xfn/11"">"& vbnewline
			html = html & "<meta charset=""UTF-8"">"& vbnewline
			html = html & "<meta http-equiv=""X-UA-Compatible"" content=""chrome=1"" />"& vbnewline			
			html = html & "<title>Campos do Jordão, Portal NetCampos em Campos do Jordão SP</title>"& vbnewline
			'html = html & "<meta name=""description"" content=""Em Campos do Jordão, as melhores dicas de Hotéis, Pousadas, Restaurantes, Passeios de Campos do Jordão estão no Portal NetCampos. Venha para Campos do Jordao SP."" />"& vbnewline
			html = html & "<meta name=""description"" content=""Em Campos do Jordão, turismo, dicas de passeios em Campos do Jordão é no Portal NetCampos. Encontre Restaurantes, Pousadas, Hotéis em Campos do Jordao, confira."" />"& vbnewline
			'html = html & "<meta name=""keywords"" content="""&db.dataCidade("keywords")&""" />"& vbnewline
			html = html & "<meta name=""robots"" content=""index,follow"" />"& vbnewline
			html = html & "<meta name=""viewport"" content=""width=device-width"" />" & vbnewline
			html = html & "<meta name=""google-site-verification"" content=""pLyWQe6jno9lCKNYilnYGVSsignAgIHA7Hts26sdxsg"" />"& vbnewline
			html = html & "<meta name=""msvalidate.01"" content=""3E098835F693253313AA78E7EE64A827"" />"& vbnewline
			html = html & "<meta name=""alexaVerifyID"" content=""p2uW8xaWsRG8Gv8Lk4kMHAIswe8"" />" & vbnewline
			html = html & "<link rel=""canonical"" href=""http://www.netcampos.com/"" />"& vbnewline
			html = html & "<link rel=""alternate"" type=""application/rss+xml"" title=""RSS 2.0"" href=""http://www.netcampos.com/feed/"" />"& vbnewline
			html = html & "<link href=""https://plus.google.com/100400043965768897941"" rel=""publisher"" />"& vbnewline
			html = html & "<link href=""/assets/images/favicon.png"" type=""image/png"" rel=""icon"">" & vbnewline
			'facebook
			html = html & "<meta property=""og:locale"" content=""pt_BR"" />"& vbnewline
			html = html & "<meta property=""fb:app_id"" content=""136304184521"" />"& vbnewline
			html = html & "<meta property=""og:title"" content=""NetCampos | Quem Visita Campos do Jordão passa por aqui"" />"& vbnewline
			html = html & "<meta property=""og:type"" content=""website"" />"& vbnewline
			html = html & "<meta property=""og:url"" content=""http://www.netcampos.com/"" />"& vbnewline
			html = html & "<meta property=""og:image"" content=""http://www.netcampos.com/assets/images/netcampos-og.png"" />"& vbnewline
			html = html & "<meta property=""og:site_name"" content=""netcampos.com"" />"& vbnewline
			html = html & "<meta property=""og:description"" content=""Campos do Jordão visite o ano inteiro, as melhores dicas de Campos do Jordão estão no Portal NetCampos, acesse, compartilhe e participe."" />"& vbnewline
			html = html & loadCSS_home() & vbnewline
			html = html & loadJSHeader_home() & vbnewline
		html = html & "</head>"& vbnewline
		html = html & "<body id=""home"" class=""gnc-body"">"& vbnewline
		html = html & "<div id=""fb-root""></div>"& vbnewline
		html = html & "<div id=""wrapper-full"">"& vbnewline
			
			html = html & user.top_toolbar()
			html = html & "<div id=""head"">"& vbnewline
				html = html & "<div class=""header"">"& vbnewline
					html = html & "<div class=""sombratopesq""></div>"& vbnewline
					html = html & "<div class=""sombratopdir""></div>"& vbnewline
					
					'logo
					html = html & "<div id=""logoPortal"" class=""logo"">"& vbnewline
						html = html & "<a class=""marca_cidade"" href=""/"" title=""Portal NetCampos - Campos do Jordão"">"& vbnewline
							html = html & "<h1 class=""logotxt"">Campos do Jordão, saiba sobre a cidade no Portal NetCampos</h1>"& vbnewline
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
							html = html & "<h3 class=""anuncie""><a title=""Publicidade no Portal NetCampos"" href=""/publicidade/""><span>Publicidade no Portal NetCampos</span></a></h3>"& vbnewline
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
			html = html & "</div>"
		html = str.clearSpace(html)
		str.write(html)
	End Sub
	
	'***********************************************************************************************************
	'* draw Page Content 
	'***********************************************************************************************************
	private sub drawContentHome()
		dim html as String
		html = html & "<div id=""wrapper"" class=""pagepub"">"
			'SkyBanner
			html = html & topADS(1,1)
			'Page Content
			html = html & "<div id=""page"">"
				html = html & "<div id=""content"">"
					html = html & "<div id=""n-col1"">"
						html = html & "<div id=""homeCanais"">"
							html = html & "<div id=""adsCanais"">" & db.ads(16,1) & "</div>"
							html = html & "<div id=""topEsqMenu""><div class=""title titleCanais""><span>Canais NetCampos</span></div></div>"
							html = html & "<div class=""leftNavMenu"">" & db.verticalNavHome() & "</div>"
							'mural de recados
							html = html & WidGet_MuralRecados()
							'enquete
							html = html & WidGet_Enquete()
							'fbfan
							html = html & WidGet_FBFan()
							'cadastro
							html = html & WidGet_Cadastro()
							html = html & "<div id=""fechaMenu""></div>"
						html = html & "</div>"
					html = html & "</div>"
					
					html = html & "<div id=""n-col2"">"
						'tv destaque
						html = html & "<div id=""topDestaques"">"
							'destaques
							html = html & WidGet_DestaquesHome()
							'hotChannel
							html = html & WidGet_HotChannel()
						html = html & "</div>"
						
						'conteudo pagina
						html = html & "<div id=""contentArea"">"
							html = html & "<div id=""rbtop""><div class=""left""></div><div class=""right""></div></div>"
							'conteudo
							html = html & "<div class=""bd"">"
								
								html = html & "<div id=""n-cont0"">"
									html = html & "<div class=""n-col1"">"
										'Campos do Jordão, visite o ano todo.
										html = html & "<div class=""n-colTitle""><h2>Visite Campos do Jordão o ano todo</h2></div>"
										html = html & "<div class=""n-content"">"
											html = html & "<div class=""col-left"">"
												html = html & "<img width=""200"" height=""140"" src=""//assets.netcampos.com/images/home/campos-do-jordao.jpg"" alt=""Visite Campos do Jordão o ano todo"" alt=""Visite Campos do Jordão o ano todo"" />"
											html = html & "</div>"
											html = html & "<div class=""col-right"">"
												html = html & "<h3>Esqueça o estress e deixe sua imaginação voar.</h3>"
												html = html & "<p>Visitar Campos do Jordão é estar em contato com uma natureza exuberante e um clima reconhecido como um dos melhores do mundo. Campos do Jordão é para ser visitada o ano todo, a cidade possui excelente infra estrutura para satisfazer os mais exigentes visitantes.</p><a href=""/campos-do-jordao/"" title=""Continue Lendo"">continue lendo...</a>"
											html = html & "</div>"
										html = html & "</div>"
									html = html & "</div>"
								html = html & "</div>"
																	
								'noticias e passeios
								html = html & "<div id=""n-cont1"">"
									'ultimas noticias
									html = html & WidGet_UltimasNoticias()
									'dicas de passeios
									html = html & WidGet_PasseiosHome()
								html = html & "</div>"
								
								'galeria de fotos
								html = html & "<div id=""n-cont2"">"
									html = html & WidGet_FotosHome()
								html = html & "</div>"
								
								'videos e agenda
								html = html & "<div id=""n-cont3"">"
									'videos
										html = html & WidGet_InstitucionaisHome()
									'agenda
										html = html & WidGet_AgendaHome()
								html = html & "</div>"
								
								'acidade
								'html = html & "<div id=""n-cont4"">"
									'html = html & WidGet_InstitucionaisHome()
								'html = html & "</div>"
								
							html = html & "</div>"
							
							html = html & "<div id=""rbbot""><div class=""left""></div><div class=""right""></div></div>"
						html = html & "</div>"
						
					html = html & "</div>"
					
				html = html & "</div>"
				
				html = html & "<div id=""sidebar"">"
					html = html & widget_topBanner(13,1)
				html = html & "</div>"
				
				html = html & "<div class=""clear""></div>"
				
			html = html & "</div>"
		html = html & "</div>"
		html = str.clearSpace(html)		
		str.write(html)
	end sub
		
	'***********************************************************************************************************
	'* draw footer Home 
	'***********************************************************************************************************
	Private sub drawFooter_Home()
		dim html as String
			html = html & "<div id=""footer"">"
				html = html & "<div class=""bg"">"
					html = html & "<div class=""sombraBotEsq""></div><div class=""sombraBotDir""></div>"
					html = html & navFooter()
				html = html & "</div>"
			html = html & "</div>"	
		html = html & "</div>"
		html = html & str.loadJS("/assets/js/jquery.tools.min.js,jquery.wslide.js,jquery.timers.js,jquery.jcarousel.js,jquery.jFav-1.0.js") & vbnewline
		html = html & str.loadJS("/assets/js/submenu.js,nyroModal.min.js,qtip.js,cycle.min.js,gnc.home.js") & vbnewline		
		html = html & footerResources()
		html = html & "</body>"
		html = html & "</html>"
		html = str.clearSpace(html)
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
						html = html & "<li><a href=""/campos-do-jordao/"" rel=""nofollow"">A Cidade</a></li>"
						html = html & "<li><a href=""/cadastro/"" rel=""nofollow"">Cadastro</a></li>"
						html = html & "<li><a href=""/fotos-campos-do-jordao/"" rel=""nofollow"">Fotos</a></li>"
						html = html & "<li><a href=""/noticias-campos-do-jordao/"" rel=""nofollow"">Notícias</a></li>"
						html = html & "<li><a href=""/passeios-campos-do-jordao/"" rel=""nofollow"">Passeios</a></li>"
						html = html & "<li><a href=""/empresas/"" rel=""nofollow"">Empresas</a></li>"
						html = html & "<li><a href=""/empresas/restaurantes/"" rel=""nofollow"">Restaurantes</a></li>"
						html = html & "<li><a href=""/campos-do-jordao/dicas-de-hospedagem.html"" rel=""nofollow"">Hospedagem</a></li>"
						html = html & "<li><a href=""/empresas/utilidade-publica/"" rel=""nofollow"">Utilidade Pública</a></li>"
						html = html & "<li><a href=""/campos-do-jordao/prefeitura-municipal.html"" rel=""nofollow"">Prefeitura</a></li>"
						html = html & "<li class=""final_lista""><a href=""/busca/"" rel=""nofollow"">Buscador</a></li>"
					html = html & "</ul>"
				html = html & "</div>"
				
				'secondary
				html = html & "<div class=""foot-secondary"">"
					html = html & "<ul>"
						html = html & "<li><a href=""/parceiros/"" rel=""nofollow"">Parceiros</a></li>"
						html = html & "<li><a href=""/press-release/expediente.html"" rel=""nofollow"">Expediente</a></li>"
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
	'* draw footer Home 
	'***********************************************************************************************************
	Private Function footerResources() as string
		dim html as string
			html = html & str.loadJS("/assets/js/crossbrowser.js") & vbnewline
			html = html & city.dados(idCidade, "estatisticas")			
		return html
	End Function
		
	'***********************************************************************************************************
	'* draw Header 
	'***********************************************************************************************************	
	Private function loadCSS_home()
		dim html as string
		'html = str.loadCSS("/assets/css/site.css,home.css", string.empty)
		'html = str.loadCSS("//assets2.netcampos.com/css/netcampos.css", string.empty)
		html = html & str.loadCSS("//assets.netcampos.com/css/site.css", string.empty)
		html = html & str.loadCSS("//assets.netcampos.com/css/home.css", string.empty)
		return html
	end function
	
	'***********************************************************************************************************
	'* draw Header 
	'***********************************************************************************************************	
	Private function loadJSHeader_home()
		dim html as string
		'html = str.loadJS("/assets/js/jquery.min.js")
		'html = str.loadJS("//assets2.netcampos.com/js/jquery.min.js")
		html = str.loadJS("//assets.netcampos.com/js/jquery.min.js")
		return html
	end function
	
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
	private function widget_topBanner(byval qtd as integer, byval local as integer) as string	
		dim html as string
		html = html & "<div id=""wrapper-col-pub"">"
			html = html & "<div class=""title""><h3>publicidade</h3></div>"
			html = html & "<div class=""content"">"
				html = html & asideADS(qtd,local)
			html = html & "</div>"
			html = html & "<div class=""bot""></div>"
			html = html & "<div id=""col-partners"">"
				html = html & "<div class=""partners"">"
					html = html & "<div class=""marca_brasil""></div>"
					html = html & "<div class=""brasil_2014""></div>"
				html = html & "</div>"
			html = html & "</div>"
		html = html & "</div>"
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet Mural de Recados
	'***********************************************************************************************************
	Private Function WidGet_MuralRecados()
		dim html as string
		html = html & "<div id=""SideMural"" class=""boxCanais"">"
			html = html & "<div class=""title titleMural""><span>Mural de Recados</span></div>"
			html = html & "<div class=""MuralHome"">"
			html = html & "</div>"
		html = html & "</div>"
		return html
	End function
	
	'***********************************************************************************************************
	'* WidGet Enquete
	'***********************************************************************************************************
	Private Function WidGet_Enquete()
		dim html as string
		html = html & "<div class=""boxCanais"">"
			html = html & "<div class=""title titleEnquete""><span>Enquete</span></div>"
			html = html & "<div class=""enquete"">"
				html = html & "<h3>" & db.tituloEnquete() & "</h3>"
				html = html & "<div class=""perguntas"">"
					html = html & db.perguntasEnquete()
				html = html & "</div>"
			html = html & "</div>"
		html = html & "</div>"
		return html
	End Function
	
	'***********************************************************************************************************
	'* WidGet facebook
	'***********************************************************************************************************
	Private Function WidGet_FBFan()
		dim html as string
		html = html & "<div class=""boxCanais"">"
			html = html & "<div id=""ads165x320"">"
				html = html & "<div id=""facebook_fan""></div>"
			html = html & "</div>"
		html = html & "</div>"
		return html
	End Function
	
	'***********************************************************************************************************
	'* WidGet facebook
	'***********************************************************************************************************
	Private Function WidGet_Cadastro()
		dim html as string
		html = html & "<div class=""boxCanais"">"
			html = html & "<div class=""title titleCadastro""><span>Cadastro Gratuíto</span></div>"
			html = html & "<div class=""fCadastro"">"
				html = html & "<h5 class=""title"">Faça seu cadastro gratuito para participar das nossas inumeras promoções e ter acesso as áreas exclusívas do portal.</h5>"
				html = html & "<form action=""/cadastro/"" method=""post"" name=""frmcadastro"" id=""frmcadastro"">"
					html = html & "<fieldset>"
						html = html & "<input type=""text"" class=""input-text"" value=""nome"" name=""nome"" onFocus=""(this.value = '');"" />"
						html = html & "<input type=""text"" class=""input-text"" value=""email"" name=""email"" onFocus=""(this.value = '');"" />"
						html = html & "<input type=""submit"" name=""btcadastro"" class=""btcadastro"" id=""btcadastro"" value=""cadastrar"" />"
					html = html & "</fieldset>"
				html = html & "</form>"
			html = html & "</div>"
		html = html & "</div>"
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet destaques home
	'***********************************************************************************************************
	Private Function WidGet_DestaquesHome()
		dim html as string
		html= html & "<div id=""playerTV"">"
			html = html & _lib.drawSWF("http://assets.netcampos.com/swf/playerNetCampos.swf",353,255,"url=http://www.netcampos.com/App_Modules/slideShow/playerHome.aspx?idCidade=" & idCidade & "&amp;corFundo=0xB0D638")
		html = html & "</div>"
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet hotchannel home
	'***********************************************************************************************************
	Private Function WidGet_HotChannel()
		dim html as string
		html = html & "<div id=""hotChannel"">"
			html = html & "<div class=""head"">"
				html = html & "<div class=""left""></div>"
				html = html & "<div class=""center"">"
					
					html = html & "<div class=""addFavorito"">"
						html = html & "<span class=""ico""><a href=""http://www.netcampos.com/"" class=""addBookmark"" title=""Portal NetCampos""></a></span>"
						html = html & "<p><a href=""http://www.netcampos.com/"" class=""addBookmark"" title=""Portal NetCampos"">Adicione o Portal aos Seus Favoritos</a></p>"
					html = html & "</div>"
					
					html = html & "<div class=""line""></div>"
					
					html = html & "<div class=""divulgueGratis"">"
						html = html & "<span class=""ico""><a href=""#""></a></span>"
						html = html & "<p><a href=""/publicidade/"">Anuncie sua empresa aqui no Portal</a></p>"
					html = html & "</div>"
					
					html = html & "<div class=""dropEmpresas"">"
						html = html & "<span class=""ico""><a href=""#""></a></span>"
						html = html & "<form id=""frmCatEmpresas"">"
						html = html & "<fieldset>"
						  html = html & "<div id=""catEmpresas"">"
							html = html & "<select name=""catEmpresas"" class=""listcatempresas"">"
							  html = html & "<option>Categorias</option>"
							html = html & "</select>"
						  html = html & "</div>"
						html = html & "</fieldset>"
						html = html & "</form>"
					html = html & "</div>"
					
				html = html & "</div>"
				html = html & "<div class=""right""></div>"
			html = html & "</div>"
			
			'advertising
			html = html & "<div class=""bd"">"
				html = html & "<div class=""left""></div>"
				html = html & "<div class=""center"">"
					html = html & "<div id=""HotMidia"">"& db.ads(17,1) &"</div>"
				html = html & "</div>"
				html = html & "<div class=""right""></div>"
			html = html & "</div>"									
			
		html = html & "</div>"
		
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet ultimas noticias home
	'***********************************************************************************************************
	Private Function WidGet_UltimasNoticias()
		dim html as string
		html = html & "<div class=""n-col1"">"
			html = html & "<div class=""n-colTitle""><h2><a href=""http://www.netcampos.com/noticias-campos-do-jordao/"" style=""color:#fff"">Notícias de Campos do Jordão</a></h2></div>"
			html = html & "<div class=""n-content"">"
				html = html & db.noticiasHome(6)
			html = html & "</div>"
			
			html = html & "<div class=""col-ads"">"
				html = html & "<p>publicidade</p>"
				html = html & "<div class=""ads"">" & db.ads(19,1) & "</div>"
			html = html & "</div>"
			
		html = html & "</div>"
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet passeios home
	'***********************************************************************************************************
	private function WidGet_PasseiosHome()
		dim html as string
		html = html & "<div class=""n-col2"">"
			html = html & "<div class=""n-colTitle""><h2><a href=""http://www.netcampos.com/passeios-campos-do-jordao/"" style=""color:#fff"">Passeios em Campos do Jordão</a></h2></div>"
			html = html & "<div class=""n-content"">"
				html = html & db.passeiosHome(6)
			html = html & "</div>"
		html = html & "</div>"
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet fotos home
	'***********************************************************************************************************
	private function WidGet_FotosHome()
		dim html as string
		'html = html & "<div id=""n-cont2"">"
			html = html & "<div class=""n-col1"">"
				html = html & "<div class=""n-colTitle""><h2>Fotos de Campos do Jordão</h2></div>"
				html = html & "<div class=""n-content"">"
					html = html & db.fotosHome(6)
				html = html & "</div>"
			html = html & "</div>"
		'html = html & "</div>"
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet institucionais home
	'***********************************************************************************************************
	private function WidGet_InstitucionaisHome()
		dim html as string
		html = html & "<div class=""n-col1 gnc-institucionais"">"
			html = html & "<div class=""n-colTitle""><h2>Sobre Campos do Jordão</h2></div>"
			html = html & "<div class=""n-content"">"
				html = html & db.institucionais(2)
			html = html & "</div>"
			html = html & "<div class=""n-footer"">"
				html = html & "<div class=""n-link""><a href=""/campos-do-jordao/"">dados gerais</a></div>"
			html = html & "</div>"
		html = html & "</div>"
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet Videos home
	'***********************************************************************************************************
	private function WidGet_VideosHome()
		dim html as string
		html = html & "<div class=""n-col1"">"
			html = html & "<div class=""n-colTitle""><h2>Vídeos da Cidade</h2></div>"
			html = html & "<div class=""n-content"">"
				html = html & "<div id=""mediaplayer"">"
					html = html & _lib.playerSWF("/App_Modules/player/netplayer.swf",300,260,"config=/App_Modules/player/confighome.aspx?idCidade=" & idCidade & "", "VideoPlayer")
				html = html & "</div>"
			html = html & "</div>"
		html = html & "</div>"
		return html
	end function
	
	'***********************************************************************************************************
	'* WidGet Videos home
	'***********************************************************************************************************
	private function WidGet_AgendaHome()
		dim html as string	
		html = html & "<div class=""n-col2"">"
			
			html = html & "<div class=""n-colTitle""><h2>Agenda</h2></div>"
			html = html & "<div class=""n-content"">"
				html = html & "<div id=""agenda"">"
					html = html & "<div class=""n-list"">"
						html = html & db.agendaHome(10)
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
			
		html = html & "</div>"
		return html		
	end function
	'***********************************************************************************************************
	'* draw footer Home 
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
	'* draw Page Content 
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
								html = html & "<img src=""//assets.netcampos.com/images/weather/cptec/" & m_node.ChildNodes.Item(1).InnerText & ".png"" alt=""" & m_node.ChildNodes.Item(1).InnerText & """ />"
								html = html & "<div class=""previsao"">"
									html = html & "<div class=""max""><div class=""ico_max""></div><div class=""txt_max"">Max: " & m_node.ChildNodes.Item(2).InnerText & "<span class=""txtgraus"">ºC</span></div> </div>"
									html = html & "<div class=""min""><div class=""ico_min""></div><div class=""txt_min"">Min: " & m_node.ChildNodes.Item(3).InnerText & "<span class=""txtgraus"">ºC</span></div> </div>"
								html = html & "</div>"
								html = html & "<p class=""weatherUpdate""><span class=""txt-update"">Fonte: Cptec</span></p>"
							html = html & "</div>"
						next
					html = html & "</div>"
					
					html = html & "<div class=""nav"">"
						html = html & "<div class=""pos-ant""><a id=""ant"" href=""#""><span>anterior</span></a></div>"
						html = html & "<div class=""pos-prox""><a id=""prox"" href=""#""><span>próximo</span></a></div>"
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