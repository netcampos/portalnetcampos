Imports Microsoft.VisualBasic
Imports System

Public Module _kernelNetPortal
	
	Public device, appController, appView, urlBase, domainName as String
	Public webSite as String = _lib2.QS("host")	
	Public idList As New ArrayList()
			
	Public Class _app
		
		Public seoTitle, seoDescription as String		
		Private str as new _netPortalStrings()
		Private user as new _npUsers2()
		Private db as new _database()
				
		'******************************************************************************************************************
		'' @DESCRIPTION:
		'******************************************************************************************************************
		Public Sub New()			
			MyBase.New()
			seoTitle = string.empty
			seoDescription = string.empty
			appController = "/_app_Content/" & _lib2.device() & "/"
			appView = "/_app_Content/" & _lib2.device() & "/"
			
			appController = "/_app_Content/desktop/"
			appView = "/_app_Content/desktop/"			
			
			if webSite = "marina.gruponetcampos.corp" or webSite = "10.0.0.2"  then webSite = "http://www.netcampos.com"
			webSite = "http://www.netcampos.com"
		End Sub
		
		'******************************************************************************************************************
		'' @DESCRIPTION:
		'******************************************************************************************************************
		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub
		
		'******************************************************************************************************************
		'' @DESCRIPTION:
		'******************************************************************************************************************		
		Public Sub appVersion()
		
		End Sub
		
		'******************************************************************************************************************
		'' @DESCRIPTION:
		'******************************************************************************************************************
		Public sub init()
			device = _lib2.device()			
			try
				if _lib2.lenb(_lib2.getController() ) > 0 then appController = appController & _lib2.getController() & "/default.aspx"
				if _lib2.fileExists(appController) then
					_lib2.PageExecute(appController)
				else
					_lib2.fileNotFound(appController)
				end if
			catch ex as exception
				_lib2.write(ex.message())			
			end try
 		End sub
		
		'******************************************************************************************************************
		'' @DESCRIPTION:
		'******************************************************************************************************************
		Public sub setHttpHeader()
			HttpContext.Current.Response.Clear()
			HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8
			Dim ts As New TimeSpan(0,60,0)
			HttpContext.Current.Response.Cache.SetMaxAge(ts)
			HttpContext.Current.Response.Cache.SetExpires(DateTime.Now.AddSeconds(3600))
			HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
			HttpContext.Current.Response.Cache.SetValidUntilExpires(true)
			HttpContext.Current.Response.Cache.VaryByHeaders("Accept-Language") = true
			HttpContext.Current.Response.Cache.VaryByHeaders("User-Agent") = true
			HttpContext.Current.Response.Cache.SetLastModified(DateTime.Now.AddHours(-2))
			HttpContext.Current.Response.AppendHeader("Server", "NetCampos/1.0")
			HttpContext.Current.Response.ContentType = "text/html"
		End sub
		
		'******************************************************************************************************************
		'' @DESCRIPTION:
		'******************************************************************************************************************
		Public Function get_header(optional byval metaHeader as string = "")
			Dim html as String = ""
			
			if _lib2.lenb(metaHeader) > 0 then
				html = metaHeader
			else
				html = html & "<!DOCTYPE html>" & vbnewline
				html = html & "<!--[if lt IE 7]>      <html lang=""pt-BR"" class=""no-js lt-ie9 lt-ie8 lt-ie7""> <![endif]-->" & vbnewline
				html = html & "<!--[if IE 7]>         <html lang=""pt-BR"" class=""no-js lt-ie9 lt-ie8""> <![endif]-->" & vbnewline
				html = html & "<!--[if IE 8]>         <html lang=""pt-BR"" class=""no-js lt-ie9""> <![endif]-->" & vbnewline
				html = html & "<!--[if gt IE 8]><!--> <html lang=""pt-BR"" class=""no-js""> <!--<![endif]-->" & vbnewline
				html = html & "<head>"& vbnewline
					html = html & "<meta charset=""utf-8"">" & vbnewline
					html = html & "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge,chrome=1"">" & vbnewline
					html = html & "<title>" & seoTitle & "</title>" & vbnewline
					html = html & "<meta name=""description"" content=""" & seoDescription & """>" & vbnewline
					html = html & "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbnewline
					html = html & "<link href=""/assets/css/bootstrap.css,style.css"" rel=""stylesheet"" media=""screen"">"
				html = html & "</head>"& vbnewline
					html = html & "<body id=" & _lib2.getController() & "  class=""gnc-body"">"& vbnewline
					html = html & "<header>"
						'fixed header
						html = html & "<div id=""fixedHeader"">"
							html = html & "<div id=""gncNavBar"" class=""navbar-fixed-top"">"
								html = html & "<div class=""container"">"
									if user.userON() then
										html = html & "<div class=""pull-left"">"
										
										html = html & "</div>"
										
										html = html & "<div id=""npTopBarRight"" class=""pull-right"">"
										
										html = html & "</div>"									
									else
										html = html & "<div class=""pull-left"">"
											html = html & "<ul>"
												html = html & "<li class=""icoHome""></li>"
												html = html & "<li class=""TotalMembers""><strong>" & user.countUsers() & "</strong> pessoas já estão conectadas à <strong>Campos do Jordão!</strong> Participe você também.</li>"
											html = html & "</ul>"											
										html = html & "</div>"
										
										html = html & "<div class=""pull-right"">"
											html = html & "<ul>"
												html = html & "<li><a href=""#""><i class=""icon icon-off""></i><span class=""text"">Entrar</span></a></li>"
												html = html & "<li><a href=""#""><i class=""icon icon-user""></i><span class=""text"">Cadastre-se</span></a></li>"
											html = html & "</ul>"
										html = html & "</div>"
									end if

								html = html & "</div>"
							html = html & "</div>"
						html = html & "</div>"
						
						'header page
						html = html & "<div id=""gnc-header"" class=""clearfix"">"
							html = html & "<div class=""container"">"
								html = html & "<div class=""content"">"
									html = html & "<div id=""logoPortal"" class=""logo"">"
										html = html & "<a href=""" & webSite & """ title=""Voltar para a Página Inicial""><div class=""logotxt""><span>Campos do Jordão - Campos do Jordão com NetCampos</span></div></a>"
									html = html & "</div>"
									
									html = html & "<div id=""gncTopAdvertising"" class=""pull-left"">"
										html = html & "<div class=""sky"">"
											html = html & "<div class=""sky-bg-left""></div>"
												html = html & "<div class=""sky-banner"">"
													html = html & "<div id=""super_banner"">"
														html = html & db.viewAds(1,1)
													html = html & "</div>"									
												html = html & "</div>"
											html = html & "<div class=""sky-bg-right""></div>"
										html = html & "</div>"
										
										html = html & "<div class=""sky2"">"
											html = html & "<div class=""sky-bg-left""></div>"
												html = html & "<div class=""sky-banner"">"
													html = html & "<a href=""#""><img src=""/assets/img/banners/Skybanner-2.gif"" /></a>"
												html = html & "</div>"
											html = html & "<div class=""sky-bg-right""></div>"
										html = html & "</div>"										
										
									html = html & "</div>"
																		
								html = html & "</div>"							
								
							html = html & "</div>"
							
								html = html & "<div id=""gncTopNav"">"
									html = html & "<nav class=""menu-primary-container"">"
										html = html & "<div class=""container"">"
										html = html & "<ul id=""NavMenu"">"
											html = html & "<li class=""first""><a href=""#"">Home</a></li>"
											html = html & "<li class=""first""><a href=""#"">A Cidade</a></li>"
											html = html & "<li><a href=""#"">Hotéis</a></li>"
											html = html & "<li><a href=""#"">Pousadas</a></li>"
											html = html & "<li><a href=""#"">Restaurantes</a></li>"
											html = html & "<li><a href=""#"">Imobiliárias</a></li>"
											html = html & "<li><a href=""#"">Empresas</a></li>"										
											html = html & "<li><a href=""#"">fotos</a></li>"
											html = html & "<li><a href=""#"">passeios</a></li>"
											html = html & "<li class=""last""><a href=""#"">notícias</a></li>"
											'html = html & "<li class=""last""><a href=""#"">Contato</a></li>"
										html = html & "</ul>"
										
										html = html & "<div id=""topSearch"">"
										
											html = html & "<form id=""busca"" name=""busca"" action=""/busca/"">"
												html = html & "<fieldset>"
													html = html & "<input type=""hidden"" name=""hl"" id=""hl"" value=""pt-BR"">" 
													html = html & "<input type=""text"" name=""q"" class=""input-search sprite-portal"" id=""q"" value="""">" 
													html = html & "<input type=""submit"" value=""Buscar"" class=""btn-search"">"
												html = html & "</fieldset>"
											html = html & "</form>"
										
										html = html & "</div>"
										html = html & "</div>"
									html = html & "</nav>"
								html = html & "</div>"								
							
						html = html & "</div>"
						
					html = html & "</header>"
					
					
					
					html = html & "<div id=""gncPage"">"

					
			end if
			
			return html
		End Function
		
		'******************************************************************************************************************
		'' @DESCRIPTION:
		'******************************************************************************************************************
		Public Function get_footer(optional byval metaHeader as string = "")
			Dim html as String = ""
			
				html = html & "</div>"
				html = html & "<div id=""gncFooter"">"
				
				html = html & "</div>"
				
				html = html & _lib2.getJS( _lib2.getController() )
				
			html = html & "</body>" & vbnewline
			html = html & "</html>"& vbnewline
			
			return html
		End Function
				
	
	End Class

end Module