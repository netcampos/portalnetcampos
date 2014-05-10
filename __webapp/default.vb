Imports System
Partial Class _Default
	Inherits System.Web.UI.Page
	
	Public html as String
	Public _lib as new _appLibrary()
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init	

		'_lib.write( _lib.currentURL() )
		'_lib.write("<br>")
		'_lib.write( _lib.getURL(0) )
		
		html = _lib.getTemplatePart("/__webapp/templates/passeios.ascx")
		
		'_lib.GenerateControlMarkup("/App/views/sobre/" & pageView & ".ascx")
		
		'_lib.write("<br>")
		'_lib.write( _lib.pageController() )		
		
    End Sub
		
End Class