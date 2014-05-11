Imports System
Partial Class _Default
	Inherits System.Web.UI.Page
	
	Public pageController as String
	Public _lib 	as new _appLibrary()
	Public _user 	as new _appUsers()
	Public _db		as new _appDB()
	Public _app		as new _appInit()
	
	'******************************************************************************************************************
	'' @DESCRIPTION:
	'******************************************************************************************************************	
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
		_lib.setHttpHeader()
		pageController = _lib.pageController()		
		if (_lib.fileExists("/__webapp/templates/" & pageController & "/default.ascx")) then
			pageController = "/__webapp/templates/" & pageController & "/default.ascx"
		else
			
			pageController = "/__webapp/templates/" & pageController & "/default.ascx"
		end if
		
		_lib.write(pageController)
		
		'_lib.write( _lib.currentURL() )
		'_lib.write("<br>")
		'_lib.write( _lib.getURL(0) )		
		
		'pageController = _lib.getTemplatePart("/__webapp/templates/" & pageController & "/default.ascx")
		
		'_lib.GenerateControlMarkup("/App/views/sobre/" & pageView & ".ascx")
		
		'_lib.write("<br>")
		'_lib.write( pageController )		
		
    End Sub
		
End Class