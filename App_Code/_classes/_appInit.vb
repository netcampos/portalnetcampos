Imports Microsoft.VisualBasic
Imports System

Public Class _appInit
	
	#Region "INIT"
		
		'public members
		Public controller as String
		Public fileTemplate as String
		
		'private members
		Private _lib 	as new _appLibrary()
		Private _user 	as new _appUsers()
		Private _db		as new _appDB()	
			
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:
		'******************************************************************************************************************
		Public Sub New()
			MyBase.New()
			
			appPageID 		= 0
			appModuleID 	= 0
			appContentID 	= 0
			
			_lib.defineHttpHeader()
			controller = String.Empty
			fileTemplate = String.Empty
			controller = _lib.pageController()
			me.checkController()
			
			
		End Sub
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:
		'******************************************************************************************************************
		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Public Function getHeader() as String
			dim html, seoTitle, seoDescription, seoKeywords as String

    		html = html & "<meta charset=""utf-8"">" & vbnewline
			
			if appPageID > 0 and appContentID = 0 then
				seoTitle 		= _db.getPage("seoTitle") & " | Portal NetCampos"
				seoDescription 	= _db.getPage("seoDescription")
				seoKeywords		= _db.getPage("seoKeywords")
			elseif appPageID = 0 and appContentID > 0 then
				
			end if
			
			if _lib.lenb(seoTitle) > 0 then html = html & "<title>" & seoTitle & "</title>" & vbnewline
			if _lib.lenb(seoDescription) > 0 then html = html & "<meta name=""description"" content=""" & seoDescription & """/>" & vbnewline
			if _lib.lenb(seoKeywords) > 0 then html = html & "<meta name=""keywords"" content=""" & seoKeywords & """/>" & vbnewline			
			
			html = html & "<!-- start: CSS -->" & vbnewline
    		html = html & "<link href=""" & _lib.domainAssets() & "/css/bootstrap.css,style.css"" rel=""stylesheet"" type=""text/css"" media=""all"" >" & vbnewline
    		html = html & "<link href=""//netdna.bootstrapcdn.com/font-awesome/3.2.1/css/font-awesome.min.css"" rel=""stylesheet"">" & vbnewline
    		html = html & "<!-- end: CSS -->" & vbnewline
			
			html = html & "<!-- start: viewport -->" & vbnewline
		   	html = html & "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & vbnewline
			html = html & "<!-- end: viewport -->" & vbnewline			
			
			return html
		End Function
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:
		'******************************************************************************************************************		
		Private Sub checkController()
			dim fileLoad as string = _lib.templatePath & controller & "/page-" & controller & ".ascx"
			dim controllerDB as String
			
			if (_lib.fileExists( fileLoad ) ) then
				fileTemplate = fileLoad
			else
				controllerDB = _db.getController(controller, "controller")
				
				if ( _lib.lenB(controllerDB) > 0 ) then
					fileTemplate 	= _lib.templatePath & controllerDB & "/default.ascx"
					appPageID 		= _db.getController(controller, "id")
					appModuleID 	= _db.getController(controller, "idModulo")
				end if
				
				if (not _lib.fileExists( fileTemplate ) ) then
					_lib.NotFound(fileTemplate)
				end if
				
			end if
				
		End Sub
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:
		'******************************************************************************************************************	
		Public Function loadPageContent() as String
			Dim html as string
			
			html = _lib.getTemplatePart(fileTemplate)
			
			
			return html				
		End Function
		
		
	#End Region
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:
	'******************************************************************************************************************		
	#Region "User Functions"
		
		'******************************************************************************************************************
		'' @SDESCRIPTION: Check if User is Connected
		'******************************************************************************************************************	
		Public Function userON() as Boolean			
			return _user.userON()				
		End Function
		
		'******************************************************************************************************************
		'' @SDESCRIPTION: Check total Members
		'******************************************************************************************************************	
		Public Function totalMembers() as Integer
		
			return 0
		End Function						
		
	#End Region
			
End Class