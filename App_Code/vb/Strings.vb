Imports Microsoft.VisualBasic
Imports System
Imports System.Configuration
Imports System.Text
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions

Public Class _netPortalStrings

	Public Sub New()
		MyBase.New()
	End Sub
	
	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************
	Public sub write(ByVal str as string)
		HttpContext.Current.response.write(str)
	End Sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	writes a string to the output in the same line
	'' @PARAM:			value [string]: output string
	'******************************************************************************************************************	
    Public Function clearSpace(ByVal html As String)
       	Dim strHtml As String = html
        strHtml = Regex.Replace(strHtml, "^\s+<", " <", RegexOptions.Multiline)
        strHtml = Regex.Replace(strHtml, ">\s+<", "> <", RegexOptions.Multiline)
        Return (strHtml.Trim)
    End Function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified stylesheet-file
	'******************************************************************************************************************
	Public function loadCSS(byVal url as String, byVal media as string)
		dim html as string = ""
			html = html & ("<link type=""text/css"" rel=""stylesheet"" href=""" & url & """" & IIF(media <> "", " media=""" & media & """", string.empty ) & " />")
		return html
	end Function		

	'******************************************************************************************************************
	'' @SDESCRIPTION:	Loads a specified stylesheet-file
	'******************************************************************************************************************
	Public function loadJS(byVal url as String)
		dim html as string
		html = html & ("<script type=""text/javascript"" src=""" & url & """></script>")
		return html
	end Function
	
	'***********************************************************************************************************
	'' @SDESCRIPTION:	executes a given javascript. input may be a string or an array. each field = a line
	'' @PARAM:			JSCode [string]. [array]: your javascript-code. e.g. <em>window.location.reload()</em>
	'***********************************************************************************************************
	Public function JS(byVal JSCode as String)
		dim html : html = ""
		
		html = html & ("<script type=""text/javascript"">")
			html = html & JSCode
		html = html & ("</script>")
		
		return html
	end function	

End Class