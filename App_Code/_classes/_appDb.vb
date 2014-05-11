Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports MySql.Data
Imports MySql.Data.MySqlClient

Imports System.IO
Imports System.Xml

Public Class _appDB
	
	#Region "INIT"
		
		Private _lib as new _appLibrary()
		
		'******************************************************************************************************************
		'' @SDESCRIPTION:
		'******************************************************************************************************************
		Public Sub New()
			MyBase.New()
		End Sub
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub
		
		'Private Property
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Private Property sqlServer() As String
			Get
				Return System.Configuration.ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString
			End Get
			Set(ByVal value As String)
			
			End Set
		End Property
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Private Property mysql5() As String
			Get
				Return System.Configuration.ConfigurationManager.ConnectionStrings("mysql5").ConnectionString
			End Get
			Set(ByVal value As String)
			
			End Set
		End Property
		
		'***********************************************************************************************************
		'* description:
		'***********************************************************************************************************
		Private Property idCidade() As Integer
			Get
				Return _lib.QS("idCidade")
			End Get
			Set(ByVal value As Integer)
			
			End Set
		End Property
		
		'***********************************************************************************************************
		'* description: 
		'***********************************************************************************************************
		Private Property domainName() As String
			Get
				Return _lib.domainName()
			End Get
			Set(ByVal value As String)
			
			End Set
		End Property
		
	#End Region
	
	'*****************************************************************************************************************
	#Region "Info Tables"
	
		'******************************************************************************************************************
		'' @SDESCRIPTION: RETORNA INFORMACOES SOBRE A AREA EM QUESTAO
		'******************************************************************************************************************
		Public function getController(byval slug as string, byval coluna as string) as String
			Dim retorno as string
			Dim conn as SqlConnection 
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(sqlServer())
			
			if _lib.lenb(retorno) = 0 then
				try	
					conn.Open()
					cmd = New SqlCommand("select top 1 * from _cms_view_cidades_controllers where idCidade = @id and seoSlug = @slug", conn)
					with cmd
						.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
						.Parameters("@id").Value = idCidade()
						.Parameters.Add(New SqlParameter("@slug", SqlDbType.VarChar))
						.Parameters("@slug").Value = slug
					end with
					dr = cmd.ExecuteReader()
					if dr.HasRows then
						dr.Read()
						retorno = dr(coluna).toString()
					end if
					dr.close
				Catch ex As Exception
					retorno = ""
				Finally
					conn.Close()
					conn.Dispose()    
					If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
				End Try
			end if
			
			return retorno
		End Function
		
		'******************************************************************************************************************
		'' @SDESCRIPTION: RETORNA INFORMACOES SOBRE A AREA EM QUESTAO
		'******************************************************************************************************************
		Public function getPage(byval coluna as string) as String
			Dim retorno as string
			Dim conn as SqlConnection 
			Dim cmd as SqlCommand
			Dim dr As SqlDataReader
			conn = New SqlConnection(sqlServer())
			
			if _lib.lenb(appPageID) > 0 then
				try	
					conn.Open()
					cmd = New SqlCommand("select top 1 * from _cms_view_cidades_controllers where id = @id and idCidade = @idCidade", conn)
					with cmd
						.Parameters.Add(New SqlParameter("@id", SqlDbType.Int))
						.Parameters("@id").Value = appPageID
						.Parameters.Add(New SqlParameter("@idCidade", SqlDbType.Int))
						.Parameters("@idCidade").Value = idCidade()
					end with
					dr = cmd.ExecuteReader()
					if dr.HasRows then
						dr.Read()
						retorno = dr(coluna).toString()
					end if
					dr.close
				Catch ex As Exception
					retorno = ""
				Finally
					conn.Close()
					conn.Dispose()    
					If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
				End Try
			end if
			
			return retorno
		End Function		
	
	#End Region
			
End Class