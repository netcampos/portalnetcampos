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

Public Class _MemberShipProvider
    Inherits MembershipProvider

    Public Overrides Property ApplicationName() As String
        Get

        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Overrides Function ChangePassword(ByVal username As String, ByVal oldPassword As String, ByVal newPassword As String) As Boolean

    End Function

    Public Overrides Function ChangePasswordQuestionAndAnswer(ByVal username As String, ByVal password As String, ByVal newPasswordQuestion As String, ByVal newPasswordAnswer As String) As Boolean

    End Function

    Public Overrides Function CreateUser(ByVal username As String, ByVal password As String, ByVal email As String, ByVal passwordQuestion As String, ByVal passwordAnswer As String, ByVal isApproved As Boolean, ByVal providerUserKey As Object, ByRef status As System.Web.Security.MembershipCreateStatus) As System.Web.Security.MembershipUser

    End Function

    Public Overrides Function DeleteUser(ByVal username As String, ByVal deleteAllRelatedData As Boolean) As Boolean

    End Function

    Public Overrides ReadOnly Property EnablePasswordReset() As Boolean
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property EnablePasswordRetrieval() As Boolean
        Get

        End Get
    End Property

    Public Overrides Function FindUsersByEmail(ByVal emailToMatch As String, ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection

    End Function

    Public Overrides Function FindUsersByName(ByVal usernameToMatch As String, ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection

    End Function

    Public Overrides Function GetAllUsers(ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection

    End Function

    Public Overrides Function GetNumberOfUsersOnline() As Integer
		Dim total as integer
		Dim conn as SqlConnection
		Dim cmd as SqlCommand
		conn =  New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		try	
			dim onlineSpan as TimeSpan = new TimeSpan(0, System.Web.Security.Membership.UserIsOnlineTimeWindow, 0)
			dim compareTime as DateTime = DateTime.Now.Subtract(onlineSpan)		
			
			conn.Open()
			cmd = New SqlCommand("select count(*) from tbl_login_clientes_netportal where dataUltimaAtividade > @lastUpdate", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@lastUpdate", SqlDbType.DateTime))
				.Parameters("@lastUpdate").Value = compareTime
			end with
			total = cmd.ExecuteScalar()
		Catch ex As Exception
			total = 0
		Finally
			conn.Close()
			conn.Dispose()
		End Try
		
		return total
    End Function

    Public Overrides Function GetPassword(ByVal username As String, ByVal answer As String) As String

    End Function

    Public Overrides Function GetUser(ByVal providerUserKey As Object, ByVal userIsOnline As Boolean) As System.Web.Security.MembershipUser

			
    End Function

    Public Overrides Function GetUser(ByVal username As String, ByVal userIsOnline As Boolean) As System.Web.Security.MembershipUser
		
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
		
		cmd = New SqlCommand("select id, idcidade, idestado, usuario, senha, ativo, md5, data_cadastro, dataUltimaAtividade, dataUltimoLogin, dataAlteraSenha from tbl_login_clientes_netportal where usuario = @usuario", conn)
		with cmd
			.Parameters.Add(New SqlParameter("@usuario", SqlDbType.VarChar))
			.Parameters("@usuario").Value = username
		end with		
		
		Dim u As _MembershipUser = Nothing
		Dim dr As SqlDataReader = Nothing		
			
		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				u = GetUserFromReader(dr)
				If userIsOnline Then
					updateLastActivate(username)
				end if 
			end if
		Catch ex As Exception
			HttpContext.Current.Response.Write(ex.StackTrace.tostring())
		Finally
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()		
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
		End Try	
		
		return u			
			
    End Function
	
	Private Function GetUserFromReader(ByVal dr As SqlDataReader) As _MembershipUser
		Dim providerUserKey As Object = dr.GetValue(6)
		Dim username As String = dr.GetString(3)
		Dim email As String = dr.GetString(3)
		Dim passwordQuestion As String = ""
		Dim comment As String = ""
		Dim isApproved As Boolean = true
		Dim isLockedOut As Boolean = false
		Dim creationDate As DateTime = New DateTime()	
		Dim lastLoginDate As DateTime = New DateTime()
		Dim lastActivityDate As DateTime = New DateTime()
		Dim lastPasswordChangedDate As DateTime = New DateTime()
		Dim lastLockedOutDate As DateTime = New DateTime()
		Dim isSubscriber As Boolean = False
		Dim customerID As String = String.Empty

		Dim u As MembershipUser = New _MembershipUser(Me.Name, username, providerUserKey, email, passwordQuestion, comment, isApproved, isLockedOut, creationDate, lastLoginDate, lastActivityDate, lastPasswordChangedDate, lastLockedOutDate, isSubscriber, customerID)	

		return u
	
	End function
	
	private sub updateLastActivate(ByVal username as String)
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)
			
		If conn.State <> ConnectionState.Open Then conn.Open()
		cmd = new sqlCommand("update tbl_login_clientes_netportal set dataUltimaAtividade = @ultimaAtividade where usuario = @usuario", conn)
		with cmd
			.Parameters.Add(New SqlParameter("@ultimaAtividade", SqlDbType.DateTime))
			.Parameters("@ultimaAtividade").Value = DateTime.Now
			.Parameters.Add(New SqlParameter("@usuario", SqlDbType.VarChar))
			.Parameters("@usuario").Value = username
			.executeNonQuery()
		end with
		If conn IsNot Nothing Then conn.Dispose()  
		If conn IsNot Nothing Then conn.Dispose()		
		If conn.State = ConnectionState.Open Then conn.Close()
	
	end sub

    Public Overrides Function GetUserNameByEmail(ByVal email As String) As String

    End Function

    Public Overrides ReadOnly Property MaxInvalidPasswordAttempts() As Integer
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property MinRequiredNonAlphanumericCharacters() As Integer
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property MinRequiredPasswordLength() As Integer
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property PasswordAttemptWindow() As Integer
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property PasswordFormat() As System.Web.Security.MembershipPasswordFormat
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property PasswordStrengthRegularExpression() As String
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property RequiresQuestionAndAnswer() As Boolean
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property RequiresUniqueEmail() As Boolean
        Get

        End Get
    End Property

    Public Overrides Function ResetPassword(ByVal username As String, ByVal answer As String) As String

    End Function

    Public Overrides Function UnlockUser(ByVal userName As String) As Boolean

    End Function

    Public Overrides Sub UpdateUser(ByVal user As System.Web.Security.MembershipUser)

    End Sub

    Public Overrides Function ValidateUser(ByVal username As String, ByVal password As String) As Boolean

		Dim boolReturn As Boolean = False
		Dim conn as SqlConnection 
		Dim cmd as SqlCommand
		Dim dr As SqlDataReader
		conn = New SqlConnection(ConfigurationManager.ConnectionStrings("sqlServer").ConnectionString)	

		try	
			If conn.State <> ConnectionState.Open Then conn.Open()
			cmd = New SqlCommand("select md5 from tbl_login_clientes_netportal where usuario = @usuario and senha = @senha", conn)
			with cmd
				.Parameters.Add(New SqlParameter("@usuario", SqlDbType.VarChar))
				.Parameters("@usuario").Value = username
				.Parameters.Add(New SqlParameter("@senha", SqlDbType.VarChar))
				.Parameters("@senha").Value = Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(password, "md5")
			end with
			dr = cmd.ExecuteReader()
			if dr.HasRows then
				dr.Read()
				boolReturn = True
			end if
			dr.close
		Catch ex As Exception

		Finally
			If conn.State = ConnectionState.Open Then conn.Close()
			If conn IsNot Nothing Then conn.Dispose()  
			If conn IsNot Nothing Then conn.Dispose()  
			If dr IsNot Nothing AndAlso Not (dr.IsClosed) Then dr.Close()
		End Try	
		
		Return boolReturn

    End Function	
	
End Class