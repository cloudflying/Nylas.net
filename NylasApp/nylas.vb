Imports System.IO
Imports System.Web
Imports System.Net
Imports System.Text
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions
Public Class nylas
    Public Shared Property nylasClientID As String
    Public Shared Property nylasClientSecret As String
    Public Shared Property nylasAPIaddress As String = "https://api.nylas.com"
    Public Shared Property redirectAddress As String



#Region "OAuth - Not finished Yet"
    ''' <summary>
    ''' Generates redirect url for OAuth.
    ''' </summary>
    ''' <param name="login_hint">Optional, Email address of user.</param>
    ''' <returns>String, Redirect user to this url.</returns>
    Public Shared Function oauthRedirect(Optional login_hint As String = "") As String
        chkParams()
        If Len(redirectAddress) < 1 Then
            Throw New ApplicationException("102 - No Redirect URI Specified.")
        End If

        Try
            If Len(login_hint) > 0 Then
                Return nylasAPIaddress & "/oauth/authorize?client_id=" & nylasClientID & "&response_type=code&scope=email&login_hint=" & login_hint & "&redirect_uri=" & redirectAddress
            Else
                Return nylasAPIaddress & "/oauth/authorize?client_id=" & nylasClientID & "&response_type=code&scope=email&redirect_uri=" & redirectAddress
            End If
        Catch ex As Exception
            Throw New ApplicationException("Error 501 - URL Generation Failed : ", ex)
        End Try
    End Function

    Public Shared Function oauthTokenExchange(code As String) As String
        chkParams()
        Dim codeExchange As New oauthTokenCode
        Try
            Dim urlBuilder As New StringBuilder
            urlBuilder.Append(nylasAPIaddress)
            urlBuilder.Append("/oauth/token?")
            urlBuilder.Append("client_id=" & nylasClientID)
            urlBuilder.Append("&client_secret=" & nylasClientSecret)
            urlBuilder.Append("&grant_type=authorization_code")
            urlBuilder.Append("&code=" & code)

            codeExchange = JsonConvert.DeserializeObject(Of oauthTokenCode)(requestGET(urlBuilder.ToString))
        Catch ex As Exception

        End Try


        If Len(codeExchange.error) > 1 Then
            Throw New ApplicationException("502 - OAuth (" & codeExchange.error & ") " & codeExchange.reason)
        End If
        If Len(codeExchange.access_token) < 1 Then
            Throw New ApplicationException("503 - OAuth - No Access Token Received.")
        End If

        Return codeExchange.access_token
    End Function
#End Region

#Region "Messages"
    ''' <summary>
    ''' Retrieves all messages from Nylas Server. API Call /messages
    ''' </summary>
    ''' <param name="token">Required, OAuth Token for User</param>
    ''' <param name="filters">Optional, Filters parameters (Querystring values)</param>
    ''' <returns></returns>
    Public Shared Function retrieveAllMessages(token As String, Optional filters As String = "") As List(Of messageObject)
        chkParams()
        Dim returnList As New List(Of messageObject)
        Try

            If Len(filters) > 1 Then filters = "?" & filters

            returnList = JsonConvert.DeserializeObject(Of List(Of messageObject))(requestGET(nylasAPIaddress & "/messages" & filters, token))
        Catch ex As Exception
            Throw New ApplicationException("600 - Unable to gather message list from Nylas. " & vbNewLine & ex.Message, ex)
        End Try

        Return returnList
    End Function

    ''' <summary>
    ''' Retrieves a single message. API Call /messages/:id
    ''' </summary>
    ''' <param name="token">Required, OAuth Token for User</param>
    ''' <param name="messageID">Required, Message ID</param>
    ''' <returns></returns>
    Public Shared Function retrieveSingleMessage(token As String, messageID As String) As messageObject
        chkParams()
        Dim returnMessage As messageObject
        Try
            returnMessage = JsonConvert.DeserializeObject(Of messageObject)(requestGET(nylasAPIaddress & "/messages/" & messageID, token))
        Catch ex As Exception
            Throw New ApplicationException("601 - Unable to gather single message from Nylas. " & vbNewLine & ex.Message, ex)
        End Try

        Return returnMessage
    End Function
#End Region

#Region "Threads"
    Public Shared Function retrieveAllThreads(token As String, Optional filters As String = "") As List(Of threadObject)
        chkParams()
        Dim returnThread As New List(Of threadObject)
        Try
            If Len(filters) > 0 Then filters = "?" & filters
            returnThread = JsonConvert.DeserializeObject(Of List(Of threadObject))(requestGET(nylasAPIaddress & "/threads" & filters, token))
        Catch ex As Exception
            Throw New ApplicationException("630 - Unable to gather thread list from Nylas. " & vbNewLine & ex.Message, ex)
        End Try
        Return returnThread
    End Function
#End Region

#Region "Account"
    Public Shared Function retrieveAccount(token As String) As accountObject
        chkParams()
        Try
            retrieveAccount = JsonConvert.DeserializeObject(Of accountObject)(requestGET(nylasAPIaddress & "/account", token))
        Catch ex As Exception
            Throw New ApplicationException("650 - Unable to retrieve account object." & vbNewLine & ex.Message, ex)
        End Try

    End Function
#End Region

#Region "Helper Functions - Do Not Edit Below Here"

    Public Shared Function convertDate(unixTime As Double) As DateTime

        Dim dtDateTime As DateTime = New DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc)
        dtDateTime = dtDateTime.AddSeconds(unixTime).ToLocalTime
        Return dtDateTime
    End Function

    Private Shared Function chkParams() As Boolean
        If Len(nylasClientID) < 1 Then
            Throw New ApplicationException("100 - No Client ID Specified.")
        End If
        If Len(nylasClientSecret) < 1 Then
            Throw New ApplicationException("101 - No Client Secret Specified.")
        End If

        Return True
    End Function


#Region "HTTP Requests"
    Private Shared Function requestGET(url As String, Optional token As String = "") As String
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        If Len(token) > 0 Then
            request.Headers.Add("AUTHORIZATION", "Basic " & Convert.ToBase64String(Encoding.UTF8.GetBytes((token & ":").ToCharArray())))
        End If
        Try
            Dim oResponse As HttpWebResponse = request.GetResponse()
            Dim reader As New StreamReader(oResponse.GetResponseStream())
            Dim tmp As String = reader.ReadToEnd()
            oResponse.Close()
            If oResponse.StatusCode = "200" Then
                Return tmp
            Else
                Throw New ApplicationException("Error Occurred, Code: " & oResponse.StatusCode)
            End If

        Catch ex As WebException
            If ex.Status = WebExceptionStatus.ProtocolError Then
                Dim read As New StreamReader(ex.Response.GetResponseStream())
                Dim tmp As String = read.ReadToEnd()
                read.Close()
                Throw New ApplicationException(CType(ex.Response, HttpWebResponse).StatusCode & " : " & tmp)
            Else
                Throw New ApplicationException(ex.Message)
            End If
        End Try
    End Function
#End Region
#End Region

#Region "HTML to Plain Text"
    Public Shared Function HTMLToText(ByVal HTMLCode As String, Optional newLineSeparator As String = vbNewLine) As String
        HTMLCode = HTMLCode.Replace("\n", newLineSeparator)
        HTMLCode = HTMLCode.Replace("\t", " ")
        HTMLCode = Regex.Replace(HTMLCode, "\\s+", "  ")
        HTMLCode = Regex.Replace(HTMLCode, "<head.*?</head>", "", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        HTMLCode = Regex.Replace(HTMLCode, "<script.*?</script>", "", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        HTMLCode = System.Text.RegularExpressions.Regex.Replace(HTMLCode, "< *p[^>]*>", "$&" & newLineSeparator, RegexOptions.IgnoreCase)
        HTMLCode = System.Text.RegularExpressions.Regex.Replace(HTMLCode, "<[^>]*>", "")
        HTMLCode = System.Text.RegularExpressions.Regex.Replace(HTMLCode, "&nbsp;", " ", RegexOptions.IgnoreCase And RegexOptions.IgnorePatternWhitespace)
        HTMLCode = System.Text.RegularExpressions.Regex.Replace(HTMLCode, "&amp;", "&", RegexOptions.IgnoreCase And RegexOptions.IgnorePatternWhitespace)
        HTMLCode = System.Text.RegularExpressions.Regex.Replace(HTMLCode, "&quot;", """", RegexOptions.IgnoreCase And RegexOptions.IgnorePatternWhitespace)
        HTMLCode = System.Text.RegularExpressions.Regex.Replace(HTMLCode, "&lt;", "<", RegexOptions.IgnoreCase And RegexOptions.IgnorePatternWhitespace)
        HTMLCode = System.Text.RegularExpressions.Regex.Replace(HTMLCode, "&gt;", ">", RegexOptions.IgnoreCase And RegexOptions.IgnorePatternWhitespace)
        Return HTMLCode
    End Function

    Public Function ConvertHTMLToText(ByVal Source As String) As String

        Dim result As String = Source

        ' Remove formatting that will prevent regex from running reliably
        ' \r - Matches a carriage return \u000D.
        ' \n - Matches a line feed \u000A.
        ' \f - Matches a form feed \u000C.
        ' For more details see http://msdn.microsoft.com/en-us/library/4edbef7e.aspx
        result = Replace(result, "[\r\n\f]", String.Empty, Text.RegularExpressions.RegexOptions.IgnoreCase)

        ' replace the most commonly used special characters:
        result = Replace(result, "&lt;", "<", RegexOptions.IgnoreCase)
        result = Replace(result, "&gt;", ">", RegexOptions.IgnoreCase)
        result = Replace(result, "&nbsp;", " ", RegexOptions.IgnoreCase)
        result = Replace(result, "&quot;", """", RegexOptions.IgnoreCase)
        result = Replace(result, "&amp;", "&", RegexOptions.IgnoreCase)

        ' Remove ASCII character code sequences such as &#nn; and &#nnn;
        result = Replace(result, "&#[0-9]{2,3};", String.Empty, RegexOptions.IgnoreCase)

        ' Remove all other special characters. More can be added - see the following for more details:
        ' http://www.degraeve.com/reference/specialcharacters.php
        ' http://www.web-source.net/symbols.htm
        result = Replace(result, "&.{2,6};", String.Empty, RegexOptions.IgnoreCase)

        ' Remove all attributes and whitespace from the <head> tag
        result = Replace(result, "< *head[^>]*>", "<head>", RegexOptions.IgnoreCase)
        ' Remove all whitespace from the </head> tag
        result = Replace(result, "< */ *head *>", "</head>", RegexOptions.IgnoreCase)
        ' Delete everything between the <head> and </head> tags
        result = Replace(result, "<head>.*</head>", String.Empty, RegexOptions.IgnoreCase)

        ' Remove all attributes and whitespace from all <script> tags
        result = Replace(result, "< *script[^>]*>", "<script>", RegexOptions.IgnoreCase)
        ' Remove all whitespace from all </script> tags
        result = Replace(result, "< */ *script *>", "</script>", RegexOptions.IgnoreCase)
        ' Delete everything between all <script> and </script> tags
        result = Replace(result, "<script>.*</script>", String.Empty, RegexOptions.IgnoreCase)

        ' Remove all attributes and whitespace from all <style> tags
        result = Replace(result, "< *style[^>]*>", "<style>", RegexOptions.IgnoreCase)
        ' Remove all whitespace from all </style> tags
        result = Replace(result, "< */ *style *>", "</style>", RegexOptions.IgnoreCase)
        ' Delete everything between all <style> and </style> tags
        result = Replace(result, "<style>.*</style>", String.Empty, RegexOptions.IgnoreCase)

        ' Insert tabs in place of <td> tags
        result = Replace(result, "< *td[^>]*>", vbTab, RegexOptions.IgnoreCase)

        ' Insert single line breaks in place of <br> and <li> tags
        result = Replace(result, "< *br[^>]*>", vbCrLf, RegexOptions.IgnoreCase)
        result = Replace(result, "< *li[^>]*>", vbCrLf, RegexOptions.IgnoreCase)

        ' Insert double line breaks in place of <p>, <div> and <tr> tags
        result = Replace(result, "< *div[^>]*>", vbCrLf + vbCrLf, RegexOptions.IgnoreCase)
        result = Replace(result, "< *tr[^>]*>", vbCrLf + vbCrLf, RegexOptions.IgnoreCase)
        result = Replace(result, "< *p[^>]*>", vbCrLf + vbCrLf, RegexOptions.IgnoreCase)

        ' Remove all reminaing html tags
        result = Replace(result, "<[^>]*>", String.Empty, RegexOptions.IgnoreCase)

        ' Replace repeating spaces with a single space
        result = Replace(result, " +", " ")

        ' Remove any trailing spaces and tabs from the end of each line
        result = Replace(result, "[ \t]+\r\n", vbCrLf)

        ' Remove any leading whitespace characters
        result = Replace(result, "^[\s]+", String.Empty)

        ' Remove any trailing whitespace characters
        result = Replace(result, "[\s]+$", String.Empty)

        ' Remove extra line breaks if there are more than two in a row
        result = Replace(result, "\r\n\r\n(\r\n)+", vbCrLf + vbCrLf)

        ' Thats it.
        Return result

    End Function
#End Region
End Class
