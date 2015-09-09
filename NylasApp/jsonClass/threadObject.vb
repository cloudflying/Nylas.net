Public Class threadObject
    Public Property id As String
    Public Property [object] As String
    Public Property account_id As String
    Public Property subject As String
    Public Property unread As Boolean
    Public Property starred As Boolean
    Public Property last_message_timestamp As Integer
    Public Property received_recent_date As Integer
    Public Property first_message_timestamp As Integer
    Public Property participants As New List(Of msgAddressArray)
    Public Property snippet As String
    Public Property folders As New List(Of Folder)
    Public Property message_ids As String()
    Public Property draft_ids As String()
    Public Property version As Integer
End Class
