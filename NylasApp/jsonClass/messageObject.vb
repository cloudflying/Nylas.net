Public Class messageObject
    Public Property id As String
    Public Property [object] As String
    Public Property account_id As String
    Public Property thread_id As String
    Public Property subject As String
    Public Property from As New List(Of msgAddressArray)
    Public Property [to] As New List(Of msgAddressArray)
    Public Property cc As New List(Of msgAddressArray)
    Public Property bcc As New List(Of msgAddressArray)
    Public Property [date] As Integer
    Public Property unread As Boolean
    Public Property starred As Boolean
    Public Property folder As Folder
    Public Property snippet As String
    Public Property body As String
    Public Property files As New List(Of msgAttachedFiles)
    Public Property events As New List(Of calendarMsgEvent)
End Class


Public Class msgAddressArray
    Public Property name As String
    Public Property email As String
End Class
Public Class Folder
    Public Property name As String
    Public Property display_name As String
    Public Property id As String
End Class

Public Class msgAttachedFiles
    Public Property content_type As String
    Public Property filename As String
    Public Property id As String
    Public Property size As Integer
    Public Property content_id As String
End Class

Public Class Participant
    Public Property email As String
    Public Property name As String
    Public Property status As String
End Class

Public Class calendarMsgWhen
    Public Property [object] As String
    Public Property end_time As Integer
    Public Property start_time As Integer
End Class

Public Class calendarMsgEvent
    Public Property [object] As String
    Public Property id As String
    Public Property calendar_id As String
    Public Property account_id As String
    Public Property description As String
    Public Property location As String
    Public Property participants As New List(Of Participant)
    Public Property read_only As Boolean
    Public Property title As String
    Public Property [when] As New calendarMsgWhen
    Public Property busy As Boolean
    Public Property status As String
End Class

