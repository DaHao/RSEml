Imports System.IO
Imports System.Text.RegularExpressions

Public Class EMLParser2
#Region "Property"
    Private _x_Sender As String
    Private _x_Receivers As String()
    Private _Received As String
    Private _Mime_Version As String
    Private _From As String
    Private _To As String
    Private _ArrTo As String() = Nothing
    Private _CC As String
    Private _ArrCC As String() = Nothing
    Private _Date As DateTime = DateTime.MinValue
    Private _Subject As String
    Private _Content_Type As String
    Private _Content_Transfer_Encoding As String
    Private _Return_Path As String
    Private _Message_ID As String
    Private _x_OriginalArrivaTime As DateTime = DateTime.MinValue
    Private _Body As String
    Private _HTMLBody As String
    Private _attStr As String
    Private _attStrZip As String
    Private _att As System.Net.Mail.Attachment
    Private _listUnsupported As Dictionary(Of String, String) = Nothing

    Public Property X_Sender As String
        Get
            Return _x_Sender
        End Get
        Set(ByVal value As String)
            _x_Sender = value
        End Set
    End Property
    Public Property X_Receivers As String()
        Get
            Return _x_Receivers
        End Get
        Set(ByVal value As String())
            _x_Receivers = value
        End Set
    End Property
    Public Property Received As String
        Get
            Return _Received
        End Get
        Set(ByVal value As String)
            _Received = value
        End Set
    End Property
    Public Property Mime_Version As String
        Get
            Return _Mime_Version
        End Get
        Set(ByVal value As String)
            _Mime_Version = value
        End Set
    End Property
    Public Property From As String
        Get
            Return _From
        End Get
        Set(ByVal value As String)
            _From = value
        End Set
    End Property
    Public Property ToStr As String
        Get
            Return _To
        End Get
        Set(ByVal value As String)
            _To = value
        End Set
    End Property
    Public Property ArrTo As String()
        Get
            Return _ArrTo
        End Get
        Set(ByVal value As String())
            _ArrTo = value
        End Set
    End Property
    Public Property CC As String
        Get
            Return _CC
        End Get
        Set(ByVal value As String)
            _CC = value
        End Set
    End Property
    Public Property ArrCC As String()
        Get
            Return _ArrCC
        End Get
        Set(ByVal value As String())
            _ArrCC = value
        End Set
    End Property
    Public Property SendTime As DateTime
        Get
            Return _Date
        End Get
        Set(ByVal value As DateTime)
            _Date = value
        End Set
    End Property
    Public Property Subject As String
        Get
            Return _Subject
        End Get
        Set(ByVal value As String)
            _Subject = value
        End Set
    End Property
    Public Property Content_Type As String
        Get
            Return _Content_Type
        End Get
        Set(ByVal value As String)
            _Content_Type = value
        End Set
    End Property
    Public Property Content_Transfer_Encoding As String
        Get
            Return _Content_Transfer_Encoding
        End Get
        Set(ByVal value As String)
            _Content_Transfer_Encoding = value
        End Set
    End Property
    Public Property Return_Path As String
        Get
            Return _Return_Path
        End Get
        Set(ByVal value As String)
            _Return_Path = value
        End Set
    End Property
    Public Property Message_ID As String
        Get
            Return _Message_ID
        End Get
        Set(ByVal value As String)
            _Message_ID = value
        End Set
    End Property
    Public Property X_OriginalArrivalTime As DateTime
        Get
            Return _x_OriginalArrivaTime
        End Get
        Set(ByVal value As DateTime)
            _x_OriginalArrivaTime = value
        End Set
    End Property
    Public Property Body As String
        Get
            Return _Body
        End Get
        Set(ByVal value As String)
            _Body = value
        End Set
    End Property
    Public Property HTMLBody As String
        Get
            Return _HTMLBody
        End Get
        Set(ByVal value As String)
            _HTMLBody = value
        End Set
    End Property
    Public Property attStr As String
        Get
            Return _attStr
        End Get
        Set(ByVal value As String)
            _attStr = value
        End Set
    End Property
    Public Property attStrZip As String
        Get
            Return _attStrZip
        End Get
        Set(ByVal value As String)
            _attStrZip = value
        End Set
    End Property
    Public Property att As System.Net.Mail.Attachment
        Get
            Return _att
        End Get
        Set(ByVal value As System.Net.Mail.Attachment)
            _att = value
        End Set
    End Property
    Public Property UnsupportedHeaders As Dictionary(Of String, String)
        Get
            Return _listUnsupported
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            _listUnsupported = value
        End Set
    End Property

#End Region

    Public Sub New(ByVal FsEML As FileStream)
        ParseEML(FsEML)
    End Sub

    Private Sub ParseEML(ByVal FsEML As FileStream)
        Dim sr As New StreamReader(FsEML)
        Dim sLine As String = sr.ReadLine
        Dim content As List(Of String)

        While sLine IsNot Nothing
            content.Add(sLine)
        End While
    End Sub
End Class
