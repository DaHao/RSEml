Imports System.IO
Imports System.Text.RegularExpressions

Namespace Hlib.MailFormats
    Public Class EmlParser

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
        Private _ContentType As String
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

        Public Property XSender As String
            Get
                Return _x_Sender
            End Get
            Set(ByVal value As String)
                _x_Sender = value
            End Set
        End Property
        Public Property XReceivers As String()
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
        Public Property MimeVersion As String
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
        Public Property ContentType As String
            Get
                Return _ContentType
            End Get
            Set(ByVal value As String)
                _ContentType = value
            End Set
        End Property
        Public Property ContentTransferEncoding As String
            Get
                Return _Content_Transfer_Encoding
            End Get
            Set(ByVal value As String)
                _Content_Transfer_Encoding = value
            End Set
        End Property
        Public Property ReturnPath As String
            Get
                Return _Return_Path
            End Get
            Set(ByVal value As String)
                _Return_Path = value
            End Set
        End Property
        Public Property MessageId As String
            Get
                Return _Message_ID
            End Get
            Set(ByVal value As String)
                _Message_ID = value
            End Set
        End Property
        Public Property XOriginalArrivalTime As DateTime
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
        Public Property HtmlBody As String
            Get
                Return _HTMLBody
            End Get
            Set(ByVal value As String)
                _HTMLBody = value
            End Set
        End Property
        Public Property AttStr As String
            Get
                Return _attStr
            End Get
            Set(ByVal value As String)
                _attStr = value
            End Set
        End Property
        Public Property AttStrZip As String
            Get
                Return _attStrZip
            End Get
            Set(ByVal value As String)
                _attStrZip = value
            End Set
        End Property
        Public Property Att As System.Net.Mail.Attachment
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

        Public Sub New(ByVal fsEml As FileStream)
            parseEML(fsEml)
        End Sub

        Private Sub parseEML(ByVal FsEML As FileStream)

            Dim myc As New System.Globalization.CultureInfo("zh-TW", False)
            System.Threading.Thread.CurrentThread.CurrentCulture = myc

            Dim sr As New StreamReader(FsEML)
            Dim sLine As String = sr.ReadLine '讀取行數
            Dim listAll As New List(Of String) '所有資料行
            While sLine IsNot Nothing
                listAll.Add(sLine)
                sLine = sr.ReadLine
            End While

            Dim list As New List(Of String)
            Dim nStartBody As Integer = -1 '目前所在行數
            Dim saAll(listAll.Count) As String '複本
            listAll.CopyTo(saAll)

            '取得信件開頭資料
            For i As Integer = LBound(saAll) To UBound(saAll)
                If saAll(i) = String.Empty Then
                    nStartBody = i
                    Exit For
                End If

                Dim sFullValue As String = saAll(i)
                GetFullValue(saAll, i, sFullValue)
                list.Add(sFullValue)

                'Debug.WriteLine(sFullValue)
            Next i

            SetFields(list.ToArray())

            If nStartBody = -1 Then
                Return
            End If

            '取得內容 
            '1.contentType = multipart/alternative
            '2.contentType = multipart/mixed
            '3.直接輸出
            If ContentType IsNot Nothing And ContentType.ToLower.Contains("multipart/alternative") Then

                Dim ix As Integer = ContentType.ToLower.IndexOf("boundary")
                If ix = -1 Then
                    Return
                End If

                Dim sBoundaryMarker As String = ContentType.Substring(ix + 8).Trim()

                list = New List(Of String)
                For n As Integer = nStartBody + 1 To saAll.Length - 1
                    If saAll(n).Contains(sBoundaryMarker) Then
                        If list.Count > 0 Then
                            SetBody(list)
                            list = New List(Of String)
                        End If
                        Continue For
                    End If
                    list.Add(saAll(n))
                Next

            ElseIf ContentType.ToLower.Contains("multipart/mixed") Then
                Dim ix As Integer = ContentType.ToLower.IndexOf("boundary")
                If ix = -1 Then
                    Return
                End If

                'boundary標記
                Dim sboundaryMarker As String = ContentType.Substring(ix + 9).Trim()
                Dim contentFlag As Boolean = False
                Dim boundaryFlag As Boolean = False

                Body = String.Empty
                For n As Integer = nStartBody + 1 To saAll.Length - 1
                    If saAll(n).Contains(sboundaryMarker) OrElse boundaryFlag Then

                        boundaryFlag = True

                        If saAll(n).Contains(sboundaryMarker) And contentFlag Then
                            Exit For
                        End If

                        If saAll(n) Is String.Empty OrElse contentFlag Then
                            If contentFlag = False Then
                                contentFlag = True
                            Else
                                Body &= saAll(n)
                            End If
                        End If
                    Else
                        nStartBody += 1
                        Continue For
                    End If

                    nStartBody += 1

                Next

            Else
                Body = String.Empty
                For n As Integer = nStartBody + 1 To saAll.Length - 1
                    Body &= saAll(n) & "\r\n"
                Next

            End If

            '抓附件
            Dim index As Integer = ContentType.ToLower.IndexOf("boundary")
            Dim sboundary As String = ContentType.Substring(index + 9).Trim()
            Dim attFlag As Boolean = False
            Dim emptyFlag As Boolean = False
            Dim zipFlag As Boolean = False
            attStr = String.Empty
            For i As Integer = nStartBody + 1 To saAll.Length - 1
                If saAll(i).ToLower.Contains("application/zip") Then
                    zipFlag = True
                End If
                If saAll(i).ToLower.Contains("attachment") Then
                    attFlag = True
                End If
                If saAll(i) Is String.Empty And attFlag Then
                    emptyFlag = True
                End If
                If emptyFlag Then
                    If saAll(i).Contains(sboundary) Then
                        Exit For
                    End If
                    attStr &= saAll(i)
                End If
            Next

            If zipFlag Then
                attStrZip = attStr
            End If

            '壓縮檔需編成utf8
            attStr = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(attStr))
            'attStr = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(attStr))

        End Sub

        Private Sub SetFields(ByVal saLines As String())

            _listUnsupported = New Dictionary(Of String, String)
            Dim listX_Receiver As New List(Of String)

            For Each sHdr As String In saLines

                Dim a() As Char = {":"}
                Dim saHdr() As String = sHdr.Split(a, 2)
                If saHdr Is Nothing Then
                    Continue For
                End If

                Select Case saHdr(0).ToLower
                    Case "x-sender"
                        XSender = saHdr(1)
                    Case "x-receiver"
                        listX_Receiver.Add(saHdr(1))
                    Case "received"
                        Received = saHdr(1)
                    Case "mime-version"
                        MimeVersion = saHdr(1)
                    Case "from"
                        Dim matches As MatchCollection = New Regex("[a-zA-Z0-9_\-\.]+@[a-zA-Z0-9\-\.]+\.[A-Za-z]{2,4}").Matches(saHdr(1).Trim)
                        If matches.Count > 0 Then
                            For Each match As Match In matches
                                From = match.Value
                            Next
                        End If
                        'From = saHdr(1)

                    Case "to"
                        Dim listTo As New List(Of String)

                        Dim matches As MatchCollection = New Regex("[a-zA-Z0-9_\-\.]+@[a-zA-Z0-9\-\.]+\.[A-Za-z]{2,4}").Matches(saHdr(1).Trim)
                        'Dim matches As MatchCollection = New Regex("[\w_]+@[\w\.]+\w+").Matches(saHdr(1))
                        If matches.Count > 0 Then
                            For Each match As Match In matches
                                listTo.Add(match.Value)
                            Next
                        End If

                        Dim tempTo(listTo.Count - 1) As String
                        ArrTo = tempTo
                        listTo.CopyTo(ArrTo)

                        ToStr = saHdr(1)

                    Case "cc"
                        Dim listCC As New List(Of String)
                        Dim matches As MatchCollection = New Regex("[a-zA-Z0-9_\-\.]+@[a-zA-Z0-9\-\.]+\.[A-Za-z]{2,4}").Matches(saHdr(1).Trim)
                        If matches.Count > 0 Then
                            For Each match As Match In matches
                                listCC.Add(match.Value)
                            Next
                        End If

                        Dim tempCC(listCC.Count - 1) As String
                        ArrCC = tempCC
                        listCC.CopyTo(ArrCC)

                        CC = saHdr(1)

                    Case "date"
                        SendTime = DateTime.Parse(saHdr(1))
                    Case "subject"
                        Subject = saHdr(1)
                    Case "content-type"
                        ContentType = saHdr(1)
                    Case "content-transfer-encoding"
                        ContentTransferEncoding = saHdr(1)
                    Case "return-path"
                        ReturnPath = saHdr(1)
                    Case "message-id"
                        MessageId = saHdr(1)
                    Case "x-originalarrivaltime"
                        Dim ix As Integer = saHdr(1).IndexOf("FILETIME")
                        If ix <> -1 Then
                            Dim sOAT As String = saHdr(1).Substring(0, ix)
                            sOAT = sOAT.Replace("(UTC)", "-0000")
                            XOriginalArrivalTime = DateTime.Parse(sOAT)
                        End If
                    Case Else
                        _listUnsupported.Add(saHdr(0), saHdr(1))
                End Select
            Next

            Dim xre(listX_Receiver.Count) As String
            XReceivers = xre
            listX_Receiver.CopyTo(XReceivers)

        End Sub

        Private Sub SetBody(ByVal list As List(Of String))
            Dim bIsHTML As Boolean = False
            Dim bIsBodyStart As Boolean = False
            Dim listBody As New List(Of String)

            For Each s As String In list
                If s.ToLower.StartsWith("content-type") Then
                    If s.ToLower.Contains("text/html") Then
                        bIsHTML = True
                    ElseIf Not s.ToLower.Contains("text/plain") Then
                        Return
                    End If
                ElseIf s = String.Empty And Not bIsBodyStart Then
                    bIsBodyStart = True
                ElseIf bIsBodyStart Then
                    listBody.Add(s)
                End If
            Next

            Dim sa(listBody.Count) As String
            listBody.CopyTo(sa)

            If bIsHTML Then
                HTMLBody = String.Join("\r\n", sa)
            Else
                Body = String.Join("\r\n", sa)
            End If
        End Sub

        '合併被空白分行的資料
        Private Sub GetFullValue(ByVal sa As String(), ByRef i As Integer, ByRef sValue As String)
            If i + 1 <= sa.Length AndAlso sa(i + 1) IsNot String.Empty AndAlso Char.IsWhiteSpace(sa(i + 1), 0) Then
                i += 1
                sValue &= " " & sa(i).Trim
                GetFullValue(sa, i, sValue)
            End If
        End Sub

    End Class
End Namespace

