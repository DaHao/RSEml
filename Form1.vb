Imports RSEml.Hlib.MailFormats
Imports System.IO
Imports System.IO.Compression
Imports System.Net.Mail
Imports System.Net.Mime
Imports ICSharpCode.SharpZipLib
Imports Ionic.Zip

Public Class Form1

#Region "功能列"

    Private Sub Memo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Memo.Click, MyBase.Shown
        smtpGroup.Visible = False
        TextBox1.Visible = True
        TextBox1.Focus()
        TextBox1.Text = "本程式可寄送單一或多個EML檔案。" & vbNewLine &
            "單一檔案請使用【BrowseFile】；" & vbNewLine &
            "如為ZIP檔，將先解壓再進行寄發作業。" & vbNewLine &
            "多個檔案請使用【BrowseFolder】選取資料夾。" & vbNewLine & vbNewLine &
            "【Ｓｅｎｄ】寄發Mail。" & vbNewLine &
            "【Ｃｌｅａｒ】清除畫面上所有欄位。" & vbNewLine &
            "【Ｅｘｉｔ】離開本程式。" & vbNewLine & vbNewLine &
            "寄送前請點選功能列SMTP，以進行設定：" & vbNewLine &
            "點選自訂選項可顯示上次儲存的設定值。" & vbNewLine &
            "【Ｓａｖｅ】儲存SMTP設定。" & vbNewLine &
            "【Ｒｅｓｅｔ】回復預設值。" & vbNewLine &
            "【Ｃａｎｃｅｌ】取消設定，返回訊息畫面。" & vbNewLine & vbNewLine &
            "點選Help > Hint可再次觀看本說明。"

    End Sub

    'Private Sub GoogleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoogleToolStripMenuItem.Click
    '    TextBox1.Visible = False
    '    smtpGroup.Visible = True
    '    smtpHost.Text = "smtp.gmail.com"
    '    smtpPort.Text = "25"
    '    smtpSsl.Checked = True
    'End Sub

    'Private Sub HiNetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HiNetToolStripMenuItem.Click
    '    TextBox1.Visible = False
    '    smtpGroup.Visible = True
    '    smtpHost.Text = "ms2.hinet.net"
    '    smtpPort.Text = "25"
    '    smtpSsl.Checked = False
    'End Sub

    Private Sub SelfToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelfToolStripMenuItem.Click
        My.Settings.Reload()
        TextBox1.Visible = False
        smtpGroup.Visible = True
        smtpHost.Text = My.Settings.SMTPHost
        smtpPort.Text = My.Settings.SMTPPort
        'smtpSsl.Checked = My.Settings.nSSL
        userAcc.Text = My.Settings.Acc
        userPw.Text = My.Settings.Pw
    End Sub

#End Region

#Region "按鈕事件"

    Private Sub BrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrowseFile.Click

        Using openFileDialog As New OpenFileDialog
            openFileDialog.Filter = "eml檔、zip檔(*.eml,*.zip)|*.eml;*.zip"
            openFileDialog.Title = "選取檔案"
            openFileDialog.Multiselect = False

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                BrowserFolder.Enabled = False
                FolderPath.Enabled = False

                If openFileDialog.FileName <> "" Then
                    FilePath.Text = openFileDialog.FileName
                End If
            End If
        End Using
    End Sub

    Private Sub BrowserFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrowserFolder.Click

        Using openFolder As New FolderBrowserDialog

            If openFolder.ShowDialog() = DialogResult.OK Then
                BrowseFile.Enabled = False
                FilePath.Enabled = False
                TextBox1.Text = ""

                FolderPath.Text = openFolder.SelectedPath
            End If
            '訊息欄顯示檔案名
            Dim fileName As String = Dir(FolderPath.Text & "\*.eml")
            Do Until String.IsNullOrEmpty(fileName)
                TextBox1.AppendText(fileName & vbNewLine)
                fileName = Dir()
            Loop
        End Using
    End Sub

    Private Sub send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click

        TextBox1.Text = ""
        TextBox1.ForeColor = Color.Black
        Dim errCount As Integer = 0
        Dim sucCount As Integer = 0

        Dim emlFiles As New List(Of String)

        If Not String.IsNullOrEmpty(FilePath.Text) Then
            '單檔，如果是zip的話要先解壓
            If FilePath.Text.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) Then
                Dim dirPath As String = ExtractZip(FilePath.Text)
                If dirPath.Equals("error") Then
                    MsgBox("解壓發生錯誤，請檢查選擇檔案!")
                    Exit Sub
                End If

                Dim dirs As String() = Directory.GetFiles(dirPath, "*.eml")

                For Each dir As String In dirs
                    TextBox1.AppendText(Path.GetFileName(dir) & vbNewLine)
                    emlFiles.Add(dir)
                Next
            Else
                emlFiles.Add(FilePath.Text)
            End If

        ElseIf Not String.IsNullOrEmpty(FolderPath.Text) Then
            '資料夾
            Dim dirs As String() = Directory.GetFiles(FolderPath.Text, "*.eml")
            For Each dir As String In dirs
                TextBox1.AppendText(Path.GetFileName(dir) & vbNewLine)
                emlFiles.Add(dir)
            Next
        Else
            MsgBox("請選擇寄送檔案！", MsgBoxStyle.Critical)
            Exit Sub
        End If

        If emlFiles.Count = 0 Then
            MsgBox("查無待發信件！")
            Exit Sub
        End If


        Dim msgStyle As Integer = MsgBoxStyle.OkCancel + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2
        Dim msg As String = "確認寄發信件？"
        Dim ret As Integer = MsgBox(msg, msgStyle)

        ' Display the ProgressBar control.
        ProgressBar1.Visible = True
        ProgressBar1.Style = ProgressBarStyle.Continuous
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = emlFiles.Count
        ProgressBar1.Value = 0
        ProgressBar1.Step = 1

        If ret = MsgBoxResult.Ok Then

            TextBox1.Text = ""
            For Each emlfile As String In emlFiles

                Try
                    Dim fs As IO.FileStream = File.Open(emlfile, FileMode.Open, FileAccess.ReadWrite)
                    Dim reader As EmlParser = New EmlParser(fs)

                    Dim emlStaus As String = ""
                    'emlStaus = SendMail(reader.ToStr, reader.From, reader.Subject, reader.Body, StreamOutputAttachment(reader.attStr, reader.Subject))
                    emlStaus = SendMail(reader)

                    If emlStaus.Equals("ok") Then
                        sucCount += 1
                        TextBox1.ForeColor = Color.Black
                        TextBox1.Text &= Path.GetFileName(emlfile) & " 寄送成功！" & vbNewLine & vbNewLine

                    ElseIf emlStaus.Equals("smtp") Then
                        'TextBox1.ForeColor = Color.Blue
                        errCount += 1
                        TextBox1.Text &= Path.GetFileName(emlfile) & "寄送失敗。請檢查SMTP設定。"
                        Exit Sub
                    Else
                        errCount += 1
                        'TextBox1.ForeColor = Color.Red
                        TextBox1.Text &= Path.GetFileName(emlfile) & " 寄送失敗。" & vbNewLine
                        TextBox1.Text &= "Error Message： " & emlStaus & vbNewLine & vbNewLine
                    End If


                Catch ex As Exception
                    errCount += 1
                    'TextBox1.ForeColor = Color.Red
                    TextBox1.Text &= Path.GetFileName(emlfile) & " 寄送失敗。" & vbNewLine
                    TextBox1.Text &= "Error Message： " & ex.Message & vbNewLine & vbNewLine
                Finally

                End Try

                ProgressBar1.PerformStep()
            Next

            MsgBox("共寄發：" & (sucCount + errCount).ToString & "封信件" & vbNewLine & "成功筆數：" & sucCount.ToString & vbNewLine & "失敗筆數：" & errCount.ToString)
            ProgressBar1.Value = 0
        Else
            Exit Sub
        End If

    End Sub

    Private Sub reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles reset.Click

        BrowseFile.Enabled = True
        FilePath.Enabled = True
        FilePath.Text = ""

        BrowserFolder.Enabled = True
        FolderPath.Enabled = True
        FolderPath.Text = ""

        TextBox1.ForeColor = Color.Black
        TextBox1.Text = ""

    End Sub

    Private Sub exitbutton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exitbutton.Click

        Dim ret As Integer
        Dim msgStyle As Integer
        Dim msg As String

        msgStyle = MsgBoxStyle.OkCancel + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2
        msg = "確定離開？"

        ret = MsgBox(msg, msgStyle)
        If ret = MsgBoxResult.Ok Then
            Me.Close()
            End
        End If

    End Sub

    Private Sub smtpok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles smtpok.Click
        Dim ret As Integer
        Dim msgStyle As Integer

        My.Settings.SMTPHost = smtpHost.Text
        My.Settings.SMTPPort = smtpPort.Text
        'My.Settings.nSSL = smtpSsl.Checked
        My.Settings.Acc = userAcc.Text
        My.Settings.Pw = userPw.Text

        msgStyle = MsgBoxStyle.OkCancel + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton2
        ret = MsgBox("確定儲存SMTP設定？", msgStyle)
        If ret = MsgBoxResult.Ok Then
            My.Settings.Save()
            smtpGroup.Visible = False
            TextBox1.Visible = True
        End If

    End Sub

    Private Sub smtpReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles smtpReset.Click
        smtpHost.Text = Nothing
        smtpPort.Text = 25
        'smtpSsl.Checked = False
    End Sub

    Private Sub smtpCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles smtpCancel.Click
        smtpGroup.Visible = False
        TextBox1.Visible = True
    End Sub

#End Region

#Region "SMTP設定"
    'Private Sub smtpSsl_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles smtpSsl.CheckedChanged
    '    If smtpSsl.Checked Then
    '        userAcc.Enabled = True
    '        userPw.Enabled = True
    '    Else
    '        userAcc.Enabled = False
    '        userAcc.Text = Nothing
    '        userPw.Enabled = False
    '        userPw.Text = Nothing
    '    End If
    'End Sub
#End Region

    Public Function SendMail(ByVal strMailTo As String, ByVal strMailFrom As String, ByVal strSubject As String, _
                                    ByVal strBody As String, Optional ByVal att As System.Net.Mail.Attachment = Nothing, Optional ByVal filePath As String = "") As String
        Dim client As New System.Net.Mail.SmtpClient
        Dim mail As New MailMessage()

        Try

            mail.To.Add(strMailTo)
            mail.From = New System.Net.Mail.MailAddress(strMailFrom)
            mail.SubjectEncoding = System.Text.Encoding.UTF8
            mail.BodyEncoding = System.Text.Encoding.UTF8
            mail.HeadersEncoding = System.Text.Encoding.UTF8
            mail.Subject = strSubject
            mail.Body = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(strBody))

            mail.IsBodyHtml = True


            If att IsNot Nothing Then
                mail.Attachments.Add(att)
            End If

            ' "smtp.syscom.com.tw" "ms2.hinet.net" "smtp.gmail.com" >> 需ssl&帳密
            client.Host = "ms2.hinet.net"
            'client.EnableSsl = False

            client.Timeout = 180000 '三分鐘
            client.Send(mail)
            Return "ok"

        Catch ex As SmtpException

            Dim strErr As String = ""
            Try
            Catch
                strErr = ex.StatusCode
            End Try
            If String.IsNullOrEmpty(strErr) Then
                strErr = ex.Message
            End If

            Return strErr

        Catch ex2 As Exception

            Return ex2.Message

        Finally
            If att IsNot Nothing Then
                att.Dispose()
            End If
        End Try

    End Function


    ''' <summary>
    ''' 改良後的寫法，所有資訊從EmlParser中提取。
    ''' </summary>
    ''' <param name="emlData">emlData類別，含內含所需資訊</param>
    ''' <returns>"ok"表示成功；其餘失敗</returns>
    ''' <remarks></remarks>
    Public Function SendMail(ByRef emlData As EmlParser) As String
        Dim client As New System.Net.Mail.SmtpClient
        Dim mail As New MailMessage()

        Try
            mail.HeadersEncoding = System.Text.Encoding.UTF8
            mail.SubjectEncoding = System.Text.Encoding.UTF8
            mail.BodyEncoding = System.Text.Encoding.UTF8


            If emlData.ArrTo IsNot Nothing AndAlso emlData.ArrTo.Count > 0 Then
                For Each Str As String In emlData.ArrTo
                    mail.To.Add(Str)
                Next
            End If
            If emlData.ArrCC IsNot Nothing AndAlso emlData.ArrCC.Count > 0 Then
                For Each Str As String In emlData.ArrCC
                    mail.CC.Add(Str)
                Next
            End If

            mail.From = New System.Net.Mail.MailAddress(emlData.From)
            mail.Subject = emlData.Subject
            mail.Body = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(emlData.Body))
            '信件優先權
            'mail.Priority = MailPriority.High
            mail.IsBodyHtml = True

            If String.IsNullOrEmpty(emlData.attStrZip) Then
                emlData.att = StreamOutputAttachment(emlData.attStr, emlData.Subject)
            Else
                'eml中已含密碼壓縮檔的寄信
                emlData.att = CreatAttZip(emlData.attStrZip, emlData.Subject)
            End If

            '把附件轉成壓縮檔
            'emlData.att = MakeZipAtt(emlData.attStr)
            '測試有密碼的壓縮檔(sharpZip)
            'emlData.att = MakeSharpZip(emlData.attStr)
            '測試有密碼的壓縮檔(ionic)
            'emlData.att = CreateSampleIonic(emlData.attStr, emlData.Subject, "1234")

            '直接加實體檔案為附件
            'emlData.att = New System.Net.Mail.Attachment("C:\Users\User\Downloads\test.zip", MediaTypeNames.Application.Zip)

            If emlData.att IsNot Nothing Then
                mail.Attachments.Add(emlData.att)
            End If

            'Attachments資訊
            'TextBox1.Text &= "ContentDisposition > " & emlData.att.ContentDisposition.ToString & vbNewLine & vbNewLine
            'TextBox1.Text &= "ContentId > " & emlData.att.ContentId & vbNewLine & vbNewLine
            'TextBox1.Text &= "ContentType > " & emlData.att.ContentType.ToString & vbNewLine & vbNewLine
            'TextBox1.Text &= "NameEncoding > " & emlData.att.NameEncoding.ToString & vbNewLine & vbNewLine
            'TextBox1.Text &= "TransferEncoding > " & emlData.att.TransferEncoding & vbNewLine & vbNewLine
            'TextBox1.Text &= "client.TargetName > " & client.TargetName & vbNewLine & vbNewLine


            'smtp設定---------------

            '"smtp.syscom.com.tw"
            'client.Host = "smtp.syscom.com.tw"
            'client.EnableSsl = False

            '"ms2.hinet.net" 
            'client.Host = "ms2.hinet.net"
            'client.EnableSsl = False

            '"smtp.gmail.com" >> 需ssl&帳密
            'client.Host = "smtp.gmail.com"
            'client.EnableSsl = True
            'client.UseDefaultCredentials = False
            'client.Credentials = New System.Net.NetworkCredential("jiash561@gmail.com", "googlen12369874")

            If Not String.IsNullOrEmpty(My.Settings.SMTPHost) Then 'smtpHost必設
                client.Host = My.Settings.SMTPHost
            Else
                Return "smtp"
            End If

            'Port如果空白的話，給預設值25
            client.Port = IIf(String.IsNullOrEmpty(My.Settings.SMTPPort.ToString.Trim), 25, My.Settings.SMTPPort)
            'client.EnableSsl = My.Settings.nSSL

            'If client.EnableSsl Then
            'client.UseDefaultCredentials = False
            'client.Credentials = New System.Net.NetworkCredential(My.Settings.Acc, My.Settings.Pw)
            client.DeliveryMethod = SmtpDeliveryMethod.Network
            client.Credentials = New System.Net.NetworkCredential(My.Settings.Acc, My.Settings.Pw)
            'End If

            client.Timeout = 180000 '逾時三分鐘
            client.Send(mail)
            Return "ok"

        Catch ex As SmtpException

            Dim strErr As String = ""
            strErr = ex.StatusCode
            If String.IsNullOrEmpty(strErr) Then
                strErr &= " " & ex.Message
            End If

            Return strErr

        Catch ex2 As Exception

            Return ex2.Message

        Finally
            If emlData.Att IsNot Nothing Then
                emlData.Att.Dispose()
            End If
            If client IsNot Nothing Then
                client.Dispose()
            End If
            If mail IsNot Nothing Then
                mail.Dispose()
            End If

        End Try

    End Function

    Public Function StreamOutputAttachment(ByVal content As String, ByVal strFileName As String, Optional ByVal mediaType As String = MediaTypeNames.Text.Plain) As System.Net.Mail.Attachment

        Dim ms As New MemoryStream(System.Text.Encoding.UTF8.GetBytes(content))
        Dim att As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(ms, strFileName & ".txt", MediaTypeNames.Text.Plain)

        att.NameEncoding = System.Text.Encoding.UTF8
        att.TransferEncoding = Net.Mime.TransferEncoding.Base64

        Return att
    End Function

    Public Function MakeZipAtt(ByVal content As String, Optional ByVal mediatype As String = MediaTypeNames.Application.Zip) As System.Net.Mail.Attachment
        '寄出壓縮檔，兩段式可命名到內容檔。
        '使用內建函式庫
        Dim inStream As New MemoryStream(System.Text.Encoding.UTF8.GetBytes(content))
        Dim outstream As New MemoryStream()
        Dim compress As New GZipStream(outstream, CompressionMode.Compress)

        inStream.CopyTo(compress)
        compress.Close()

        'Dim ms As New MemoryStream(outstream.ToArray())
        Using ms As New MemoryStream(outstream.ToArray())
            'Dim att As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(ms, "test.txt.rar", MediaTypeNames.Application.Zip)
            Dim att As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(ms, "test.txt.rar", mediatype)

            att.NameEncoding = System.Text.Encoding.UTF8
            att.TransferEncoding = Net.Mime.TransferEncoding.Base64
            Return att
        End Using

    End Function

#Region "Sample SharpZip"

    Public Function MakeSharpZip(ByVal content As String) As Attachment

        Dim inStream As New MemoryStream(System.Text.Encoding.UTF8.GetBytes(content))
        Dim outStream As New MemoryStream()
        Dim zipStream As New Zip.ZipOutputStream(outStream)

        zipStream.Password = "1234"
        zipStream.SetLevel(1)

        Dim newEntry As New Zip.ZipEntry("textZipEntry.txt")
        newEntry.DateTime = DateTime.Now

        zipStream.PutNextEntry(newEntry)
        Core.StreamUtils.Copy(inStream, zipStream, New Byte(4095) {})
        zipStream.CloseEntry()

        zipStream.IsStreamOwner = False
        zipStream.Close()
        outStream.Position = 0

        Dim att As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(outStream, "456", MediaTypeNames.Text.Plain)

        Return att

    End Function

    Public Sub CreateSample(ByVal outPathName As String, ByVal password As String, ByVal folderName As String)

        'Dim fsOut As FileStream = File.Create("D:\987.zip") 'outPathname) '輸出檔案
        Using fsOut As FileStream = File.Create("D:\987.zip")
            Dim zipStream As New Zip.ZipOutputStream(fsOut)

            folderName = "D:\456" '輸入資料夾

            zipStream.SetLevel(3)       '0-9, 9 being the highest level of compression
            zipStream.Password = "1234"   ' optional. Null is the same as not setting.

            ' This setting will strip the leading part of the folder path in the entries, to
            ' make the entries relative to the starting folder.
            ' To include the full path for each entry up to the drive root, assign folderOffset = 0.
            Dim folderOffset As Integer = folderName.Length + (If(folderName.EndsWith("\"), 0, 1))

            CompressFolder(folderName, zipStream, folderOffset)

            zipStream.IsStreamOwner = True
            ' Makes the Close also Close the underlying stream
            'zipStream.Dispose()
        End Using

    End Sub

    ' Recurses down the folder structure
    Private Sub CompressFolder(ByVal path As String, ByVal zipStream As Zip.ZipOutputStream, ByVal folderOffset As Integer)

        Dim files As String() = Directory.GetFiles(path)

        For Each filename As String In files

            Dim fi As New FileInfo(filename)

            Dim entryName As String = filename.Substring(folderOffset)  ' Makes the name in zip based on the folder
            entryName = Zip.ZipEntry.CleanName(entryName)       ' Removes drive from name and fixes slash direction
            Dim newEntry As New Zip.ZipEntry(entryName)
            newEntry.DateTime = fi.LastWriteTime            ' Note the zip format stores 2 second granularity

            ' Specifying the AESKeySize triggers AES encryption. Allowable values are 0 (off), 128 or 256.
            '   newEntry.AESKeySize = 256;

            ' To permit the zip to be unpacked by built-in extractor in WinXP and Server2003, WinZip 8, Java, and other older code,
            ' you need to do one of the following: Specify UseZip64.Off, or set the Size.
            ' If the file may be bigger than 4GB, or you do not need WinXP built-in compatibility, you do not need either,
            ' but the zip will be in Zip64 format which not all utilities can understand.
            '   zipStream.UseZip64 = UseZip64.Off;
            newEntry.Size = fi.Length

            zipStream.PutNextEntry(newEntry)

            ' Zip the file in buffered chunks
            ' the "using" will close the stream even if an exception occurs
            Dim buffer As Byte() = New Byte(4095) {}
            Using streamReader As FileStream = File.OpenRead(filename)
                Core.StreamUtils.Copy(streamReader, zipStream, buffer)
            End Using
            zipStream.CloseEntry()
        Next

        Dim folders As String() = Directory.GetDirectories(path)
        For Each folder As String In folders
            CompressFolder(folder, zipStream, folderOffset)
        Next
    End Sub
#End Region

#Region "SampleIonic"
    Public Function CreateSampleIonic(ByVal content As String, ByVal strFileName As String, ByVal pw As String) As System.Net.Mail.Attachment

        Using zip As New ZipFile(System.Text.Encoding.UTF8)
            'Dim outputStream As New MemoryStream()
            Using outputStream As New MemoryStream
                zip.Password = pw
                zip.AddEntry(strFileName & ".txt", System.Text.Encoding.UTF8.GetBytes(content))
                zip.Save(outputStream)
                outputStream.Position = 0

                Dim att As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(outputStream, strFileName & ".zip", MediaTypeNames.Text.Plain)
                att.NameEncoding = System.Text.Encoding.UTF8
                att.TransferEncoding = Net.Mime.TransferEncoding.Base64

                Return att
            End Using
        End Using
    End Function

    Private Function ExtractZip(ByVal ZipToUnpack As String) As String

        'Dim UnZipDir As String = Path.GetDirectoryName(ZipToUnpack) + "\" + Split(Path.GetFileName(ZipToUnpack), ".")(0)
        Try
            Dim UnZipDir As String = Split(ZipToUnpack, ".")(0)
            Dim Zips As ZipFile = ZipFile.Read(ZipToUnpack)
            '處理相同檔案的問題
            AddHandler Zips.ExtractProgress, AddressOf handleOverWrite
            Dim e As ZipEntry
            For Each e In Zips
                e.Extract(UnZipDir, ExtractExistingFileAction.InvokeExtractProgressEvent)
            Next

            Return UnZipDir

        Catch ex As Exception
            Return "error"
        End Try

    End Function

#End Region

    ''' <summary>
    ''' Create a Attachment of Zip
    ''' </summary>
    ''' <param name="content">Must use attStrZip Property </param>
    ''' <param name="strFileName">The Name of Zip</param>
    ''' <returns>Attachment</returns>
    ''' <remarks></remarks>
    Public Function CreatAttZip(ByVal content As String, ByVal strFileName As String) As System.Net.Mail.Attachment
        
        Dim arrByte As Byte() = System.Convert.FromBase64String(content)

        'Dim stream As MemoryStream = New MemoryStream(arrByte)
        Using stream As MemoryStream = New MemoryStream(arrByte)

            stream.Position = 0

            Dim att As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(stream, strFileName & ".zip", MediaTypeNames.Application.Zip)
            att.NameEncoding = System.Text.Encoding.UTF8
            att.TransferEncoding = Net.Mime.TransferEncoding.Base64

            Return att
        End Using

    End Function

    Private Sub handleOverWrite(ByVal sender As Object, ByVal e As ExtractProgressEventArgs)
        'If (e.EventType = ZipProgressEventType.Extracting_BeforeExtractEntry) Then

        If e.EventType = ZipProgressEventType.Extracting_ExtractEntryWouldOverwrite Then
            Dim entry As ZipEntry = e.CurrentEntry
            'Dim response As String = Nothing

            Dim msgStyle As Integer = MsgBoxStyle.OkCancel + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2
            Dim msg As String = entry.FileName & vbNewLine & "檔案已存在，是否覆蓋？"
            Dim ret As Integer = MsgBox(msg, msgStyle)
            If ret = MsgBoxResult.Ok Then
                entry.ExtractExistingFile = ExtractExistingFileAction.OverwriteSilently
            Else
                entry.ExtractExistingFile = ExtractExistingFileAction.DoNotOverwrite
            End If

        End If

    End Sub

End Class
