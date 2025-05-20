Imports System.IO
Imports System.Net.NetworkInformation
Imports System.IO.File
Imports System.Windows.Forms

Imports System.Web
Imports System.Management

Imports System.Data.OleDb
Imports System.Net
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class MainFrm

    <DllImport("user32.dll", entrypoint:="FlashWindow")> _
Public Shared Function flashwindow(ByVal hwnd As Integer, ByVal binvert As Integer) As Integer

    End Function

    Private Sub EitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        flashwindow(Me.Handle, 1)
        'End
    End Sub

    'Private Sub EntityConversionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EntityConversionToolStripMenuItem.Click
    '    Try

    '        'Code for database connection

    '        Dim g_strConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DbEntity.mdb" & ";Persist Security Info=False"

    '        Dim con As New OleDbConnection(g_strConnString)

    '        Dim cmd As New OleDbCommand

    '        con.Open()

    '        Dim conn As New OleDbConnection(g_strConnString)
    '        Dim da As OleDb.OleDbDataAdapter
    '        Dim sql As String = Nothing

    '        Try
    '            conn.Open()
    '            sql = "SELECT symbols, characterentity, namedentity FROM EntityList"
    '            Dim ds As New DataSet
    '            da = New OleDb.OleDbDataAdapter(sql, conn)
    '            da.Fill(ds)

    '            For i = 0 To ds.Tables(0).Rows.Count - 1
    '                Dim symbol_val As String = (ds.Tables(0).Rows(i).Item(0))
    '                Dim Character_ent As String = (ds.Tables(0).Rows(i).Item(1))
    '                Dim Named_ent As String = (ds.Tables(0).Rows(i).Item(2))

    '                ' MS_CTRL.Run("Con_Entity", char_val, char_ent
    '            Next
    '        Catch ex As OleDbException
    '            MsgBox("Error: " & ex.ToString())
    '        Finally
    '            conn.Close()
    '        End Try

    '        'Code for inset
    '        'cmd.CommandText = "INSERT INTO table1(Neyms) VALUES('" + var1 + "')"
    '        'cmd.ExecuteNonQuery()
    '        'End

    '        'End by mohamed
    '        'Dim count As Integer = doc.Words.Count

    '        Kill_word_Process()

    '        Dim msword As New Microsoft.Office.Interop.Word.Application

    '        'Code for opened specified file
    '        'msword.Documents.Open("C:\Idries\Epub_Samples\Inputword.rtf")


    '        msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")

    '        '*********Conversion for named to hexadecimal entity************
    '        Dim rng As Word.Range
    '        rng = msword.ActiveDocument.Content
    '        With rng.Find

    '            'For kk As Integer = 1 To 10

    '            .ClearFormatting()
    '            .MatchCase = True
    '            .Execute(FindText:="&aacute;", _
    '            ReplaceWith:="&#x00E1;", _
    '            Replace:=Word.WdReplace.wdReplaceAll)
    '            .Execute(FindText:="©", _
    '            ReplaceWith:="XXXXXX", _
    '            Replace:=Word.WdReplace.wdReplaceAll)
    '            .Execute(FindText:="&Aacute;", _
    '            ReplaceWith:="&#x00C1;", _
    '            Replace:=Word.WdReplace.wdReplaceAll)

    '            'Next

    '        End With
    '        Marshal.ReleaseComObject(rng)


    '        msword.ActiveDocument.Save()

    '        MsgBox("Entity conversion is completed.", MsgBoxStyle.Information, "Word Automation")
    '        'end proces

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Word Automation")
    '    End Try

    'End Sub

    Private Sub TableConversionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            Kill_word_Process()

            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")


            'code for find bold
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Bold = True
            'msword.Selection.Find.Font.BoldBi = True
            'msword.Selection.Find.Font.Italic = True
            ''msword.Selection.Find.Font.ItalicBi = True
            'msword.Selection.Find.Font.Underline = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Bold = True Then
                    .Replacement.Text = "<b>^&</b>"
                    '.Replacement.Text = "<b id=""^&"">^&</b>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If

            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find italic
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Italic = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Italic = True Then
                    .Replacement.Text = "<i>^&</i>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'Code for super & subscript

            'Selection.HomeKey(wdswdStory)
            'With Selection.Find
            '    .ClearFormatting()
            '    .Replacement.ClearFormatting()
            '    .Font.Superscript = True
            '    .Replacement.Text = "^^{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            '    .Font.Subscript = True
            '    .Replacement.Text = "_{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            'End With


            Dim rng As Word.Range
            rng = msword.ActiveDocument.Content

            'msword.ActiveDocument.FormattingShowFont.ToString()

            'rng.Find.Font.Bold = True

            With rng.Find
                '.Font.Italic = True
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Superscript = True
                '.Replacement.Text = "<a id=""fn^&"" href=""#ft^&""><sup>^&</sup></a>" '"^^{^&}"
                .Replacement.Text = "<sup>^&</sup>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Subscript = True
                .Replacement.Text = "<sub>^&</sub>" '"_{^&}"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

            End With

            Marshal.ReleaseComObject(rng)

            'End 

            'code for find underline
            ''Selection.Font.Underline = wdUnderlineSingle
            'msword.Selection.Find.ClearFormatting()
            'msword.Selection.Find.Font.Underline = True
            'msword.Selection.Find.Replacement.ClearFormatting()
            'With msword.Selection.Find
            '    .Text = ""
            '    If WdUnderline.wdUnderlineSingle = True Then
            '        .Replacement.Text = "<u>^&</u>"
            '        .Forward = True
            '        .Wrap = WdFindWrap.wdFindContinue
            '        .Format = True
            '        .MatchCase = False
            '        .MatchWholeWord = False
            '        .MatchWildcards = False
            '        .MatchSoundsLike = False
            '        .MatchAllWordForms = False
            '    ElseIf WdUnderline.wdUnderlineNone = True Then
            '        MsgBox("ok")
            '    End If
            'End With
            'msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            '***************format end

            'Code for table processing
            If msword.ActiveDocument.Tables.Count > 0 Then

                For TblCount As Integer = 1 To msword.ActiveDocument.Tables.Count
                    For RowCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Rows.Count

                        For ColumnCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Columns.Count
                            msword.ActiveDocument.Tables.Item(TblCount).Cell(RowCnt, ColumnCnt).Select()

                            msword.Selection.TypeText(Text:="<table >" & msword.Selection.Text.ToString() & "</table>")
                        Next

                    Next

                    'If msword.ActiveDocument.Tables.Item(i).Rows.Count > 0 Then
                    '    Dim hkjh As String = msword.ActiveDocument.Tables.Item(i).Rows.Item(1).ConvertToText.ToString
                    'End If

                Next

            End If


            msword.ActiveDocument.Save()

            MsgBox("Table and formating conversion is completed.", MsgBoxStyle.Information, "Word Automation")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Word Automation")
        End Try
    End Sub

    Friend Sub Kill_word_Process()
        Try
            Dim ExistingProcess() As Process
            Dim ChkProcess As Process
            ExistingProcess = Process.GetProcessesByName("winword")
            For Each ChkProcess In ExistingProcess
                If (LCase(ChkProcess.ProcessName) = "winword") And ChkProcess.MainWindowTitle = "" Then
                    ChkProcess.Kill()
                End If
            Next
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub EpubConversionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        FrmMetaInfo.ShowDialog()

    End Sub



    Private Sub MainFrm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'ComboBox1.Items.Add("xxxx")
        'ComboBox1.Items.Add("yyyyy")
        'ComboBox1.Text = "khh"

        'ComboBox1.SelectedIndex = 1

        'ex.StackTrace 
        Try
            'Dim id As Guid = Guid.NewGuid()          '8db23ba4-687d-4956-a9fa-2f32bc9d4bf4
            ''
            'MsgBox(id.ToString)
            ' flashwindow(Me.Handle, 1)

            'FrmFlash.Show()
            Dim PhysicalAddrList As New ArrayList
            PhysicalAddrList.Add("2AD224144265") 'Edreethe
            PhysicalAddrList.Add("1078D2D5C6D0")
            PhysicalAddrList.Add("001FD067E8AF")
            PhysicalAddrList.Add("00E04C448338")
            PhysicalAddrList.Add("00E04C44815B")
            PhysicalAddrList.Add("00E04C4480C1")
            PhysicalAddrList.Add("00E04C448158")
            PhysicalAddrList.Add("00FF586D583D")
            PhysicalAddrList.Add("00E04C2EEFE2")




            Dim PhysicalAddr As String = ""
            For Each nic As NetworkInformation.NetworkInterface In NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                'MessageBox.Show(String.Format("The MAC address of {0} is{1}{2}", nic.Description, Environment.NewLine, nic.GetPhysicalAddress())) '2AD224144265
                PhysicalAddr = nic.GetPhysicalAddress.ToString
                Exit For

            Next

            Dim PhysicalAddrFound As Boolean = False
            For Each Addr In PhysicalAddrList
                If Regex.IsMatch(Addr.ToString, "^" & PhysicalAddr & "$", RegexOptions.IgnoreCase) Then
                    PhysicalAddrFound = True
                End If
            Next


            If PhysicalAddrFound = False Then
                MsgBox("Unauthorized user. Please Contact Administrator.", MsgBoxStyle.Exclamation, "Access denied")
                End
            End If

            'Dim cpuID As String = String.Empty
            'Dim mc As ManagementCla = New ManagementClass("Win32_NetworkAdapterConfiguration")
            'Dim moc As ManagementObjectCollection = mc.GetInstances()
            'For Each mo As ManagementObject In moc
            '    If (cpuID = String.Empty And CBool(mo.Properties("IPEnabled").Value) = True) Then
            '        cpuID = mo.Properties("MacAddress").Value.ToString()
            '    End If
            'Next
            ' Return cpuID


            'Code for get ipAddress
            'Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName)
            'Dim kj As String = h.AddressList.GetValue(0).ToString
            'End

            '54BEF71306D3 My machine physical address 48D224144265
            'Dim PhysicalAddrList As New ArrayList
            'Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
            'Dim PhysicalAddrList As New ArrayList
            'PhysicalAddrList.Add("54BEF71306D3")
            'PhysicalAddrList.Add("1078D2D5C6D0")
            'PhysicalAddrList.Add("00E04C448338")
            'PhysicalAddrList.Add("00E04C44815B")
            'PhysicalAddrList.Add("00E04C4480C1")
            'PhysicalAddrList.Add("00E04C448158")
            'PhysicalAddrList.Add("001FD067E8AF")

            'PhysicalAddrList.Add("1078D2D5C6D0")

            'MsgBox(nics(2).GetPhysicalAddress.ToString)
            'Dim PhysicalAddrFound As Boolean = False
            'For Each Addr In PhysicalAddrList
            '    If Regex.IsMatch(Addr.ToString, "^" & nics(2).GetPhysicalAddress.ToString & "$", RegexOptions.IgnoreCase) Then
            '        PhysicalAddrFound = True
            '    End If
            'Next


            'If PhysicalAddrFound = False Then
            '    MsgBox("Unauthorized user. Please Contact Administrator.", MsgBoxStyle.Exclamation, "Access denied")
            '    End
            'End If

            'End process



            Dim App_path As String = System.Windows.Forms.Application.StartupPath

            'stylesheet.css
            'language.ini
            If Not File.Exists(App_path & "\supportingFiles\language.ini") Then
                MsgBox("Supporting file 'language.ini' is missing. Please contact the developer.", MsgBoxStyle.Exclamation, "Epub conversion")
                End
                'DbEntity
            ElseIf Not File.Exists(App_path & "\DbEntity.mdb") Then
                MsgBox("Supporting file 'DbEntity.mdb' is missing. Please contact the developer.", MsgBoxStyle.Exclamation, "Epub conversion")
                End

            ElseIf Not File.Exists(App_path & "\supportingFiles\stylesheet.css") Then
                MsgBox("Supporting file 'stylesheet.css' is missing. Please contact the developer.", MsgBoxStyle.Exclamation, "Epub conversion")
                End

            ElseIf Not File.Exists(App_path & "\supportingFiles\container.xml") Then
                MsgBox("Supporting file 'container.xml' is missing. Please contact the developer.", MsgBoxStyle.Exclamation, "Epub conversion")
                End

            ElseIf Not File.Exists(App_path & "\supportingFiles\backcover.ini") Then
                MsgBox("Supporting file 'backcover.ini' is missing. Please contact the developer.", MsgBoxStyle.Exclamation, "Epub conversion")
                End
            ElseIf Not File.Exists(App_path & "\supportingFiles\cover.ini") Then
                MsgBox("Supporting file 'cover.ini' is missing. Please contact the developer.", MsgBoxStyle.Exclamation, "Epub conversion")
                End
            ElseIf Not File.Exists(App_path & "\supportingFiles\chaptimageContent.ini") Then
                MsgBox("Supporting file 'chaptimageContent.ini' is missing. Please contact the developer.", MsgBoxStyle.Exclamation, "Epub conversion")
                End
            ElseIf Not File.Exists(App_path & "\office.dll") Then
                MsgBox("Supporting file 'office.dll' is missing. Please contact the developer.", MsgBoxStyle.Exclamation, "Epub conversion")
                End
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Epub conversion")
        End Try

    End Sub


    Private Sub Symbol2CharacterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

            'Code for database connection

            Dim g_strConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DbEntity.mdb" & ";Persist Security Info=False"

            Dim con As New OleDbConnection(g_strConnString)

            Dim cmd As New OleDbCommand

            con.Open()

            Dim conn As New OleDbConnection(g_strConnString)
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String = Nothing

            ' Try
            conn.Open()
            sql = "SELECT symbols, characterentity FROM EntityList"
            Dim ds As New DataSet
            da = New OleDb.OleDbDataAdapter(sql, conn)
            da.Fill(ds)


            'For doc file process
            Kill_word_Process()
            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")
            'End

            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim symbol_val As String = String.Empty
                Try
                    symbol_val = (ds.Tables(0).Rows(i).Item(0))
                Catch ex As Exception
                End Try
                Dim Character_ent As String = String.Empty
                Try
                    Character_ent = (ds.Tables(0).Rows(i).Item(1))
                Catch ex As Exception
                End Try

                'Dim Named_ent As String = (ds.Tables(0).Rows(i).Item(2))

                If Not symbol_val = Nothing AndAlso Not Character_ent = Nothing Then
                    FunEntityConvn(symbol_val, Character_ent)
                End If

                ' MS_CTRL.Run("Con_Entity", char_val, char_ents
            Next

            'cmd.Connection = conn
            'cmd.CommandText = "INSERT INTO EntityList(symbols) VALUES('" + "txttext" + "')"
            'cmd.ExecuteNonQuery()


            'save word doc
            msword.ActiveDocument.Save()

            MsgBox("Symbol to character entity is converted successfully.", MsgBoxStyle.Information, "Entity conversion")

            'Catch ex As OleDbException
            '    MsgBox("Error: " & ex.ToString())
            'Finally
            conn.Close()
            'End Try

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())

        End Try

    End Sub


    Private Sub Symbol2NamedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

            'Code for database connection

            Dim g_strConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DbEntity.mdb" & ";Persist Security Info=False"

            Dim con As New OleDbConnection(g_strConnString)

            Dim cmd As New OleDbCommand

            con.Open()

            Dim conn As New OleDbConnection(g_strConnString)
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String = Nothing

            Try
                conn.Open()
                sql = "SELECT symbols, namedentity FROM EntityList"
                Dim ds As New DataSet
                da = New OleDb.OleDbDataAdapter(sql, conn)
                da.Fill(ds)

                'For doc file process
                Kill_word_Process()
                Dim msword As New Microsoft.Office.Interop.Word.Application
                msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")
                'End
                Dim symbol_val As String = String.Empty
                Dim Named_ent As String = String.Empty
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    Try
                        symbol_val = (ds.Tables(0).Rows(i).Item(0))
                    Catch ex As Exception
                    End Try
                    'Dim Character_ent As String = (ds.Tables(0).Rows(i).Item(1))
                    Try
                        Named_ent = (ds.Tables(0).Rows(i).Item(1))
                    Catch ex As Exception
                    End Try
                    If Not symbol_val = Nothing AndAlso Not Named_ent = Nothing Then
                        FunEntityConvn(symbol_val, Named_ent)
                    End If
                    ' MS_CTRL.Run("Con_Entity", char_val, char_ent
                Next

                'save word doc
                msword.ActiveDocument.Save()

                MsgBox("Symbol to named entity is converted successfully.", MsgBoxStyle.Information, "Entity Conversion")

            Catch ex As OleDbException
                MsgBox("Error: " & ex.ToString())
            Finally
                conn.Close()
            End Try

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try

    End Sub

    Public Function FunEntityConvn(ByVal Symbol As String, ByVal ConvertingEntity As String)

        Try

            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")

            Dim rng As Word.Range
            rng = msword.ActiveDocument.Content
            
            With rng.Find
                .ClearFormatting()
                .MatchCase = True
                .Execute(FindText:=Symbol, _
                ReplaceWith:=ConvertingEntity, _
                Replace:=Word.WdReplace.wdReplaceAll)
            End With

            Marshal.ReleaseComObject(rng)

            ' msword.ActiveDocument.Save()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return ""
    End Function

    Private Sub EntityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        FrmEntityAdd.ShowDialog()
    End Sub




    Private Sub WebLinkConversionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Code for replace text


    End Sub




    Private Sub Symbol2CharacterToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Symbol2NamedToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub EntityToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub WebLinkConversionToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Kill_word_Process()

        Dim msword As New Microsoft.Office.Interop.Word.Application
        msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")
        Dim rng As Word.Range
        rng = msword.ActiveDocument.Content

        'Start
        Dim ActiveDoc_Content As String = msword.ActiveDocument.Content.Text
        Dim TempOverlapList As New ArrayList


        Dim WEBPatrn As MatchCollection = Regex.Matches(ActiveDoc_Content, " (.*?)(\.com|\.org|\\.net|\.in) ", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        For Each Patrn As Match In WEBPatrn

            msword.ActiveDocument.FormattingShowFont.ToString()

            rng.Find.Font.Bold = True

            If Not TempOverlapList.Contains(Patrn.ToString) Then
                With rng.Find
                    '.Font.Italic = True
                    .ClearFormatting()
                    .Execute(FindText:=Patrn.ToString, _
                    ReplaceWith:="*******xxxxxtbl2****", _
                    Replace:=Word.WdReplace.wdReplaceAll)
                    TempOverlapList.Add(Patrn.ToString)
                End With
            End If

        Next

        'End


        Marshal.ReleaseComObject(rng)
        msword.ActiveDocument.Save()

        MsgBox("WebLink process is completed.", MsgBoxStyle.Information, "WebLink Automation")
    End Sub



    Private Sub EpubConversionToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FrmMetaInfo.ShowDialog()
    End Sub

    Private Sub EitToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        End
    End Sub


    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.Close()
    End Sub

    'Private Sub EntityToolStripMenuItem_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EntityToolStripMenuItem.Click
    '    FrmEntityAdd.ShowDialog()
    'End Sub

    Private Sub EpubConversionToolStripMenuItem_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EpubConversionToolStripMenuItem.Click
        FrmMetaInfo.Show()
    End Sub

   Private Sub TableConversionToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TableConversionToolStripMenuItem.Click
        Try
            'Dim nnh As String = ComboBox1.SelectedText.ToString()


            'nnh = ComboBox1.Items(0).ToString()

            'If ComboBox1.CanFocus() = True Then
            'MsgBox("k")
            'End If

            'Me.WindowState = FormWindowState.Minimized

            Kill_word_Process()

            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")

            'Code for remove field codes
            'msword.ActiveDocument.Fields.Unlink()

            'initial replace for >,< and & entity


            'msword.Selection.Find.ClearFormatting()
            'msword.Selection.Find.Replacement.ClearFormatting()
            'msword.Selection.End = 0
            'With msword.Selection.Find
            '    .Text = "&"
            '    .Replacement.Text = "&amp;"
            '    .Forward = True
            '    .Wrap = WdFindWrap.wdFindContinue
            '    .Format = False
            '    .MatchCase = False
            '    .MatchWholeWord = False
            '    .MatchWildcards = False
            '    .MatchSoundsLike = False
            '    .MatchAllWordForms = False
            'End With
            'msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'msword.Selection.Find.ClearFormatting()
            'msword.Selection.Find.Replacement.ClearFormatting()
            'msword.Selection.End = 0
            'With msword.Selection.Find
            '    .Text = ">"
            '    .Replacement.Text = "&gt;"
            '    .Forward = True
            '    .Wrap = WdFindWrap.wdFindContinue
            '    .Format = False
            '    .MatchCase = False
            '    .MatchWholeWord = False
            '    .MatchWildcards = False
            '    .MatchSoundsLike = False
            '    .MatchAllWordForms = False
            'End With
            'msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'msword.Selection.Find.ClearFormatting()
            'msword.Selection.Find.Replacement.ClearFormatting()
            'msword.Selection.End = 0
            'With msword.Selection.Find
            '    .Text = "<"
            '    .Replacement.Text = "&lt;"
            '    .Forward = True
            '    .Wrap = WdFindWrap.wdFindContinue
            '    .Format = False
            '    .MatchCase = False
            '    .MatchWholeWord = False
            '    .MatchWildcards = False
            '    .MatchSoundsLike = False
            '    .MatchAllWordForms = False
            'End With
            'msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
            'Initial replacement end

            'Initial cleanup process
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "  "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = vbCr & vbCr
                .Replacement.Text = vbCr
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^t"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^b"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^m"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            'To remove alphabet and number for increase speed to find only needed symbols
            Dim WordContent As String = msword.ActiveDocument.Content.Text.ToString
            WordContent = Regex.Replace(WordContent, "[@|$|%|\*|!|a-z|0-9| |;|\.|\:|,|<|>|\/|\(|\-|\)]+", "", RegexOptions.IgnoreCase Or RegexOptions.Singleline)


            ' Dim jkjk As String = Regex.Replace("!!!!@!!", "!{4}", "", RegexOptions.IgnoreCase)
            'Process for entity conversion
            ProgressBar1.Visible = True
            ProgressBar1.Minimum = 0
            'ProgressBar1.Maximum = msword.ActiveDocument.Content.Text.Count
            ProgressBar1.Maximum = WordContent.ToString.Count
            'Test number entity
            Dim rngs As Word.Range

            'Iterate through each character
            'For Each c As Char In msword.ActiveDocument.Content.Text
            For Each c As Char In WordContent.ToString

                rngs = msword.ActiveDocument.Content
                If AscW(c) > 127 Then '255 Then '127 Then '

                    With rngs.Find
                        .ClearFormatting()
                        .MatchCase = True
                        .Execute(FindText:=c.ToString, _
                        ReplaceWith:=String.Format("&#{0}", AscW(c) & ";"), _
                        Replace:=Word.WdReplace.wdReplaceAll)
                    End With

                End If

                If ProgressBar1.Value > WordContent.ToString.Count Then
                    ProgressBar1.Value = ProgressBar1.Value - 10
                Else
                    ProgressBar1.Value = ProgressBar1.Value + 1
                End If
            Next

            Marshal.ReleaseComObject(rngs)

            'Code for replace bold tag
            msword.Selection.End = 0
            With msword.Selection.Find
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Bold = True
                .Replacement.Font.Bold = False
                .Execute(FindText:="", ReplaceWith:="<strong>^&</strong>", MatchWildcards:=True, Replace:=Word.WdReplace.wdReplaceAll)
            End With


            'Code for insert proper <Strong> closing and opening tag
            Dim rnge As Word.Range
            rnge = msword.ActiveDocument.Content

            msword.ActiveDocument.FormattingShowFont.ToString()

            rnge.Find.Font.Bold = True

            With rnge.Find 'Replace the unstructured </strong> tag
                .ClearFormatting()
                .Execute(FindText:=vbCr & "</strong>", _
                ReplaceWith:="</strong>" & vbCr, _
                Replace:=Word.WdReplace.wdReplaceAll)
            End With

            'rnge = msword.ActiveDocument.Content

            'msword.ActiveDocument.FormattingShowFont.ToString()

            'rnge.Find.Font.Bold = True

            'With rnge.Find 'Replace the unstructured </strong> tag
            '    .ClearFormatting()
            '    .Execute(FindText:="<strong>*</strong>", _
            '    ReplaceWith:="</strong>^p<strong>", _
            '    Replace:=Word.WdReplace.wdReplaceAll)
            'End With



            Dim StrongTag As String = msword.ActiveDocument.Content.Text.ToString
            Dim StrongTags As MatchCollection = Regex.Matches(StrongTag, "<strong>(((?!</strong>).)+)</strong>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            Dim TmpList As New ArrayList

            For Each STag As Match In StrongTags


                Try


                    Dim Entermark() As String = Regex.Split(STag.ToString, vbCr)
                    If Entermark.Count > 0 Then
                        For Each item As String In Entermark

                            If Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then

                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:=item.ToString & "</strong>", _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With

                                    TmpList.Add(item.ToString)

                                End If
                            ElseIf Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then
                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:="<strong>" & item.ToString, _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With
                                    TmpList.Add(item.ToString)
                                End If
                            ElseIf Not Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then
                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:="<strong>" & item.ToString & "</strong>", _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With
                                    TmpList.Add(item.ToString)
                                End If

                            End If

                        Next
                    End If
                Catch ex As Exception
                    'MsgBox(ex.Message)
                End Try
            Next


            Marshal.ReleaseComObject(rnge)


            'code for find italic
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Italic = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Italic = True Then
                    .Replacement.Text = "<em>^&</em>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find underline
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Underline = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Underline = WdUnderline.wdUnderlineSingle Then
                    .Replacement.Text = "<span class=" & ChrW(34) & "underline" & ChrW(34) & ">^&</span>"
                    '"<strong>^&</strong>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            'code for find strikethrough
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.StrikeThrough = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.StrikeThrough = True Then
                    .Replacement.Text = "<span class=" & ChrW(34) & "strike" & ChrW(34) & ">^&</span>"
                    '"<strong>^&</strong>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find smallcaps
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.SmallCaps = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.SmallCaps = True Then
                    .Replacement.Text = "<small>^&</small>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find image
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = "^g"
                .Replacement.Text = "<img></img>"
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            Dim ImgTag As MatchCollection = Regex.Matches(msword.ActiveDocument.Content.Text, "<img></img>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)


            Dim FileSeqcount As Integer = 0
            Dim FileSeqVar As String = ""

            For Each Img As Match In ImgTag

                'Code for append initial zero sequence
                FileSeqcount = FileSeqcount + 1
                If Regex.IsMatch(FileSeqcount, "^\d{1}$", RegexOptions.None) Then
                    FileSeqVar = "000" & FileSeqcount
                ElseIf Regex.IsMatch(FileSeqcount, "^\d{2}$", RegexOptions.None) Then
                    FileSeqVar = "00" & FileSeqcount
                ElseIf Regex.IsMatch(FileSeqcount, "^\d{3}$", RegexOptions.None) Then
                    FileSeqVar = "0" & FileSeqcount
                Else
                    FileSeqVar = FileSeqcount
                End If

                Dim rng_Img As Word.Range
                rng_Img = msword.ActiveDocument.Content

                With rng_Img.Find
                    .ClearFormatting()

                    .Execute(FindText:=Img.ToString, _
                  ReplaceWith:="<p class=" & ChrW(34) & "image" & ChrW(34) & "><img src=" & ChrW(34) & "images/f" & FileSeqVar & ".jpg" & ChrW(34) & " alt=" & ChrW(34) & "image" & ChrW(34) & "/></p>", _
                  Replace:=Word.WdReplace.wdReplaceOne)

                End With

            Next

            'Code for super & subscript

            'Selection.HomeKey(wdswdStory)
            'With Selection.Find
            '    .ClearFormatting()
            '    .Replacement.ClearFormatting()
            '    .Font.Superscript = True
            '    .Replacement.Text = "^^{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            '    .Font.Subscript = True
            '    .Replacement.Text = "_{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            'End With


            'Super & supscript
            Dim rng As Word.Range
            rng = msword.ActiveDocument.Content
            With rng.Find
                '.Font.Italic = True
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Superscript = True
                '.Replacement.Text = "<a id=""fn^&"" href=""#ft^&""><sup>^&</sup></a>" 
                .Replacement.Text = "<sup>^&</sup>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Subscript = True
                .Replacement.Text = "<sub>^&</sub>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

            End With

            Marshal.ReleaseComObject(rng)

            'Code for table processing
            If msword.ActiveDocument.Tables.Count > 0 Then

                Try

                    For TblCount As Integer = 1 To msword.ActiveDocument.Tables.Count
                        For RowCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Rows.Count

                            For ColumnCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Columns.Count
                                msword.ActiveDocument.Tables.Item(TblCount).Cell(RowCnt, ColumnCnt).Select()

                                msword.Selection.TypeText(Text:="<table>" & msword.Selection.Text.ToString() & "</table>")
                            Next

                        Next

                        'If msword.ActiveDocument.Tables.Item(i).Rows.Count > 0 Then
                        '    Dim hkjh As String = msword.ActiveDocument.Tables.Item(i).Rows.Item(1).ConvertToText.ToString
                        'End If

                    Next

                Catch ex As Exception

                End Try
            End If


            'msword.Selection.Find.ClearFormatting()
            'msword.Selection.Find.Replacement.ClearFormatting()
            'msword.Selection.End = 0
            'With msword.Selection.Find
            '    .Text = vbCrLf & vbCrLf
            '    .Replacement.Text = vbCrLf
            '    .Forward = True
            '    .Wrap = WdFindWrap.wdFindContinue
            '    .Format = False
            '    .MatchCase = False
            '    .MatchWholeWord = False
            '    .MatchWildcards = False
            '    .MatchSoundsLike = False
            '    .MatchAllWordForms = False
            'End With
            'msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            'End

            'To replace <strong> tag into normal formatting
            With msword.ActiveDocument.Content.Find
                .ClearFormatting()
                .Font.Bold = True
                With .Replacement
                    .ClearFormatting()
                    .Font.Bold = False
                End With
                .Execute(FindText:="<strong>", ReplaceWith:="<strong>", _
                    Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
            End With

            msword.ActiveDocument.Save()

            Me.Refresh()

            Me.WindowState = FormWindowState.Normal


            MsgBox("Table and formating conversion is completed.", MsgBoxStyle.Information, "Word Automation")

            ProgressBar1.Value = 0
            ProgressBar1.Visible = False


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Word Automation")
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
        End Try

    End Sub

    Private Sub WebLinkConversionToolStripMenuItem_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WebLinkConversionToolStripMenuItem.Click

        Try

            Kill_word_Process()

            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")
            Dim rng As Word.Range
            rng = msword.ActiveDocument.Content

           
            Dim ActiveDoc_Content As String = msword.ActiveDocument.Content.Text

            Dim TempOverlapList As New ArrayList

            'Code for replace Actual hyperlink
            'comment by mohamed for testing
            For i As Integer = 1 To rng.Hyperlinks.Count
                Try

                    rng.Hyperlinks(1).Delete()
                    '        With rng.Find

                    '            If Not TempOverlapList.Contains(rng.Hyperlinks(i).TextToDisplay.ToString()) Then
                    '                .ClearFormatting()
                    '                TempOverlapList.Add(rng.Hyperlinks(i).TextToDisplay.ToString)
                    '                rng.Hyperlinks(i).Name.ToString()
                    '                .Execute(FindText:=rng.Hyperlinks(i).TextToDisplay.ToString(), _
                    '                      ReplaceWith:="<a href=" & ChrW(34) & rng.Hyperlinks(i).Name.ToString() & ChrW(34) & ">" & rng.Hyperlinks(i).TextToDisplay.ToString() & "</a> ", _
                    '                      Replace:=Word.WdReplace.wdReplaceOne)
                    '            End If

                    '        End With
                Catch ex As Exception
                End Try
            Next

                    'upto Space or entermark (?:http://)?(www\.)(.*?)[ |^p]
            Dim WEBPatrn As MatchCollection = Regex.Matches(ActiveDoc_Content, "((?:http://)?(?:www\.(?:(?:(?!\.pdf|\.html|\.com|\.org|\.net|\.de|\.in|\.ru|\.au|\.da|\.ms).)+)(?:\.pdf|\.html|\.com|\.org|\.net|\.de|\.in|\.ru|\.au|\.da|\.ms)+))([ |" & vbCrLf & "]+)", RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.Multiline)

                    ''<p>(((?!</p>).)+)</p>
                    For Each Patrn As Match In WEBPatrn

                        Try
                            If Not TempOverlapList.Contains(Patrn.ToString) Then
                                rng = msword.ActiveDocument.Content

                                With rng.Find
                                    .ClearFormatting()
                                    .Execute(FindText:=Patrn.ToString, _
                                  ReplaceWith:="<a href=" & ChrW(34) & "http://" & Patrn.Groups(1).ToString & ChrW(34) & ">" & Patrn.Groups(1).ToString & "</a>" & Patrn.Groups(2).ToString, _
                                  Replace:=Word.WdReplace.wdReplaceAll)
                                    TempOverlapList.Add(Patrn.ToString)
                                End With

                            End If
                        Catch ex As Exception
                        End Try

                    Next

            'Comment by mohamed
                    'Dim WEBPatrn As MatchCollection = Regex.Matches(ActiveDoc_Content, "((?:http://)?(www\.(((?!\.pdf|\.com|\.org|\.net|\.de|\.in|\.ru|\.au|\.da|\.ms).)+)(\.pdf|\.com|\.org|\.net|\.de|\.in|\.ru|\.au|\.da|\.ms)+))(?:\/?[^\w]+)", RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.Multiline)
            ' ''<p>(((?!</p>).)+)</p>
                    'For Each Patrn As Match In WEBPatrn

                    '    Try
                    '        If Not TempOverlapList.Contains(Patrn.ToString) Then
                    '            rng = msword.ActiveDocument.Content

                    '            With rng.Find
                    '                .ClearFormatting()
                    '                .Execute(FindText:=Patrn.ToString, _
                    '              ReplaceWith:="<a href=" & ChrW(34) & "http://" & Patrn.Groups(1).ToString & ChrW(34) & ">" & Patrn.Groups(1).ToString & "</a>", _
                    '              Replace:=Word.WdReplace.wdReplaceAll)
                    '                TempOverlapList.Add(Patrn.ToString)
                    '            End With

                    '        End If
                    '    Catch ex As Exception
                    '    End Try

                    'Next
            Marshal.ReleaseComObject(rng)
                    msword.ActiveDocument.Save()

                    MsgBox("WebLink process is completed.", MsgBoxStyle.Information, "WebLink Automation")

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

    End Sub


    Private Sub EitToolStripMenuItem_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EitToolStripMenuItem.Click
        Me.Close()
    End Sub

  Private Sub FootNoteConversionToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FootNoteConversionToolStripMenuItem1.Click
        FrmFootNote.Show()

    End Sub

    Private Sub FToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FToolStripMenuItem.Click

        Dim InputContent As String = String.Empty
        Dim TempList As New ArrayList

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            If Regex.IsMatch(Path.GetExtension(OpenFileDialog1.FileName), "html", RegexOptions.IgnoreCase) Then
                InputContent = File.ReadAllText(OpenFileDialog1.FileName)

                Dim Fig As MatchCollection = Regex.Matches(InputContent, "(fig[?:ures]* (\d+\-*\.*(\d*)))", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                For Each figure As Match In Fig
                    '<a href="#fig1">figure 1</a>
                    If Not TempList.Contains(figure.ToString) Then

                        If Regex.IsMatch(figure.ToString, "\d+[\.|\-]+\d+", RegexOptions.IgnoreCase) Then
                            'InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#fig" & figure.Groups(3).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#fig" & figure.Groups(2).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            TempList.Add(figure.ToString)
                        Else
                            'InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#fig" & figure.Groups(2).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#fig" & figure.Groups(2).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            TempList.Add(figure.ToString)
                        End If

                       End If
                Next

                '<a id="fig1"/>
                Dim ReplaceFig As String = String.Empty

                Dim FigCaption As MatchCollection = Regex.Matches(InputContent, "<p class=""caption(((?!</p>).)+)</p>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                For Each Captn As Match In FigCaption

                    'Dim FigMatch As MatchCollection = Regex.Matches(Captn.ToString, "<a href=""#fig(\d+\.*\d*)(((?!</a>).)+)</a>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                    Dim FigMatch As MatchCollection = Regex.Matches(Captn.ToString, "<a href=""#fig(\d+\.*\-*\d*)"">(((?!</a>).)+)</a>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                    '<p class="caption"><a href="#fig1">figure 1</a></p>
                    For Each Ahref As Match In FigMatch
                        ReplaceFig = Regex.Replace(Captn.ToString, "<a href=""#fig(\d+\.*\-*\d*)"">(((?!</a>).)+)</a>", "<a id=""fig$1""" & "/>$2", RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                        InputContent = Regex.Replace(InputContent, Captn.ToString, ReplaceFig, RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                    Next

                Next


                File.WriteAllText(OpenFileDialog1.FileName, InputContent)

                MsgBox("Figure citation is completed.", MsgBoxStyle.Information, "Figure citation")

            Else
                MsgBox("It allows html or xhtml file. Please check.", MsgBoxStyle.Information, "Figure citation")
            End If

        End If

    End Sub

  
    Private Sub TableCitationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TableCitationToolStripMenuItem.Click
        Dim InputContent As String = String.Empty
        Dim TempList As New ArrayList

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            If Regex.IsMatch(Path.GetExtension(OpenFileDialog1.FileName), "html", RegexOptions.IgnoreCase) Then
                InputContent = File.ReadAllText(OpenFileDialog1.FileName)

                Dim Fig As MatchCollection = Regex.Matches(InputContent, "(tab[?:les]* (\d+\-\.*(\d*)))", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                For Each figure As Match In Fig
                    '<a href="#fig1">figure 1</a>
                    If Not TempList.Contains(figure.ToString) Then

                        If Regex.IsMatch(figure.ToString, "\d+[\.|\-]+\d+", RegexOptions.IgnoreCase) Then
                            'InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#tab" & figure.Groups(3).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#tab" & figure.Groups(2).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                            TempList.Add(figure.ToString)
                        Else
                            InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#tab" & figure.Groups(2).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            TempList.Add(figure.ToString)
                        End If

                    End If

                Next

                '<a id="fig1"/>
                Dim ReplaceFig As String = String.Empty

                Dim FigCaption As MatchCollection = Regex.Matches(InputContent, "<p class=""caption(((?!</p>).)+)</p>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                For Each Captn As Match In FigCaption

                    Dim FigMatch As MatchCollection = Regex.Matches(Captn.ToString, "<a href=""#tab(\d+\-*\.*\d*)(((?!</a>).)+)</a>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                    For Each Ahref As Match In FigMatch

                        ReplaceFig = Regex.Replace(Captn.ToString, "<a href=""#tab(\d+\.*\-*\d*)"">(((?!</a>).)+)</a>", "<a id=""tab$1""" & "/>$2", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                        InputContent = Regex.Replace(InputContent, Captn.ToString, ReplaceFig, RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                    Next

                    File.WriteAllText(OpenFileDialog1.FileName, InputContent)


                Next
                MsgBox("Table citataion is completed.", MsgBoxStyle.Information, "Exhibit citation")
            Else
                MsgBox("It allows html or xhtml file. Please check.", MsgBoxStyle.Information, "Table citation")
            End If

        End If

    End Sub

    Private Sub ExhibitCitataionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExhibitCitataionToolStripMenuItem.Click
        Dim InputContent As String = String.Empty
        Dim TempList As New ArrayList

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            If Regex.IsMatch(Path.GetExtension(OpenFileDialog1.FileName), "html", RegexOptions.IgnoreCase) Then
                InputContent = File.ReadAllText(OpenFileDialog1.FileName)

                Dim Fig As MatchCollection = Regex.Matches(InputContent, "(exhibit[?:s]* (\d+\-*\.*(\d*)))", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                For Each figure As Match In Fig
                    '<a href="#fig1">figure 1</a>
                    If Not TempList.Contains(figure.ToString) Then
                      
                        If Regex.IsMatch(figure.ToString, "\d+[\.|\-]+\d+", RegexOptions.IgnoreCase) Then
                            'InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#exh" & figure.Groups(3).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#exh" & figure.Groups(2).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            TempList.Add(figure.ToString)
                        Else
                            InputContent = Regex.Replace(InputContent, figure.ToString, "<a href=""#exh" & figure.Groups(2).ToString & """>" & figure.Groups(1).ToString & "</a>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                            TempList.Add(figure.ToString)
                        End If

                    End If
                Next

                '<a id="fig1"/>
                Dim ReplaceFig As String = String.Empty

                Dim FigCaption As MatchCollection = Regex.Matches(InputContent, "<p class=""caption(((?!</p>).)+)</p>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                For Each Captn As Match In FigCaption

                    Dim FigMatch As MatchCollection = Regex.Matches(Captn.ToString, "<a href=""#exh(\d+\.*\-*\d*)"">(((?!</a>).)+)</a>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                    For Each Ahref As Match In FigMatch
                        ReplaceFig = Regex.Replace(Captn.ToString, "<a href=""#exh(\d+\.*\-*\d*)"">(((?!</a>).)+)</a>", "<a id=""exh$1""" & "/>$2", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                        InputContent = Regex.Replace(InputContent, Captn.ToString, ReplaceFig, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                    Next

                Next

                File.WriteAllText(OpenFileDialog1.FileName, InputContent)


                MsgBox("Exhibit citation completed.", MsgBoxStyle.Information, "Exhibit citation")
            Else
                MsgBox("It allows html or xhtml file. Please check.", MsgBoxStyle.Information, "Exhibit citation")
            End If

        End If

    End Sub

    'Coment by mohamed number entity
    'Private Sub NoEntityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NoEntityToolStripMenuItem.Click
    '    Try

    '        Kill_word_Process()

    '        Dim msword As New Microsoft.Office.Interop.Word.Application
    '        msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")

    '        Dim rng As Word.Range

    '        'Process for entity conversion

    '        'Iterate through each character
    '        For Each c As Char In msword.ActiveDocument.Content.Text
    '            rng = msword.ActiveDocument.Content
    '            If AscW(c) > 127 Then '255 Then '127 Then '

    '                With rng.Find

    '                    Dim myEncodedString As String = HttpUtility.HtmlEncode("")

    '                    .ClearFormatting()
    '                    .MatchCase = True

    '                    .Execute(FindText:=c.ToString, _
    '                    ReplaceWith:=String.Format("&#{0}", AscW(c) & ";"), _
    '                    Replace:=Word.WdReplace.wdReplaceAll)
    '                End With

    '            End If
    '        Next

    '        Marshal.ReleaseComObject(rng)

    '        'End process

    '        msword.ActiveDocument.Save()

    '        MsgBox("Entity conversion is completed.", MsgBoxStyle.Information, "WebLink Automation")

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try

    'End Sub

    Private Sub CharEntityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CharEntityToolStripMenuItem.Click
        'Try

        '   Kill_word_Process()

        '    Dim msword As New Microsoft.Office.Interop.Word.Application
        '    msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")

        '    'To remove alphabet and number for increase speed
        '    Dim WordContent As String = msword.ActiveDocument.Content.Text.ToString
        '    'WordContent = Regex.Replace(WordContent, "[a-z|0-9|<|>| ]+", "", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        '    WordContent = Regex.Replace(WordContent, "[@|$|%|\*|!|a-z|0-9| |;|\.|\:|,|<|>|\/|\(|\-|\)]+", "", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        '    ProgressBar1.Visible = True
        '    ProgressBar1.Minimum = 0
        '    ProgressBar1.Maximum = WordContent.ToString.Count ''msword.ActiveDocument.Content.Words.Count
        '    '1596 1612  458
        '    'Process for entity conversion
        '    Dim rng As Word.Range
        '    'Iterate through each character
        '    For Each c As Char In WordContent.ToString

        '        ' If AscW(c) > 127 Then '255 Then '127 Then '

        '        'If Not Regex.IsMatch(c.ToString, "<|>", RegexOptions.Singleline) Then

        '        rng = msword.ActiveDocument.Content

        '        Dim kjhj As String = Convert.ToSByte(c).ToString("X4")

        '        With rng.Find

        '            .ClearFormatting()
        '            .MatchCase = True
        '            .Execute(FindText:=c.ToString, _
        '            ReplaceWith:=HttpUtility.HtmlEncode(c), _
        '            Replace:=Word.WdReplace.wdReplaceAll)

        '        End With

        '        ' End If

        '        'If ProgressBar1.Value = 80 Then
        '        '    MsgBox("ok")
        '        'End If


        '        If ProgressBar1.Value > WordContent.ToString.Count Then
        '            ProgressBar1.Value = ProgressBar1.Value - 10
        '        Else
        '            ProgressBar1.Value = ProgressBar1.Value + 1
        '        End If

        '    Next

        '    Marshal.ReleaseComObject(rng)

        '    'End process

        '    msword.ActiveDocument.Save()

        '    MsgBox("Entity conversion is completed.", MsgBoxStyle.Information, "Character Entity Conversion.")

        '    ProgressBar1.Value = 0
        '    ProgressBar1.Visible = False

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        '    ProgressBar1.Value = 0
        '    ProgressBar1.Visible = False
        'End Try



        Try

            'Code for database connection

            Dim g_strConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DbEntity.mdb" & ";Persist Security Info=False"

            Dim con As New OleDbConnection(g_strConnString)

            Dim cmd As New OleDbCommand

            con.Open()

            Dim conn As New OleDbConnection(g_strConnString)
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String = Nothing

            ' Try
            conn.Open()
            sql = "SELECT symbols, characterentity FROM EntityList"
            Dim ds As New DataSet
            da = New OleDb.OleDbDataAdapter(sql, conn)
            da.Fill(ds)


            'For doc file process
            Kill_word_Process()
            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")
            'End


          
            'Initial cleanup process
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "  "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = vbCr & vbCr
                .Replacement.Text = vbCr
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^t"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^b"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^m"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)





            'ProgressBar declaration
            ProgressBar1.Visible = True
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = ds.Tables(0).Rows.Count 

            For i = 0 To ds.Tables(0).Rows.Count - 1

                If ProgressBar1.Value > ds.Tables(0).Rows.Count Then
                    ProgressBar1.Value = ProgressBar1.Value - 10
                Else
                    ProgressBar1.Value = ProgressBar1.Value + 1
                End If

                Dim symbol_val As String = String.Empty
                Try
                    symbol_val = (ds.Tables(0).Rows(i).Item(0))
                Catch ex As Exception
                End Try
                Dim Character_ent As String = String.Empty
                Try
                    Character_ent = (ds.Tables(0).Rows(i).Item(1))
                Catch ex As Exception
                End Try

                'Dim Named_ent As String = (ds.Tables(0).Rows(i).Item(2))

                If Not symbol_val = Nothing AndAlso Not Character_ent = Nothing Then
                    FunEntityConvn(symbol_val, Character_ent)
                End If

                ' MS_CTRL.Run("Con_Entity", char_val, char_ents
            Next


            '****************************Test*******************
            'Code for replace bold tag
            msword.Selection.End = 0
            With msword.Selection.Find
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Bold = True
                .Replacement.Font.Bold = False
                .Execute(FindText:="", ReplaceWith:="<strong>^&</strong>", MatchWildcards:=True, Replace:=Word.WdReplace.wdReplaceAll)
            End With


            'Code for insert proper <Strong> closing and opening tag
            Dim rnge As Word.Range
            rnge = msword.ActiveDocument.Content

            msword.ActiveDocument.FormattingShowFont.ToString()

            rnge.Find.Font.Bold = True

            With rnge.Find 'Replace the unstructured </strong> tag
                .ClearFormatting()
                .Execute(FindText:=vbCr & "</strong>", _
                ReplaceWith:="</strong>" & vbCr, _
                Replace:=Word.WdReplace.wdReplaceAll)
            End With

            'rnge = msword.ActiveDocument.Content

            'msword.ActiveDocument.FormattingShowFont.ToString()

            'rnge.Find.Font.Bold = True

            'With rnge.Find 'Replace the unstructured </strong> tag
            '    .ClearFormatting()
            '    .Execute(FindText:="<strong>*</strong>", _
            '    ReplaceWith:="</strong>^p<strong>", _
            '    Replace:=Word.WdReplace.wdReplaceAll)
            'End With



            Dim StrongTag As String = msword.ActiveDocument.Content.Text.ToString
            Dim StrongTags As MatchCollection = Regex.Matches(StrongTag, "<strong>(((?!</strong>).)+)</strong>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            Dim TmpList As New ArrayList

            For Each STag As Match In StrongTags


                Try


                    Dim Entermark() As String = Regex.Split(STag.ToString, vbCr)
                    If Entermark.Count > 0 Then
                        For Each item As String In Entermark

                            If Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then

                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:=item.ToString & "</strong>", _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With

                                    TmpList.Add(item.ToString)

                                End If
                            ElseIf Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then
                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:="<strong>" & item.ToString, _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With
                                    TmpList.Add(item.ToString)
                                End If
                            ElseIf Not Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then
                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:="<strong>" & item.ToString & "</strong>", _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With
                                    TmpList.Add(item.ToString)
                                End If

                            End If

                        Next
                    End If
                Catch ex As Exception
                    'MsgBox(ex.Message)
                End Try
            Next


            Marshal.ReleaseComObject(rnge)


            'code for find italic
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Italic = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Italic = True Then
                    .Replacement.Text = "<em>^&</em>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find underline
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Underline = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Underline = WdUnderline.wdUnderlineSingle Then
                    .Replacement.Text = "<span class=" & ChrW(34) & "underline" & ChrW(34) & ">^&</span>"
                    '"<strong>^&</strong>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            'code for find strikethrough
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.StrikeThrough = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.StrikeThrough = True Then
                    .Replacement.Text = "<span class=" & ChrW(34) & "strike" & ChrW(34) & ">^&</span>"
                    '"<strong>^&</strong>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find smallcaps
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.SmallCaps = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.SmallCaps = True Then
                    .Replacement.Text = "<small>^&</small>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find image
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = "^g"
                .Replacement.Text = "<img></img>"
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            Dim ImgTag As MatchCollection = Regex.Matches(msword.ActiveDocument.Content.Text, "<img></img>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)


            Dim FileSeqcount As Integer = 0
            Dim FileSeqVar As String = ""

            For Each Img As Match In ImgTag

                'Code for append initial zero sequence
                FileSeqcount = FileSeqcount + 1
                If Regex.IsMatch(FileSeqcount, "^\d{1}$", RegexOptions.None) Then
                    FileSeqVar = "000" & FileSeqcount
                ElseIf Regex.IsMatch(FileSeqcount, "^\d{2}$", RegexOptions.None) Then
                    FileSeqVar = "00" & FileSeqcount
                ElseIf Regex.IsMatch(FileSeqcount, "^\d{3}$", RegexOptions.None) Then
                    FileSeqVar = "0" & FileSeqcount
                Else
                    FileSeqVar = FileSeqcount
                End If

                Dim rng_Img As Word.Range
                rng_Img = msword.ActiveDocument.Content

                With rng_Img.Find
                    .ClearFormatting()

                    .Execute(FindText:=Img.ToString, _
                  ReplaceWith:="<p class=" & ChrW(34) & "image" & ChrW(34) & "><img src=" & ChrW(34) & "images/f" & FileSeqVar & ".jpg" & ChrW(34) & " alt=" & ChrW(34) & "image" & ChrW(34) & "/></p>", _
                  Replace:=Word.WdReplace.wdReplaceOne)

                End With

            Next

            'Code for super & subscript

            'Selection.HomeKey(wdswdStory)
            'With Selection.Find
            '    .ClearFormatting()
            '    .Replacement.ClearFormatting()
            '    .Font.Superscript = True
            '    .Replacement.Text = "^^{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            '    .Font.Subscript = True
            '    .Replacement.Text = "_{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            'End With


            'Super & supscript
            Dim rng As Word.Range
            rng = msword.ActiveDocument.Content
            With rng.Find
                '.Font.Italic = True
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Superscript = True
                '.Replacement.Text = "<a id=""fn^&"" href=""#ft^&""><sup>^&</sup></a>" 
                .Replacement.Text = "<sup>^&</sup>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Subscript = True
                .Replacement.Text = "<sub>^&</sub>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

            End With

            Marshal.ReleaseComObject(rng)

            'Code for table processing
            If msword.ActiveDocument.Tables.Count > 0 Then

                Try

                    For TblCount As Integer = 1 To msword.ActiveDocument.Tables.Count
                        For RowCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Rows.Count

                            For ColumnCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Columns.Count
                                msword.ActiveDocument.Tables.Item(TblCount).Cell(RowCnt, ColumnCnt).Select()

                                msword.Selection.TypeText(Text:="<table>" & msword.Selection.Text.ToString() & "</table>")
                            Next

                        Next

                        'If msword.ActiveDocument.Tables.Item(i).Rows.Count > 0 Then
                        '    Dim hkjh As String = msword.ActiveDocument.Tables.Item(i).Rows.Item(1).ConvertToText.ToString
                        'End If

                    Next

                Catch ex As Exception

                End Try
            End If

            'To replace <strong> tag into normal formatting
            With msword.ActiveDocument.Content.Find
                .ClearFormatting()
                .Font.Bold = True
                With .Replacement
                    .ClearFormatting()
                    .Font.Bold = False
                End With
                .Execute(FindText:="<strong>", ReplaceWith:="<strong>", _
                    Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
            End With

            '****************************End********************









            'save word doc
            msword.ActiveDocument.Save()

            MsgBox("Symbol to character entity is converted successfully.", MsgBoxStyle.Information, "Entity conversion")

            ProgressBar1.Value = 0
            ProgressBar1.Visible = False

          
            conn.Close()
         
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
        End Try



    End Sub


    Private Sub PageIdSequenceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PageIdSequenceToolStripMenuItem.Click

        Try

            Kill_word_Process()

            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")

            'code for generate page id sequence
            Dim rngs As Word.Range
            rngs = msword.ActiveDocument.Content

            Dim pageNo As Word.Pages
            pageNo = msword.ActiveDocument.ActiveWindow.Panes(1).Pages

            For EachPage As Integer = 1 To pageNo.Count

                With msword.ActiveDocument
                    rngs = .GoTo(What:=Word.WdGoToItem.wdGoToPage, Count:=EachPage)
                    rngs.Select()
                    msword.Selection.InsertBefore("<a id=" & ChrW(34) & "page_" & EachPage & "" & ChrW(34) & "></a>")
                    'Rng = Rng.GoTo(What:=wdGoToBookmark, Name:="\page")
                    'Rng.Delete()

                End With

            Next

            msword.ActiveDocument.Save()

            MsgBox("PageId sequence is completed.", MsgBoxStyle.Information, "PageId sequence.")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "PageId sequence.")
        End Try

    End Sub

   
    Private Sub AddEntityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddEntityToolStripMenuItem.Click
        FrmEntityAdd.Show()
    End Sub

    
    Private Sub ToolStripMenuNamedEnt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuNamedEnt.Click
        Try

            'Code for database connection

            Dim g_strConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DbEntity.mdb" & ";Persist Security Info=False"

            Dim con As New OleDbConnection(g_strConnString)

            Dim cmd As New OleDbCommand

            con.Open()

            Dim conn As New OleDbConnection(g_strConnString)
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String = Nothing

            ' Try
            conn.Open()
            sql = "SELECT symbols, namedentity FROM EntityList"
            Dim ds As New DataSet
            da = New OleDb.OleDbDataAdapter(sql, conn)
            da.Fill(ds)


            'For doc file process
            Kill_word_Process()
            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")
            'End



            'Initial cleanup process
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "  "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = vbCr & vbCr
                .Replacement.Text = vbCr
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^t"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^b"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^m"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            


            'ProgressBar declaration
            ProgressBar1.Visible = True
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = ds.Tables(0).Rows.Count ''msword.ActiveDocument.Content.Words.Count


            For i = 0 To ds.Tables(0).Rows.Count - 1

                If ProgressBar1.Value > ds.Tables(0).Rows.Count Then
                    ProgressBar1.Value = ProgressBar1.Value - 10
                Else
                    ProgressBar1.Value = ProgressBar1.Value + 1
                End If

                Dim symbol_val As String = String.Empty
                Try
                    symbol_val = (ds.Tables(0).Rows(i).Item(0))
                Catch ex As Exception
                End Try
                Dim Named_ent As String = String.Empty
                Try
                    Named_ent = (ds.Tables(0).Rows(i).Item(1))
                Catch ex As Exception
                End Try

                'Dim Named_ent As String = (ds.Tables(0).Rows(i).Item(2))

                If Not symbol_val = Nothing AndAlso Not Named_ent = Nothing Then
                    FunEntityConvn(symbol_val, Named_ent)
                End If

                ' MS_CTRL.Run("Con_Entity", char_val, char_ents
            Next

          


            '****************************Test*******************
            'Code for replace bold tag
            msword.Selection.End = 0
            With msword.Selection.Find
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Bold = True
                .Replacement.Font.Bold = False
                .Execute(FindText:="", ReplaceWith:="<strong>^&</strong>", MatchWildcards:=True, Replace:=Word.WdReplace.wdReplaceAll)
            End With



            'Code for insert proper <Strong> closing and opening tag
            Dim rnge As Word.Range
            rnge = msword.ActiveDocument.Content

            msword.ActiveDocument.FormattingShowFont.ToString()

            rnge.Find.Font.Bold = True

            With rnge.Find 'Replace the unstructured </strong> tag
                .ClearFormatting()
                .Execute(FindText:=vbCr & "</strong>", _
                ReplaceWith:="</strong>" & vbCr, _
                Replace:=Word.WdReplace.wdReplaceAll)
            End With

            'rnge = msword.ActiveDocument.Content

            'msword.ActiveDocument.FormattingShowFont.ToString()

            'rnge.Find.Font.Bold = True

            'With rnge.Find 'Replace the unstructured </strong> tag
            '    .ClearFormatting()
            '    .Execute(FindText:="<strong>*</strong>", _
            '    ReplaceWith:="</strong>^p<strong>", _
            '    Replace:=Word.WdReplace.wdReplaceAll)
            'End With



            Dim StrongTag As String = msword.ActiveDocument.Content.Text.ToString
            Dim StrongTags As MatchCollection = Regex.Matches(StrongTag, "<strong>(((?!</strong>).)+)</strong>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            Dim TmpList As New ArrayList

            For Each STag As Match In StrongTags


                Try


                    Dim Entermark() As String = Regex.Split(STag.ToString, vbCr)
                    If Entermark.Count > 0 Then
                        For Each item As String In Entermark

                            If Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then

                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:=item.ToString & "</strong>", _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With

                                    TmpList.Add(item.ToString)

                                End If
                            ElseIf Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then
                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:="<strong>" & item.ToString, _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With
                                    TmpList.Add(item.ToString)
                                End If
                            ElseIf Not Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then
                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:="<strong>" & item.ToString & "</strong>", _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With
                                    TmpList.Add(item.ToString)
                                End If

                            End If

                        Next
                    End If
                Catch ex As Exception
                    'MsgBox(ex.Message)
                End Try
            Next


            Marshal.ReleaseComObject(rnge)


            'code for find italic
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Italic = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Italic = True Then
                    .Replacement.Text = "<em>^&</em>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find underline
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Underline = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Underline = WdUnderline.wdUnderlineSingle Then
                    .Replacement.Text = "<span class=" & ChrW(34) & "underline" & ChrW(34) & ">^&</span>"
                    '"<strong>^&</strong>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            'code for find strikethrough
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.StrikeThrough = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.StrikeThrough = True Then
                    .Replacement.Text = "<span class=" & ChrW(34) & "strike" & ChrW(34) & ">^&</span>"
                    '"<strong>^&</strong>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find smallcaps
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.SmallCaps = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.SmallCaps = True Then
                    .Replacement.Text = "<small>^&</small>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find image
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = "^g"
                .Replacement.Text = "<img></img>"
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            Dim ImgTag As MatchCollection = Regex.Matches(msword.ActiveDocument.Content.Text, "<img></img>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)


            Dim FileSeqcount As Integer = 0
            Dim FileSeqVar As String = ""

            For Each Img As Match In ImgTag

                'Code for append initial zero sequence
                FileSeqcount = FileSeqcount + 1
                If Regex.IsMatch(FileSeqcount, "^\d{1}$", RegexOptions.None) Then
                    FileSeqVar = "000" & FileSeqcount
                ElseIf Regex.IsMatch(FileSeqcount, "^\d{2}$", RegexOptions.None) Then
                    FileSeqVar = "00" & FileSeqcount
                ElseIf Regex.IsMatch(FileSeqcount, "^\d{3}$", RegexOptions.None) Then
                    FileSeqVar = "0" & FileSeqcount
                Else
                    FileSeqVar = FileSeqcount
                End If

                Dim rng_Img As Word.Range
                rng_Img = msword.ActiveDocument.Content

                With rng_Img.Find
                    .ClearFormatting()

                    .Execute(FindText:=Img.ToString, _
                  ReplaceWith:="<p class=" & ChrW(34) & "image" & ChrW(34) & "><img src=" & ChrW(34) & "images/f" & FileSeqVar & ".jpg" & ChrW(34) & " alt=" & ChrW(34) & "image" & ChrW(34) & "/></p>", _
                  Replace:=Word.WdReplace.wdReplaceOne)

                End With

            Next

            'Code for super & subscript

            'Selection.HomeKey(wdswdStory)
            'With Selection.Find
            '    .ClearFormatting()
            '    .Replacement.ClearFormatting()
            '    .Font.Superscript = True
            '    .Replacement.Text = "^^{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            '    .Font.Subscript = True
            '    .Replacement.Text = "_{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            'End With


            'Super & supscript
            Dim rng As Word.Range
            rng = msword.ActiveDocument.Content
            With rng.Find
                '.Font.Italic = True
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Superscript = True
                '.Replacement.Text = "<a id=""fn^&"" href=""#ft^&""><sup>^&</sup></a>" 
                .Replacement.Text = "<sup>^&</sup>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Subscript = True
                .Replacement.Text = "<sub>^&</sub>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

            End With

            Marshal.ReleaseComObject(rng)

            'Code for table processing
            If msword.ActiveDocument.Tables.Count > 0 Then

                Try

                    For TblCount As Integer = 1 To msword.ActiveDocument.Tables.Count
                        For RowCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Rows.Count

                            For ColumnCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Columns.Count
                                msword.ActiveDocument.Tables.Item(TblCount).Cell(RowCnt, ColumnCnt).Select()

                                msword.Selection.TypeText(Text:="<table>" & msword.Selection.Text.ToString() & "</table>")
                            Next

                        Next

                        'If msword.ActiveDocument.Tables.Item(i).Rows.Count > 0 Then
                        '    Dim hkjh As String = msword.ActiveDocument.Tables.Item(i).Rows.Item(1).ConvertToText.ToString
                        'End If

                    Next

                Catch ex As Exception

                End Try
            End If

            'To replace <strong> tag into normal formatting
            With msword.ActiveDocument.Content.Find
                .ClearFormatting()
                .Font.Bold = True
                With .Replacement
                    .ClearFormatting()
                    .Font.Bold = False
                End With
                .Execute(FindText:="<strong>", ReplaceWith:="<strong>", _
                    Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
            End With

            '****************************End********************








            'save word doc
            msword.ActiveDocument.Save()

            MsgBox("Symbol to Named entity is converted successfully.", MsgBoxStyle.Information, "Entity conversion")

            ProgressBar1.Value = 0
            ProgressBar1.Visible = False

            'Catch ex As OleDbException
            '    MsgBox("Error: " & ex.ToString())
            'Finally
            conn.Close()
            'End Try

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
        End Try

    End Sub

    'Private Sub ToolStripCharEntity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripCharEntity.Click

    '    Try
    '        Kill_word_Process()

    '        Dim msword As New Microsoft.Office.Interop.Word.Application
    '        msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")

    '        'To store entity value
    '        Dim ChaptPageColltn As New Dictionary(Of String, String)
    '        ChaptPageColltn.Add("34", "quot")
    '        ChaptPageColltn.Add("38", "amp")
    '        ChaptPageColltn.Add("39", "apos")
    '        ChaptPageColltn.Add("60", "lt")
    '        ChaptPageColltn.Add("62", "gt")
    '        ChaptPageColltn.Add("160", "nbsp")
    '        ChaptPageColltn.Add("161", "iexcl")
    '        ChaptPageColltn.Add("162", "cent")
    '        ChaptPageColltn.Add("163", "pound")
    '        ChaptPageColltn.Add("164", "curren")
    '        ChaptPageColltn.Add("165", "yen")
    '        ChaptPageColltn.Add("166", "brvbar")
    '        ChaptPageColltn.Add("167", "sect")
    '        ChaptPageColltn.Add("168", "uml")
    '        ChaptPageColltn.Add("169", "copy")
    '        ChaptPageColltn.Add("170", "ordf")
    '        ChaptPageColltn.Add("171", "laquo")
    '        ChaptPageColltn.Add("172", "not")
    '        ChaptPageColltn.Add("173", "shy")
    '        ChaptPageColltn.Add("174", "reg")
    '        ChaptPageColltn.Add("175", "macr")
    '        ChaptPageColltn.Add("176", "deg")
    '        ChaptPageColltn.Add("177", "plusmn")
    '        ChaptPageColltn.Add("178", "sup2")
    '        ChaptPageColltn.Add("179", "sup3")
    '        ChaptPageColltn.Add("180", "acute")
    '        ChaptPageColltn.Add("181", "micro")
    '        ChaptPageColltn.Add("182", "para")
    '        ChaptPageColltn.Add("183", "middot")
    '        ChaptPageColltn.Add("184", "cedil")
    '        ChaptPageColltn.Add("185", "sup1")
    '        ChaptPageColltn.Add("186", "ordm")
    '        ChaptPageColltn.Add("187", "raquo")
    '        ChaptPageColltn.Add("188", "frac14")
    '        ChaptPageColltn.Add("189", "frac12")
    '        ChaptPageColltn.Add("190", "frac34")
    '        ChaptPageColltn.Add("191", "iquest")
    '        ChaptPageColltn.Add("192", "Agrave")
    '        ChaptPageColltn.Add("193", "Aacute")
    '        ChaptPageColltn.Add("194", "Acirc")
    '        ChaptPageColltn.Add("195", "Atilde")
    '        ChaptPageColltn.Add("196", "Auml")
    '        ChaptPageColltn.Add("197", "Aring")
    '        ChaptPageColltn.Add("198", "AElig")
    '        ChaptPageColltn.Add("199", "Ccedil")
    '        ChaptPageColltn.Add("200", "Egrave")
    '        ChaptPageColltn.Add("201", "Eacute")
    '        ChaptPageColltn.Add("202", "Ecirc")
    '        ChaptPageColltn.Add("203", "Euml")
    '        ChaptPageColltn.Add("204", "Igrave")
    '        ChaptPageColltn.Add("205", "Iacute")
    '        ChaptPageColltn.Add("206", "Icirc")
    '        ChaptPageColltn.Add("207", "Iuml")
    '        ChaptPageColltn.Add("208", "ETH")
    '        ChaptPageColltn.Add("209", "Ntilde")
    '        ChaptPageColltn.Add("210", "Ograve")
    '        ChaptPageColltn.Add("211", "Oacute")
    '        ChaptPageColltn.Add("212", "Ocirc")
    '        ChaptPageColltn.Add("213", "Otilde")
    '        ChaptPageColltn.Add("214", "Ouml")
    '        ChaptPageColltn.Add("215", "times")
    '        ChaptPageColltn.Add("216", "Oslash")
    '        ChaptPageColltn.Add("217", "Ugrave")
    '        ChaptPageColltn.Add("218", "Uacute")
    '        ChaptPageColltn.Add("219", "Ucirc")
    '        ChaptPageColltn.Add("220", "Uuml")
    '        ChaptPageColltn.Add("221", "Yacute")
    '        ChaptPageColltn.Add("222", "THORN")
    '        ChaptPageColltn.Add("223", "szlig")
    '        ChaptPageColltn.Add("224", "agrave")
    '        ChaptPageColltn.Add("225", "aacute")
    '        ChaptPageColltn.Add("226", "acirc")
    '        ChaptPageColltn.Add("227", "atilde")
    '        ChaptPageColltn.Add("228", "auml")
    '        ChaptPageColltn.Add("229", "aring")
    '        ChaptPageColltn.Add("230", "aelig")
    '        ChaptPageColltn.Add("231", "ccedil")
    '        ChaptPageColltn.Add("232", "egrave")
    '        ChaptPageColltn.Add("233", "eacute")
    '        ChaptPageColltn.Add("234", "ecirc")
    '        ChaptPageColltn.Add("235", "euml")
    '        ChaptPageColltn.Add("236", "igrave")
    '        ChaptPageColltn.Add("237", "iacute")
    '        ChaptPageColltn.Add("238", "icirc")
    '        ChaptPageColltn.Add("239", "iuml")
    '        ChaptPageColltn.Add("240", "eth")
    '        ChaptPageColltn.Add("241", "ntilde")
    '        ChaptPageColltn.Add("242", "ograve")
    '        ChaptPageColltn.Add("243", "oacute")
    '        ChaptPageColltn.Add("244", "ocirc")
    '        ChaptPageColltn.Add("245", "otilde")
    '        ChaptPageColltn.Add("246", "ouml")
    '        ChaptPageColltn.Add("247", "divide")
    '        ChaptPageColltn.Add("248", "oslash")
    '        ChaptPageColltn.Add("249", "ugrave")
    '        ChaptPageColltn.Add("250", "uacute")
    '        ChaptPageColltn.Add("251", "ucirc")
    '        ChaptPageColltn.Add("252", "uuml")
    '        ChaptPageColltn.Add("253", "yacute")
    '        ChaptPageColltn.Add("254", "thorn")
    '        ChaptPageColltn.Add("255", "yuml")
    '        ChaptPageColltn.Add("402", "fnof")
    '        ChaptPageColltn.Add("913", "Alpha")
    '        ChaptPageColltn.Add("914", "Beta")
    '        ChaptPageColltn.Add("915", "Gamma")
    '        ChaptPageColltn.Add("916", "Delta")
    '        ChaptPageColltn.Add("917", "Epsilon")
    '        ChaptPageColltn.Add("918", "Zeta")
    '        ChaptPageColltn.Add("919", "Eta")
    '        ChaptPageColltn.Add("920", "Theta")
    '        ChaptPageColltn.Add("921", "Iota")
    '        ChaptPageColltn.Add("922", "Kappa")
    '        ChaptPageColltn.Add("923", "Lambda")
    '        ChaptPageColltn.Add("924", "Mu")
    '        ChaptPageColltn.Add("925", "Nu")
    '        ChaptPageColltn.Add("926", "Xi")
    '        ChaptPageColltn.Add("927", "Omicron")
    '        ChaptPageColltn.Add("928", "Pi")
    '        ChaptPageColltn.Add("929", "Rho")
    '        ChaptPageColltn.Add("931", "Sigma")
    '        ChaptPageColltn.Add("932", "Tau")
    '        ChaptPageColltn.Add("933", "Upsilon")
    '        ChaptPageColltn.Add("934", "Phi")
    '        ChaptPageColltn.Add("935", "Chi")
    '        ChaptPageColltn.Add("936", "Psi")
    '        ChaptPageColltn.Add("937", "Omega")
    '        ChaptPageColltn.Add("945", "alpha")
    '        ChaptPageColltn.Add("946", "beta")
    '        ChaptPageColltn.Add("947", "gamma")
    '        ChaptPageColltn.Add("948", "delta")
    '        ChaptPageColltn.Add("949", "epsilon")
    '        ChaptPageColltn.Add("950", "zeta")
    '        ChaptPageColltn.Add("951", "eta")
    '        ChaptPageColltn.Add("952", "theta")
    '        ChaptPageColltn.Add("953", "iota")
    '        ChaptPageColltn.Add("954", "kappa")
    '        ChaptPageColltn.Add("955", "lambda")
    '        ChaptPageColltn.Add("956", "mu")
    '        ChaptPageColltn.Add("957", "nu")
    '        ChaptPageColltn.Add("958", "xi")
    '        ChaptPageColltn.Add("959", "omicron")
    '        ChaptPageColltn.Add("960", "pi")
    '        ChaptPageColltn.Add("961", "rho")
    '        ChaptPageColltn.Add("962", "sigmaf")
    '        ChaptPageColltn.Add("963", "sigma")
    '        ChaptPageColltn.Add("964", "tau")
    '        ChaptPageColltn.Add("965", "upsilon")
    '        ChaptPageColltn.Add("966", "phi")
    '        ChaptPageColltn.Add("967", "chi")
    '        ChaptPageColltn.Add("968", "psi")
    '        ChaptPageColltn.Add("969", "omega")
    '        ChaptPageColltn.Add("977", "thetasym")
    '        ChaptPageColltn.Add("978", "upsih")
    '        ChaptPageColltn.Add("982", "piv")
    '        ChaptPageColltn.Add("8226", "bull")
    '        ChaptPageColltn.Add("8230", "hellip")
    '        ChaptPageColltn.Add("8242", "prime")
    '        ChaptPageColltn.Add("8243", "Prime")
    '        ChaptPageColltn.Add("8254", "oline")
    '        ChaptPageColltn.Add("8260", "frasl")
    '        ChaptPageColltn.Add("8472", "weierp")
    '        ChaptPageColltn.Add("8465", "image")
    '        ChaptPageColltn.Add("8476", "real")
    '        ChaptPageColltn.Add("8482", "trade")
    '        ChaptPageColltn.Add("8501", "alefsym")
    '        ChaptPageColltn.Add("8592", "larr")
    '        ChaptPageColltn.Add("8593", "uarr")
    '        ChaptPageColltn.Add("8594", "rarr")
    '        ChaptPageColltn.Add("8595", "darr")
    '        ChaptPageColltn.Add("8596", "harr")
    '        ChaptPageColltn.Add("8629", "crarr")
    '        ChaptPageColltn.Add("8656", "lArr")
    '        ChaptPageColltn.Add("8657", "uArr")
    '        ChaptPageColltn.Add("8658", "rArr")
    '        ChaptPageColltn.Add("8659", "dArr")
    '        ChaptPageColltn.Add("8660", "hArr")
    '        ChaptPageColltn.Add("8704", "forall")
    '        ChaptPageColltn.Add("8706", "part")
    '        ChaptPageColltn.Add("8707", "exist")
    '        ChaptPageColltn.Add("8709", "empty")
    '        ChaptPageColltn.Add("8711", "nabla")
    '        ChaptPageColltn.Add("8712", "isin")
    '        ChaptPageColltn.Add("8713", "notin")
    '        ChaptPageColltn.Add("8715", "ni")
    '        ChaptPageColltn.Add("8719", "prod")
    '        ChaptPageColltn.Add("8721", "sum")
    '        ChaptPageColltn.Add("8722", "minus")
    '        ChaptPageColltn.Add("8727", "lowast")
    '        ChaptPageColltn.Add("8730", "radic")
    '        ChaptPageColltn.Add("8733", "prop")
    '        ChaptPageColltn.Add("8734", "infin")
    '        ChaptPageColltn.Add("8736", "ang")
    '        ChaptPageColltn.Add("8743", "and")
    '        ChaptPageColltn.Add("8744", "or")
    '        ChaptPageColltn.Add("8745", "cap")
    '        ChaptPageColltn.Add("8746", "cup")
    '        ChaptPageColltn.Add("8747", "int")
    '        ChaptPageColltn.Add("8756", "there4")
    '        ChaptPageColltn.Add("8764", "sim")
    '        ChaptPageColltn.Add("8773", "cong")
    '        ChaptPageColltn.Add("8776", "asymp")
    '        ChaptPageColltn.Add("8800", "ne")
    '        ChaptPageColltn.Add("8801", "equiv")
    '        ChaptPageColltn.Add("8804", "le")
    '        ChaptPageColltn.Add("8805", "ge")
    '        ChaptPageColltn.Add("8834", "sub")
    '        ChaptPageColltn.Add("8835", "sup")
    '        ChaptPageColltn.Add("8836", "nsub")
    '        ChaptPageColltn.Add("8838", "sube")
    '        ChaptPageColltn.Add("8839", "supe")
    '        ChaptPageColltn.Add("8853", "oplus")
    '        ChaptPageColltn.Add("8855", "otimes")
    '        ChaptPageColltn.Add("8869", "perp")
    '        ChaptPageColltn.Add("8901", "sdot")
    '        ChaptPageColltn.Add("8968", "lceil")
    '        ChaptPageColltn.Add("8969", "rceil")
    '        ChaptPageColltn.Add("8970", "lfloor")
    '        ChaptPageColltn.Add("8971", "rfloor")
    '        ChaptPageColltn.Add("9001", "lang")
    '        ChaptPageColltn.Add("9002", "rang")
    '        ChaptPageColltn.Add("9674", "loz")
    '        ChaptPageColltn.Add("9824", "spades")
    '        ChaptPageColltn.Add("9827", "clubs")
    '        ChaptPageColltn.Add("9829", "hearts")
    '        ChaptPageColltn.Add("9830", "diams")
    '        ChaptPageColltn.Add("338", "OElig")
    '        ChaptPageColltn.Add("339", "oelig")
    '        ChaptPageColltn.Add("352", "Scaron")
    '        ChaptPageColltn.Add("353", "scaron")
    '        ChaptPageColltn.Add("376", "Yuml")
    '        ChaptPageColltn.Add("710", "circ")
    '        ChaptPageColltn.Add("732", "tilde")
    '        ChaptPageColltn.Add("8194", "ensp")
    '        ChaptPageColltn.Add("8195", "emsp")
    '        ChaptPageColltn.Add("8201", "thinsp")
    '        ChaptPageColltn.Add("8204", "zwnj")
    '        ChaptPageColltn.Add("8205", "zwj")
    '        ChaptPageColltn.Add("8206", "lrm")
    '        ChaptPageColltn.Add("8207", "rlm")
    '        ChaptPageColltn.Add("8211", "ndash")
    '        ChaptPageColltn.Add("8212", "mdash")
    '        ChaptPageColltn.Add("8216", "lsquo")
    '        ChaptPageColltn.Add("8217", "rsquo")
    '        ChaptPageColltn.Add("8218", "sbquo")
    '        ChaptPageColltn.Add("8220", "ldquo")
    '        ChaptPageColltn.Add("8221", "rdquo")
    '        ChaptPageColltn.Add("8222", "bdquo")
    '        ChaptPageColltn.Add("8224", "dagger")
    '        ChaptPageColltn.Add("8225", "Dagger")
    '        ChaptPageColltn.Add("8240", "permil")
    '        ChaptPageColltn.Add("8249", "lsaquo")
    '        ChaptPageColltn.Add("8250", "rsaquo")
    '        ChaptPageColltn.Add("8364", "euro")

    '        ProgressBar1.Visible = True
    '        ProgressBar1.Minimum = 0
    '        ProgressBar1.Maximum = ChaptPageColltn.Count


    '        For Each Id As String In ChaptPageColltn.Keys

    '            If ProgressBar1.Value > ChaptPageColltn.Count Then
    '                ProgressBar1.Value = ProgressBar1.Value - 10
    '            Else
    '                ProgressBar1.Value = ProgressBar1.Value + 1
    '            End If

    '            msword.Selection.Find.ClearFormatting()
    '            msword.Selection.Find.Replacement.ClearFormatting()
    '            msword.Selection.End = 0
    '            With msword.Selection.Find
    '                .Text = "&#" & Id.ToString & ";"
    '                .Replacement.Text = "&" & ChaptPageColltn.Item(Id) & ";"
    '                .Forward = True
    '                .Wrap = WdFindWrap.wdFindContinue
    '                .Format = False
    '                .MatchCase = False
    '                .MatchWholeWord = False
    '                .MatchWildcards = False
    '                .MatchSoundsLike = False
    '                .MatchAllWordForms = False
    '            End With
    '            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

    '        Next
    '        'End

    '        'ProgressBar1.Minimum = 0
    '        'ProgressBar1.Maximum = ""

    '      msword.ActiveDocument.Save()

    '        MsgBox("Charater Entity conversion is completed.", MsgBoxStyle.Information, "Entity Conversion")

    '        ProgressBar1.Value = 0
    '        ProgressBar1.Visible = False

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Character Entity Conversion")
    '        ProgressBar1.Value = 0
    '        ProgressBar1.Visible = False
    '    End Try

    'End Sub


    Private Sub GeneralCitationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GeneralCitationToolStripMenuItem.Click
        FrmGenrlCitationLink.Show()
    End Sub

  

    Private Sub ChapterCitationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChapterCitationToolStripMenuItem.Click



        FrmFolderBrowse.Show()
    End Sub

   
    Private Sub WordFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WordFileToolStripMenuItem.Click

        Try
            Kill_word_Process()

            Dim msword As New Microsoft.Office.Interop.Word.Application
            msword = System.Runtime.InteropServices.Marshal.GetActiveObject("word.Application")



            'Initial cleanup process
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "  "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = vbCr & vbCr
                .Replacement.Text = vbCr
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^t"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^b"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            msword.Selection.End = 0
            With msword.Selection.Find
                .Text = "^m"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)







            'To store entity value
            Dim ChaptPageColltn As New Dictionary(Of String, String)

            ChaptPageColltn.Add("34", "quot")
            ChaptPageColltn.Add("38", "amp")
            ChaptPageColltn.Add("39", "apos")
            ChaptPageColltn.Add("60", "lt")
            ChaptPageColltn.Add("62", "gt")
            ChaptPageColltn.Add("160", "nbsp")
            ChaptPageColltn.Add("161", "iexcl")
            ChaptPageColltn.Add("162", "cent")
            ChaptPageColltn.Add("163", "pound")
            ChaptPageColltn.Add("164", "curren")
            ChaptPageColltn.Add("165", "yen")
            ChaptPageColltn.Add("166", "brvbar")
            ChaptPageColltn.Add("167", "sect")
            ChaptPageColltn.Add("168", "uml")
            ChaptPageColltn.Add("169", "copy")
            ChaptPageColltn.Add("170", "ordf")
            ChaptPageColltn.Add("171", "laquo")
            ChaptPageColltn.Add("172", "not")
            ChaptPageColltn.Add("173", "shy")
            ChaptPageColltn.Add("174", "reg")
            ChaptPageColltn.Add("175", "macr")
            ChaptPageColltn.Add("176", "deg")
            ChaptPageColltn.Add("177", "plusmn")
            ChaptPageColltn.Add("178", "sup2")
            ChaptPageColltn.Add("179", "sup3")
            ChaptPageColltn.Add("180", "acute")
            ChaptPageColltn.Add("181", "micro")
            ChaptPageColltn.Add("182", "para")
            ChaptPageColltn.Add("183", "middot")
            ChaptPageColltn.Add("184", "cedil")
            ChaptPageColltn.Add("185", "sup1")
            ChaptPageColltn.Add("186", "ordm")
            ChaptPageColltn.Add("187", "raquo")
            ChaptPageColltn.Add("188", "frac14")
            ChaptPageColltn.Add("189", "frac12")
            ChaptPageColltn.Add("190", "frac34")
            ChaptPageColltn.Add("191", "iquest")
            ChaptPageColltn.Add("192", "Agrave")
            ChaptPageColltn.Add("193", "Aacute")
            ChaptPageColltn.Add("194", "Acirc")
            ChaptPageColltn.Add("195", "Atilde")
            ChaptPageColltn.Add("196", "Auml")
            ChaptPageColltn.Add("197", "Aring")
            ChaptPageColltn.Add("198", "AElig")
            ChaptPageColltn.Add("199", "Ccedil")
            ChaptPageColltn.Add("200", "Egrave")
            ChaptPageColltn.Add("201", "Eacute")
            ChaptPageColltn.Add("202", "Ecirc")
            ChaptPageColltn.Add("203", "Euml")
            ChaptPageColltn.Add("204", "Igrave")
            ChaptPageColltn.Add("205", "Iacute")
            ChaptPageColltn.Add("206", "Icirc")
            ChaptPageColltn.Add("207", "Iuml")
            ChaptPageColltn.Add("208", "ETH")
            ChaptPageColltn.Add("209", "Ntilde")
            ChaptPageColltn.Add("210", "Ograve")
            ChaptPageColltn.Add("211", "Oacute")
            ChaptPageColltn.Add("212", "Ocirc")
            ChaptPageColltn.Add("213", "Otilde")
            ChaptPageColltn.Add("214", "Ouml")
            ChaptPageColltn.Add("215", "times")
            ChaptPageColltn.Add("216", "Oslash")
            ChaptPageColltn.Add("217", "Ugrave")
            ChaptPageColltn.Add("218", "Uacute")
            ChaptPageColltn.Add("219", "Ucirc")
            ChaptPageColltn.Add("220", "Uuml")
            ChaptPageColltn.Add("221", "Yacute")
            ChaptPageColltn.Add("222", "THORN")
            ChaptPageColltn.Add("223", "szlig")
            ChaptPageColltn.Add("224", "agrave")
            ChaptPageColltn.Add("225", "aacute")
            ChaptPageColltn.Add("226", "acirc")
            ChaptPageColltn.Add("227", "atilde")
            ChaptPageColltn.Add("228", "auml")
            ChaptPageColltn.Add("229", "aring")
            ChaptPageColltn.Add("230", "aelig")
            ChaptPageColltn.Add("231", "ccedil")
            ChaptPageColltn.Add("232", "egrave")
            ChaptPageColltn.Add("233", "eacute")
            ChaptPageColltn.Add("234", "ecirc")
            ChaptPageColltn.Add("235", "euml")
            ChaptPageColltn.Add("236", "igrave")
            ChaptPageColltn.Add("237", "iacute")
            ChaptPageColltn.Add("238", "icirc")
            ChaptPageColltn.Add("239", "iuml")
            ChaptPageColltn.Add("240", "eth")
            ChaptPageColltn.Add("241", "ntilde")
            ChaptPageColltn.Add("242", "ograve")
            ChaptPageColltn.Add("243", "oacute")
            ChaptPageColltn.Add("244", "ocirc")
            ChaptPageColltn.Add("245", "otilde")
            ChaptPageColltn.Add("246", "ouml")
            ChaptPageColltn.Add("247", "divide")
            ChaptPageColltn.Add("248", "oslash")
            ChaptPageColltn.Add("249", "ugrave")
            ChaptPageColltn.Add("250", "uacute")
            ChaptPageColltn.Add("251", "ucirc")
            ChaptPageColltn.Add("252", "uuml")
            ChaptPageColltn.Add("253", "yacute")
            ChaptPageColltn.Add("254", "thorn")
            ChaptPageColltn.Add("255", "yuml")
            ChaptPageColltn.Add("402", "fnof")
            ChaptPageColltn.Add("913", "Alpha")
            ChaptPageColltn.Add("914", "Beta")
            ChaptPageColltn.Add("915", "Gamma")
            ChaptPageColltn.Add("916", "Delta")
            ChaptPageColltn.Add("917", "Epsilon")
            ChaptPageColltn.Add("918", "Zeta")
            ChaptPageColltn.Add("919", "Eta")
            ChaptPageColltn.Add("920", "Theta")
            ChaptPageColltn.Add("921", "Iota")
            ChaptPageColltn.Add("922", "Kappa")
            ChaptPageColltn.Add("923", "Lambda")
            ChaptPageColltn.Add("924", "Mu")
            ChaptPageColltn.Add("925", "Nu")
            ChaptPageColltn.Add("926", "Xi")
            ChaptPageColltn.Add("927", "Omicron")
            ChaptPageColltn.Add("928", "Pi")
            ChaptPageColltn.Add("929", "Rho")
            ChaptPageColltn.Add("931", "Sigma")
            ChaptPageColltn.Add("932", "Tau")
            ChaptPageColltn.Add("933", "Upsilon")
            ChaptPageColltn.Add("934", "Phi")
            ChaptPageColltn.Add("935", "Chi")
            ChaptPageColltn.Add("936", "Psi")
            ChaptPageColltn.Add("937", "Omega")
            ChaptPageColltn.Add("945", "alpha")
            ChaptPageColltn.Add("946", "beta")
            ChaptPageColltn.Add("947", "gamma")
            ChaptPageColltn.Add("948", "delta")
            ChaptPageColltn.Add("949", "epsilon")
            ChaptPageColltn.Add("950", "zeta")
            ChaptPageColltn.Add("951", "eta")
            ChaptPageColltn.Add("952", "theta")
            ChaptPageColltn.Add("953", "iota")
            ChaptPageColltn.Add("954", "kappa")
            ChaptPageColltn.Add("955", "lambda")
            ChaptPageColltn.Add("956", "mu")
            ChaptPageColltn.Add("957", "nu")
            ChaptPageColltn.Add("958", "xi")
            ChaptPageColltn.Add("959", "omicron")
            ChaptPageColltn.Add("960", "pi")
            ChaptPageColltn.Add("961", "rho")
            ChaptPageColltn.Add("962", "sigmaf")
            ChaptPageColltn.Add("963", "sigma")
            ChaptPageColltn.Add("964", "tau")
            ChaptPageColltn.Add("965", "upsilon")
            ChaptPageColltn.Add("966", "phi")
            ChaptPageColltn.Add("967", "chi")
            ChaptPageColltn.Add("968", "psi")
            ChaptPageColltn.Add("969", "omega")
            ChaptPageColltn.Add("977", "thetasym")
            ChaptPageColltn.Add("978", "upsih")
            ChaptPageColltn.Add("982", "piv")
            ChaptPageColltn.Add("8226", "bull")
            ChaptPageColltn.Add("8230", "hellip")
            ChaptPageColltn.Add("8242", "prime")
            ChaptPageColltn.Add("8243", "Prime")
            ChaptPageColltn.Add("8254", "oline")
            ChaptPageColltn.Add("8260", "frasl")
            ChaptPageColltn.Add("8472", "weierp")
            ChaptPageColltn.Add("8465", "image")
            ChaptPageColltn.Add("8476", "real")
            ChaptPageColltn.Add("8482", "trade")
            ChaptPageColltn.Add("8501", "alefsym")
            ChaptPageColltn.Add("8592", "larr")
            ChaptPageColltn.Add("8593", "uarr")
            ChaptPageColltn.Add("8594", "rarr")
            ChaptPageColltn.Add("8595", "darr")
            ChaptPageColltn.Add("8596", "harr")
            ChaptPageColltn.Add("8629", "crarr")
            ChaptPageColltn.Add("8656", "lArr")
            ChaptPageColltn.Add("8657", "uArr")
            ChaptPageColltn.Add("8658", "rArr")
            ChaptPageColltn.Add("8659", "dArr")
            ChaptPageColltn.Add("8660", "hArr")
            ChaptPageColltn.Add("8704", "forall")
            ChaptPageColltn.Add("8706", "part")
            ChaptPageColltn.Add("8707", "exist")
            ChaptPageColltn.Add("8709", "empty")
            ChaptPageColltn.Add("8711", "nabla")
            ChaptPageColltn.Add("8712", "isin")
            ChaptPageColltn.Add("8713", "notin")
            ChaptPageColltn.Add("8715", "ni")
            ChaptPageColltn.Add("8719", "prod")
            ChaptPageColltn.Add("8721", "sum")
            ChaptPageColltn.Add("8722", "minus")
            ChaptPageColltn.Add("8727", "lowast")
            ChaptPageColltn.Add("8730", "radic")
            ChaptPageColltn.Add("8733", "prop")
            ChaptPageColltn.Add("8734", "infin")
            ChaptPageColltn.Add("8736", "ang")
            ChaptPageColltn.Add("8743", "and")
            ChaptPageColltn.Add("8744", "or")
            ChaptPageColltn.Add("8745", "cap")
            ChaptPageColltn.Add("8746", "cup")
            ChaptPageColltn.Add("8747", "int")
            ChaptPageColltn.Add("8756", "there4")
            ChaptPageColltn.Add("8764", "sim")
            ChaptPageColltn.Add("8773", "cong")
            ChaptPageColltn.Add("8776", "asymp")
            ChaptPageColltn.Add("8800", "ne")
            ChaptPageColltn.Add("8801", "equiv")
            ChaptPageColltn.Add("8804", "le")
            ChaptPageColltn.Add("8805", "ge")
            ChaptPageColltn.Add("8834", "sub")
            ChaptPageColltn.Add("8835", "sup")
            ChaptPageColltn.Add("8836", "nsub")
            ChaptPageColltn.Add("8838", "sube")
            ChaptPageColltn.Add("8839", "supe")
            ChaptPageColltn.Add("8853", "oplus")
            ChaptPageColltn.Add("8855", "otimes")
            ChaptPageColltn.Add("8869", "perp")
            ChaptPageColltn.Add("8901", "sdot")
            ChaptPageColltn.Add("8968", "lceil")
            ChaptPageColltn.Add("8969", "rceil")
            ChaptPageColltn.Add("8970", "lfloor")
            ChaptPageColltn.Add("8971", "rfloor")
            ChaptPageColltn.Add("9001", "lang")
            ChaptPageColltn.Add("9002", "rang")
            ChaptPageColltn.Add("9674", "loz")
            ChaptPageColltn.Add("9824", "spades")
            ChaptPageColltn.Add("9827", "clubs")
            ChaptPageColltn.Add("9829", "hearts")
            ChaptPageColltn.Add("9830", "diams")
            ChaptPageColltn.Add("338", "OElig")
            ChaptPageColltn.Add("339", "oelig")
            ChaptPageColltn.Add("352", "Scaron")
            ChaptPageColltn.Add("353", "scaron")
            ChaptPageColltn.Add("376", "Yuml")
            ChaptPageColltn.Add("710", "circ")
            ChaptPageColltn.Add("732", "tilde")
            ChaptPageColltn.Add("8194", "ensp")
            ChaptPageColltn.Add("8195", "emsp")
            ChaptPageColltn.Add("8201", "thinsp")
            ChaptPageColltn.Add("8204", "zwnj")
            ChaptPageColltn.Add("8205", "zwj")
            ChaptPageColltn.Add("8206", "lrm")
            ChaptPageColltn.Add("8207", "rlm")
            ChaptPageColltn.Add("8211", "ndash")
            ChaptPageColltn.Add("8212", "mdash")
            ChaptPageColltn.Add("8216", "lsquo")
            ChaptPageColltn.Add("8217", "rsquo")
            ChaptPageColltn.Add("8218", "sbquo")
            ChaptPageColltn.Add("8220", "ldquo")
            ChaptPageColltn.Add("8221", "rdquo")
            ChaptPageColltn.Add("8222", "bdquo")
            ChaptPageColltn.Add("8224", "dagger")
            ChaptPageColltn.Add("8225", "Dagger")
            ChaptPageColltn.Add("8240", "permil")
            ChaptPageColltn.Add("8249", "lsaquo")
            ChaptPageColltn.Add("8250", "rsaquo")
            ChaptPageColltn.Add("8364", "euro")

            ProgressBar1.Visible = True
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = ChaptPageColltn.Count


            For Each Id As String In ChaptPageColltn.Keys

                If ProgressBar1.Value > ChaptPageColltn.Count Then
                    ProgressBar1.Value = ProgressBar1.Value - 10
                Else
                    ProgressBar1.Value = ProgressBar1.Value + 1
                End If

                msword.Selection.Find.ClearFormatting()
                msword.Selection.Find.Replacement.ClearFormatting()
                msword.Selection.End = 0

                With msword.Selection.Find
                    .Text = "&#" & Id.ToString & ";"
                    .Replacement.Text = "&" & ChaptPageColltn.Item(Id) & ";"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            Next
            'End


            '****************************Test*******************
            'Code for replace bold tag
            msword.Selection.End = 0
            With msword.Selection.Find
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Bold = True
                .Replacement.Font.Bold = False
                .Execute(FindText:="", ReplaceWith:="<strong>^&</strong>", MatchWildcards:=True, Replace:=Word.WdReplace.wdReplaceAll)
            End With

            'Code for insert proper <Strong> closing and opening tag
            Dim rnge As Word.Range
            rnge = msword.ActiveDocument.Content

            msword.ActiveDocument.FormattingShowFont.ToString()

            rnge.Find.Font.Bold = True

            With rnge.Find 'Replace the unstructured </strong> tag
                .ClearFormatting()
                .Execute(FindText:=vbCr & "</strong>", _
                ReplaceWith:="</strong>" & vbCr, _
                Replace:=Word.WdReplace.wdReplaceAll)
            End With

            'rnge = msword.ActiveDocument.Content

            'msword.ActiveDocument.FormattingShowFont.ToString()

            'rnge.Find.Font.Bold = True

            'With rnge.Find 'Replace the unstructured </strong> tag
            '    .ClearFormatting()
            '    .Execute(FindText:="<strong>*</strong>", _
            '    ReplaceWith:="</strong>^p<strong>", _
            '    Replace:=Word.WdReplace.wdReplaceAll)
            'End With



            Dim StrongTag As String = msword.ActiveDocument.Content.Text.ToString
            Dim StrongTags As MatchCollection = Regex.Matches(StrongTag, "<strong>(((?!</strong>).)+)</strong>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            Dim TmpList As New ArrayList

            For Each STag As Match In StrongTags


                Try


                    Dim Entermark() As String = Regex.Split(STag.ToString, vbCr)
                    If Entermark.Count > 0 Then
                        For Each item As String In Entermark

                            If Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then

                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:=item.ToString & "</strong>", _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With

                                    TmpList.Add(item.ToString)

                                End If
                            ElseIf Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then
                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:="<strong>" & item.ToString, _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With
                                    TmpList.Add(item.ToString)
                                End If
                            ElseIf Not Regex.IsMatch(item.ToString, "</strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) AndAlso Not Regex.IsMatch(item.ToString, "<strong>", RegexOptions.Singleline Or RegexOptions.IgnoreCase) Then
                                If Not TmpList.Contains(item.ToString) Then
                                    rnge = msword.ActiveDocument.Content
                                    With rnge.Find
                                        .Format = True
                                        .ClearFormatting()
                                        .Execute(FindText:=item.ToString, _
                                        ReplaceWith:="<strong>" & item.ToString & "</strong>", _
                                        Replace:=Word.WdReplace.wdReplaceOne)
                                    End With
                                    TmpList.Add(item.ToString)
                                End If

                            End If

                        Next
                    End If
                Catch ex As Exception
                    'MsgBox(ex.Message)
                End Try
            Next


            Marshal.ReleaseComObject(rnge)


            'code for find italic
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Italic = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Italic = True Then
                    .Replacement.Text = "<em>^&</em>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find underline
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.Underline = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.Underline = WdUnderline.wdUnderlineSingle Then
                    .Replacement.Text = "<span class=" & ChrW(34) & "underline" & ChrW(34) & ">^&</span>"
                    '"<strong>^&</strong>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)


            'code for find strikethrough
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.StrikeThrough = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.StrikeThrough = True Then
                    .Replacement.Text = "<span class=" & ChrW(34) & "strike" & ChrW(34) & ">^&</span>"
                    '"<strong>^&</strong>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find smallcaps
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Font.SmallCaps = True
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = ""
                If msword.Selection.Find.Font.SmallCaps = True Then
                    .Replacement.Text = "<small>^&</small>"
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End If
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            'code for find image
            msword.Selection.Find.ClearFormatting()
            msword.Selection.Find.Replacement.ClearFormatting()
            With msword.Selection.Find
                .Text = "^g"
                .Replacement.Text = "<img></img>"
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            msword.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

            Dim ImgTag As MatchCollection = Regex.Matches(msword.ActiveDocument.Content.Text, "<img></img>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)


            Dim FileSeqcount As Integer = 0
            Dim FileSeqVar As String = ""

            For Each Img As Match In ImgTag

                'Code for append initial zero sequence
                FileSeqcount = FileSeqcount + 1
                If Regex.IsMatch(FileSeqcount, "^\d{1}$", RegexOptions.None) Then
                    FileSeqVar = "000" & FileSeqcount
                ElseIf Regex.IsMatch(FileSeqcount, "^\d{2}$", RegexOptions.None) Then
                    FileSeqVar = "00" & FileSeqcount
                ElseIf Regex.IsMatch(FileSeqcount, "^\d{3}$", RegexOptions.None) Then
                    FileSeqVar = "0" & FileSeqcount
                Else
                    FileSeqVar = FileSeqcount
                End If

                Dim rng_Img As Word.Range
                rng_Img = msword.ActiveDocument.Content

                With rng_Img.Find
                    .ClearFormatting()

                    .Execute(FindText:=Img.ToString, _
                  ReplaceWith:="<p class=" & ChrW(34) & "image" & ChrW(34) & "><img src=" & ChrW(34) & "images/f" & FileSeqVar & ".jpg" & ChrW(34) & " alt=" & ChrW(34) & "image" & ChrW(34) & "/></p>", _
                  Replace:=Word.WdReplace.wdReplaceOne)

                End With

            Next

            'Code for super & subscript

            'Selection.HomeKey(wdswdStory)
            'With Selection.Find
            '    .ClearFormatting()
            '    .Replacement.ClearFormatting()
            '    .Font.Superscript = True
            '    .Replacement.Text = "^^{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            '    .Font.Subscript = True
            '    .Replacement.Text = "_{^&}"
            '    .Execute(Replace:=wdReplaceAll)
            'End With


            'Super & supscript
            Dim rng As Word.Range
            rng = msword.ActiveDocument.Content
            With rng.Find
                '.Font.Italic = True
                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Superscript = True
                '.Replacement.Text = "<a id=""fn^&"" href=""#ft^&""><sup>^&</sup></a>" 
                .Replacement.Text = "<sup>^&</sup>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

                .ClearFormatting()
                .Replacement.ClearFormatting()
                .Font.Subscript = True
                .Replacement.Text = "<sub>^&</sub>"
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)

            End With

            Marshal.ReleaseComObject(rng)

            'Code for table processing
            If msword.ActiveDocument.Tables.Count > 0 Then

                Try

                    For TblCount As Integer = 1 To msword.ActiveDocument.Tables.Count
                        For RowCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Rows.Count

                            For ColumnCnt As Integer = 1 To msword.ActiveDocument.Tables.Item(TblCount).Columns.Count
                                msword.ActiveDocument.Tables.Item(TblCount).Cell(RowCnt, ColumnCnt).Select()

                                msword.Selection.TypeText(Text:="<table>" & msword.Selection.Text.ToString() & "</table>")
                            Next

                        Next

                        'If msword.ActiveDocument.Tables.Item(i).Rows.Count > 0 Then
                        '    Dim hkjh As String = msword.ActiveDocument.Tables.Item(i).Rows.Item(1).ConvertToText.ToString
                        'End If

                    Next

                Catch ex As Exception

                End Try
            End If

            'To replace <strong> tag into normal formatting
            With msword.ActiveDocument.Content.Find
                .ClearFormatting()
                .Font.Bold = True
                With .Replacement
                    .ClearFormatting()
                    .Font.Bold = False
                End With
                .Execute(FindText:="<strong>", ReplaceWith:="<strong>", _
                    Format:=True, Replace:=Word.WdReplace.wdReplaceAll)
            End With

            '****************************End********************










            'ProgressBar1.Minimum = 0
            'ProgressBar1.Maximum = ""

            msword.ActiveDocument.Save()

            MsgBox("Charater Entity conversion is completed.", MsgBoxStyle.Information, "Entity Conversion")

            ProgressBar1.Value = 0
            ProgressBar1.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Character Entity Conversion")
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
        End Try

    End Sub

    Private Sub HtmlFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HtmlFileToolStripMenuItem.Click

        Try
           
            'To store entity value
            Dim ChaptPageColltn As New Dictionary(Of String, String)
            ChaptPageColltn.Add("34", "quot")
            ChaptPageColltn.Add("38", "amp")
            ChaptPageColltn.Add("39", "apos")
            ChaptPageColltn.Add("60", "lt")
            ChaptPageColltn.Add("62", "gt")
            ChaptPageColltn.Add("160", "nbsp")
            ChaptPageColltn.Add("161", "iexcl")
            ChaptPageColltn.Add("162", "cent")
            ChaptPageColltn.Add("163", "pound")
            ChaptPageColltn.Add("164", "curren")
            ChaptPageColltn.Add("165", "yen")
            ChaptPageColltn.Add("166", "brvbar")
            ChaptPageColltn.Add("167", "sect")
            ChaptPageColltn.Add("168", "uml")
            ChaptPageColltn.Add("169", "copy")
            ChaptPageColltn.Add("170", "ordf")
            ChaptPageColltn.Add("171", "laquo")
            ChaptPageColltn.Add("172", "not")
            ChaptPageColltn.Add("173", "shy")
            ChaptPageColltn.Add("174", "reg")
            ChaptPageColltn.Add("175", "macr")
            ChaptPageColltn.Add("176", "deg")
            ChaptPageColltn.Add("177", "plusmn")
            ChaptPageColltn.Add("178", "sup2")
            ChaptPageColltn.Add("179", "sup3")
            ChaptPageColltn.Add("180", "acute")
            ChaptPageColltn.Add("181", "micro")
            ChaptPageColltn.Add("182", "para")
            ChaptPageColltn.Add("183", "middot")
            ChaptPageColltn.Add("184", "cedil")
            ChaptPageColltn.Add("185", "sup1")
            ChaptPageColltn.Add("186", "ordm")
            ChaptPageColltn.Add("187", "raquo")
            ChaptPageColltn.Add("188", "frac14")
            ChaptPageColltn.Add("189", "frac12")
            ChaptPageColltn.Add("190", "frac34")
            ChaptPageColltn.Add("191", "iquest")
            ChaptPageColltn.Add("192", "Agrave")
            ChaptPageColltn.Add("193", "Aacute")
            ChaptPageColltn.Add("194", "Acirc")
            ChaptPageColltn.Add("195", "Atilde")
            ChaptPageColltn.Add("196", "Auml")
            ChaptPageColltn.Add("197", "Aring")
            ChaptPageColltn.Add("198", "AElig")
            ChaptPageColltn.Add("199", "Ccedil")
            ChaptPageColltn.Add("200", "Egrave")
            ChaptPageColltn.Add("201", "Eacute")
            ChaptPageColltn.Add("202", "Ecirc")
            ChaptPageColltn.Add("203", "Euml")
            ChaptPageColltn.Add("204", "Igrave")
            ChaptPageColltn.Add("205", "Iacute")
            ChaptPageColltn.Add("206", "Icirc")
            ChaptPageColltn.Add("207", "Iuml")
            ChaptPageColltn.Add("208", "ETH")
            ChaptPageColltn.Add("209", "Ntilde")
            ChaptPageColltn.Add("210", "Ograve")
            ChaptPageColltn.Add("211", "Oacute")
            ChaptPageColltn.Add("212", "Ocirc")
            ChaptPageColltn.Add("213", "Otilde")
            ChaptPageColltn.Add("214", "Ouml")
            ChaptPageColltn.Add("215", "times")
            ChaptPageColltn.Add("216", "Oslash")
            ChaptPageColltn.Add("217", "Ugrave")
            ChaptPageColltn.Add("218", "Uacute")
            ChaptPageColltn.Add("219", "Ucirc")
            ChaptPageColltn.Add("220", "Uuml")
            ChaptPageColltn.Add("221", "Yacute")
            ChaptPageColltn.Add("222", "THORN")
            ChaptPageColltn.Add("223", "szlig")
            ChaptPageColltn.Add("224", "agrave")
            ChaptPageColltn.Add("225", "aacute")
            ChaptPageColltn.Add("226", "acirc")
            ChaptPageColltn.Add("227", "atilde")
            ChaptPageColltn.Add("228", "auml")
            ChaptPageColltn.Add("229", "aring")
            ChaptPageColltn.Add("230", "aelig")
            ChaptPageColltn.Add("231", "ccedil")
            ChaptPageColltn.Add("232", "egrave")
            ChaptPageColltn.Add("233", "eacute")
            ChaptPageColltn.Add("234", "ecirc")
            ChaptPageColltn.Add("235", "euml")
            ChaptPageColltn.Add("236", "igrave")
            ChaptPageColltn.Add("237", "iacute")
            ChaptPageColltn.Add("238", "icirc")
            ChaptPageColltn.Add("239", "iuml")
            ChaptPageColltn.Add("240", "eth")
            ChaptPageColltn.Add("241", "ntilde")
            ChaptPageColltn.Add("242", "ograve")
            ChaptPageColltn.Add("243", "oacute")
            ChaptPageColltn.Add("244", "ocirc")
            ChaptPageColltn.Add("245", "otilde")
            ChaptPageColltn.Add("246", "ouml")
            ChaptPageColltn.Add("247", "divide")
            ChaptPageColltn.Add("248", "oslash")
            ChaptPageColltn.Add("249", "ugrave")
            ChaptPageColltn.Add("250", "uacute")
            ChaptPageColltn.Add("251", "ucirc")
            ChaptPageColltn.Add("252", "uuml")
            ChaptPageColltn.Add("253", "yacute")
            ChaptPageColltn.Add("254", "thorn")
            ChaptPageColltn.Add("255", "yuml")
            ChaptPageColltn.Add("402", "fnof")
            ChaptPageColltn.Add("913", "Alpha")
            ChaptPageColltn.Add("914", "Beta")
            ChaptPageColltn.Add("915", "Gamma")
            ChaptPageColltn.Add("916", "Delta")
            ChaptPageColltn.Add("917", "Epsilon")
            ChaptPageColltn.Add("918", "Zeta")
            ChaptPageColltn.Add("919", "Eta")
            ChaptPageColltn.Add("920", "Theta")
            ChaptPageColltn.Add("921", "Iota")
            ChaptPageColltn.Add("922", "Kappa")
            ChaptPageColltn.Add("923", "Lambda")
            ChaptPageColltn.Add("924", "Mu")
            ChaptPageColltn.Add("925", "Nu")
            ChaptPageColltn.Add("926", "Xi")
            ChaptPageColltn.Add("927", "Omicron")
            ChaptPageColltn.Add("928", "Pi")
            ChaptPageColltn.Add("929", "Rho")
            ChaptPageColltn.Add("931", "Sigma")
            ChaptPageColltn.Add("932", "Tau")
            ChaptPageColltn.Add("933", "Upsilon")
            ChaptPageColltn.Add("934", "Phi")
            ChaptPageColltn.Add("935", "Chi")
            ChaptPageColltn.Add("936", "Psi")
            ChaptPageColltn.Add("937", "Omega")
            ChaptPageColltn.Add("945", "alpha")
            ChaptPageColltn.Add("946", "beta")
            ChaptPageColltn.Add("947", "gamma")
            ChaptPageColltn.Add("948", "delta")
            ChaptPageColltn.Add("949", "epsilon")
            ChaptPageColltn.Add("950", "zeta")
            ChaptPageColltn.Add("951", "eta")
            ChaptPageColltn.Add("952", "theta")
            ChaptPageColltn.Add("953", "iota")
            ChaptPageColltn.Add("954", "kappa")
            ChaptPageColltn.Add("955", "lambda")
            ChaptPageColltn.Add("956", "mu")
            ChaptPageColltn.Add("957", "nu")
            ChaptPageColltn.Add("958", "xi")
            ChaptPageColltn.Add("959", "omicron")
            ChaptPageColltn.Add("960", "pi")
            ChaptPageColltn.Add("961", "rho")
            ChaptPageColltn.Add("962", "sigmaf")
            ChaptPageColltn.Add("963", "sigma")
            ChaptPageColltn.Add("964", "tau")
            ChaptPageColltn.Add("965", "upsilon")
            ChaptPageColltn.Add("966", "phi")
            ChaptPageColltn.Add("967", "chi")
            ChaptPageColltn.Add("968", "psi")
            ChaptPageColltn.Add("969", "omega")
            ChaptPageColltn.Add("977", "thetasym")
            ChaptPageColltn.Add("978", "upsih")
            ChaptPageColltn.Add("982", "piv")
            ChaptPageColltn.Add("8226", "bull")
            ChaptPageColltn.Add("8230", "hellip")
            ChaptPageColltn.Add("8242", "prime")
            ChaptPageColltn.Add("8243", "Prime")
            ChaptPageColltn.Add("8254", "oline")
            ChaptPageColltn.Add("8260", "frasl")
            ChaptPageColltn.Add("8472", "weierp")
            ChaptPageColltn.Add("8465", "image")
            ChaptPageColltn.Add("8476", "real")
            ChaptPageColltn.Add("8482", "trade")
            ChaptPageColltn.Add("8501", "alefsym")
            ChaptPageColltn.Add("8592", "larr")
            ChaptPageColltn.Add("8593", "uarr")
            ChaptPageColltn.Add("8594", "rarr")
            ChaptPageColltn.Add("8595", "darr")
            ChaptPageColltn.Add("8596", "harr")
            ChaptPageColltn.Add("8629", "crarr")
            ChaptPageColltn.Add("8656", "lArr")
            ChaptPageColltn.Add("8657", "uArr")
            ChaptPageColltn.Add("8658", "rArr")
            ChaptPageColltn.Add("8659", "dArr")
            ChaptPageColltn.Add("8660", "hArr")
            ChaptPageColltn.Add("8704", "forall")
            ChaptPageColltn.Add("8706", "part")
            ChaptPageColltn.Add("8707", "exist")
            ChaptPageColltn.Add("8709", "empty")
            ChaptPageColltn.Add("8711", "nabla")
            ChaptPageColltn.Add("8712", "isin")
            ChaptPageColltn.Add("8713", "notin")
            ChaptPageColltn.Add("8715", "ni")
            ChaptPageColltn.Add("8719", "prod")
            ChaptPageColltn.Add("8721", "sum")
            ChaptPageColltn.Add("8722", "minus")
            ChaptPageColltn.Add("8727", "lowast")
            ChaptPageColltn.Add("8730", "radic")
            ChaptPageColltn.Add("8733", "prop")
            ChaptPageColltn.Add("8734", "infin")
            ChaptPageColltn.Add("8736", "ang")
            ChaptPageColltn.Add("8743", "and")
            ChaptPageColltn.Add("8744", "or")
            ChaptPageColltn.Add("8745", "cap")
            ChaptPageColltn.Add("8746", "cup")
            ChaptPageColltn.Add("8747", "int")
            ChaptPageColltn.Add("8756", "there4")
            ChaptPageColltn.Add("8764", "sim")
            ChaptPageColltn.Add("8773", "cong")
            ChaptPageColltn.Add("8776", "asymp")
            ChaptPageColltn.Add("8800", "ne")
            ChaptPageColltn.Add("8801", "equiv")
            ChaptPageColltn.Add("8804", "le")
            ChaptPageColltn.Add("8805", "ge")
            ChaptPageColltn.Add("8834", "sub")
            ChaptPageColltn.Add("8835", "sup")
            ChaptPageColltn.Add("8836", "nsub")
            ChaptPageColltn.Add("8838", "sube")
            ChaptPageColltn.Add("8839", "supe")
            ChaptPageColltn.Add("8853", "oplus")
            ChaptPageColltn.Add("8855", "otimes")
            ChaptPageColltn.Add("8869", "perp")
            ChaptPageColltn.Add("8901", "sdot")
            ChaptPageColltn.Add("8968", "lceil")
            ChaptPageColltn.Add("8969", "rceil")
            ChaptPageColltn.Add("8970", "lfloor")
            ChaptPageColltn.Add("8971", "rfloor")
            ChaptPageColltn.Add("9001", "lang")
            ChaptPageColltn.Add("9002", "rang")
            ChaptPageColltn.Add("9674", "loz")
            ChaptPageColltn.Add("9824", "spades")
            ChaptPageColltn.Add("9827", "clubs")
            ChaptPageColltn.Add("9829", "hearts")
            ChaptPageColltn.Add("9830", "diams")
            ChaptPageColltn.Add("338", "OElig")
            ChaptPageColltn.Add("339", "oelig")
            ChaptPageColltn.Add("352", "Scaron")
            ChaptPageColltn.Add("353", "scaron")
            ChaptPageColltn.Add("376", "Yuml")
            ChaptPageColltn.Add("710", "circ")
            ChaptPageColltn.Add("732", "tilde")
            ChaptPageColltn.Add("8194", "ensp")
            ChaptPageColltn.Add("8195", "emsp")
            ChaptPageColltn.Add("8201", "thinsp")
            ChaptPageColltn.Add("8204", "zwnj")
            ChaptPageColltn.Add("8205", "zwj")
            ChaptPageColltn.Add("8206", "lrm")
            ChaptPageColltn.Add("8207", "rlm")
            ChaptPageColltn.Add("8211", "ndash")
            ChaptPageColltn.Add("8212", "mdash")
            ChaptPageColltn.Add("8216", "lsquo")
            ChaptPageColltn.Add("8217", "rsquo")
            ChaptPageColltn.Add("8218", "sbquo")
            ChaptPageColltn.Add("8220", "ldquo")
            ChaptPageColltn.Add("8221", "rdquo")
            ChaptPageColltn.Add("8222", "bdquo")
            ChaptPageColltn.Add("8224", "dagger")
            ChaptPageColltn.Add("8225", "Dagger")
            ChaptPageColltn.Add("8240", "permil")
            ChaptPageColltn.Add("8249", "lsaquo")
            ChaptPageColltn.Add("8250", "rsaquo")
            ChaptPageColltn.Add("8364", "euro")

            ProgressBar1.Visible = True
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = ChaptPageColltn.Count



            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                Dim Fpath As String = OpenFileDialog1.FileName
                Dim FileInput As String = File.ReadAllText(Fpath)

                For Each Id As String In ChaptPageColltn.Keys

                    If ProgressBar1.Value > ChaptPageColltn.Count Then
                        ProgressBar1.Value = ProgressBar1.Value - 10
                    Else
                        ProgressBar1.Value = ProgressBar1.Value + 1
                    End If

                    FileInput = Regex.Replace(FileInput, "&#" & Id.ToString & ";", "&" & ChaptPageColltn.Item(Id) & ";", RegexOptions.IgnoreCase Or RegexOptions.Singleline)

                Next
               
                File.WriteAllText(Fpath, FileInput)

                MsgBox("Charater Entity conversion is completed.", MsgBoxStyle.Information, "Entity Conversion")

                ProgressBar1.Value = 0
                ProgressBar1.Visible = False

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Character Entity Conversion")
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim jhg As String = ComboBox1.SelectedText.ToString()

        'If ComboBox1.Items.Count = CInt(ComboBox1.SelectedIndex.ToString) + 1 Then
        'MsgBox("ok")

        'End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim jhg As String = ComboBox1.SelectedValueChanged
    End Sub
End Class


'Dim gg As Match = Regex.Match("", "", RegexOptions.IgnoreCase)
'If gg.Success Then
'    MsgBox("ok")
'End If
'If Regex.IsMatch("ip", "i", RegexOptions.IgnoreCase) Then
'    MsgBox("ok")
'End If


'msword.ActiveDocument.Save()

'*********code for bold , italic 
'With msword.ActiveDocument.Range.Find '.Tables(1).Range.Find
'    .ClearFormatting()
'    With .Font
'        .Bold = True
'        .Italic = True
'        .Underline = WdUnderline.wdUnderlineSingle
'    End With
'    .Text = ""
'    .Replacement.Text = ""
'    .Forward = True
'    .Wrap = WdFindWrap.wdFindContinue
'    .Format = True
'    .MatchCase = False
'    .MatchWholeWord = False
'    .MatchWildcards = False
'    .MatchSoundsLike = False
'    .MatchAllWordForms = False
'    .Execute()
'    If .Found Then
'        MsgBox(msword.ActiveDocument.Selection.Text)
'    End If
'End With

'*********End**********

'code for find format (bold,italic,bolditalic)
'msword.Selection.Find.ClearFormatting()
'msword.Selection.Find.Font.Bold = True
'With msword.Selection.Find
'    .Text = ""
'    .Replacement.Text = "<bold>" & msword.Selection.Text.ToString() & "</bold>"
'    'To get selection string
'    'msword.Selection.Text.ToString()

'    .Forward = True
'    .Wrap = WdFindWrap.wdFindContinue
'    .Format = True
'    .MatchCase = False
'    .MatchWholeWord = False
'    .MatchWildcards = False
'    .MatchSoundsLike = False
'    .MatchAllWordForms = False
'    msword.Selection.Find.Execute()
'End With
'msword.Selection.Find.Execute()

'Code for find styles
'Dim kj As String = ""
'If msword.ActiveDocument.Styles.Count > 0 Then
'    For i As Integer = 1 To msword.ActiveDocument.Styles.Count - 1
'        kj &= msword.ActiveDocument.Styles.Item(i).NameLocal.ToString & vbNewLine
'        Next
'End If

'msword.FontNames.Application.ToString()

'Code for replace text
'Dim rng As Word.Range
'rng = msword.ActiveDocument.Content

'msword.ActiveDocument.FormattingShowFont.ToString()

'rng.Find.Font.Bold = True

'With rng.Find
'    '.Font.Italic = True
'    .ClearFormatting()
'    .Execute(FindText:="November", _
'    ReplaceWith:="xxxxx", _
'    Replace:=Word.WdReplace.wdReplaceAll)

'    .Execute(FindText:=Font.Bold, _
'    ReplaceWith:="yyyyy", _
'    Replace:=Word.WdReplace.wdReplaceAll)


'    '.Execute(FindText:=.Font.Bold, ReplaceWith:="<bold>", _
'    '         Replace:=Word.WdReplace.wdReplaceAll)
'End With

'Marshal.ReleaseComObject(rng)

'msword.ActiveDocument.ConvertNumbersToText()

'msword.ActiveDocument.Save()

'Dim nn As Words = Nothin

'Dim kk As String = nn.Application.ActiveDocument.Name

'Dim jlj As String = nn.Application.ActiveDocument.Words.Count.ToString()


'code for check usb device is inserted
'For Each drv As IO.DriveInfo In IO.DriveInfo.GetDrives
'    Debug.WriteLine(drv.Name & " " & drv.DriveType.ToString)
'Next
