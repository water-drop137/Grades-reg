Imports System.Data.SqlClient

Public Class 得点入力画面

    Inherits System.Windows.Forms.Form
    Public Sub New(ByVal cmb1 As String)
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()
        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        ComboBox1.Text = cmb1
    End Sub
    Private strName As String = ""
    Private Sub 得点入力画面_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.MaxLength = 3
        TextBox3.MaxLength = 3
        TextBox4.MaxLength = 3
        TextBox5.MaxLength = 3
        TextBox6.MaxLength = 3

    End Sub

    '文字入力制限
    Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress, TextBox3.KeyPress, TextBox4.KeyPress, TextBox5.KeyPress, TextBox6.KeyPress
        'キーが [0]～[9] または [BackSpace] 以外の場合イベントをキャンセル
        If Not (("0"c <= e.KeyChar And e.KeyChar <= "9"c) Or e.KeyChar = ControlChars.Back) Then
            'コントロールの既定の処理を省略する場合は true
            e.Handled = True
        End If
    End Sub

    '氏名をCBに表示
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.Click

        Dim strFamiName As String = ""
        Dim strId As String = ""

        ComboBox1.Items.Clear()

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try
                'SQL作成
                Dim sql As New System.Text.StringBuilder
                sql.AppendLine("SELECT")
                sql.AppendLine("  * ")
                sql.AppendLine("FROM 氏名 ")

                'SQL実行
                Dim command As New SqlCommand(sql.ToString, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                Dim dtNameS As New DataTable
                adapter.Fill(dtNameS)

                '実行結果をコンボボックスに格納
                For i As Integer = 0 To dtNameS.Rows.Count - 1

                    strId = dtNameS.Rows(i).Item(0).ToString
                    strName = dtNameS.Rows(i).Item(1).ToString
                    strFamiName = dtNameS.Rows(i).Item(2).ToString
                    strName = strId & " " & strName & " " & strFamiName

                    ComboBox1.Items.Add(strName)
                Next

            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

    End Sub

    '登録ボタン押下時
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Jap As String = ""
        Dim Maths1 As String = ""
        Dim Sience As String = ""
        Dim SocialS As String = ""
        Dim English As String = ""
        Dim Simei As String
        Dim HitoCD As String = ""
        Dim bRet As Boolean = False
        Jap = TextBox2.Text
        Maths1 = TextBox3.Text
        Sience = TextBox4.Text
        SocialS = TextBox5.Text
        English = TextBox6.Text

        '入力された得点のチェック
        bRet = CheckTxtbox()
        If bRet = False Then
            Exit Sub
        End If

        If TextBox2.Text = "" And TextBox3.Text = "" And TextBox4.Text = "" And TextBox5.Text = "" And TextBox6.Text = "" Then
            MsgBox("登録データはありません。", vbYes)

        ElseIf TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            Dim Answer As Long
            Answer = MsgBox("未入力の得点がありますが、よろしいですか？", vbYesNo, "確認")

            Select Case Answer

                Case vbYes

                Case vbNo
                    Exit Sub
            End Select
        End If

        Simei = ComboBox1.Text
        Simei = Simei.Substring(0, 2) 'Idデータ

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " 人コード"
                SqlSelect &= " FROM 氏名"
                SqlSelect &= " WHERE 人コード = '" & Simei & "'"

                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                Dim dtNameS As New DataTable
                adapter.Fill(dtNameS)

                '実行結果をコンボボックスに格納
                For i As Integer = 0 To dtNameS.Rows.Count - 1

                    HitoCD = dtNameS.Rows(i).Item(0).ToString
                    ' strName = strName & " " & dtNameS.Rows(i).Item(2).ToString
                    ComboBox1.Items.Add(strName)
                Next

            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try

        If Jap <> "" Then
            Try
                'SQLServerの接続開始
                Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
                sqlconn.Open()
                Try
                    'SQL作成
                    Dim SqlUpdate As String = ""
                    SqlUpdate &= "UPDATE"
                    SqlUpdate &= " 得点"
                    SqlUpdate &= " SET"
                    SqlUpdate &= " 得点 = '" & Jap & "'"
                    SqlUpdate &= " WHERE 人コード = '" & HitoCD & "'"
                    SqlUpdate &= " AND 科目コード = 01 "

                    'SQL実行
                    Dim command As New SqlCommand(SqlUpdate, sqlconn)
                    command.Connection = sqlconn
                    command.ExecuteNonQuery()

                Catch ex As Exception
                    Throw
                Finally
                    sqlconn.Close()
                End Try

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

        If Maths1 <> "" Then
            Try
                'SQLServerの接続開始
                Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
                sqlconn.Open()
                Try
                    'SQL作成
                    Dim SqlUpdate As String = ""
                    SqlUpdate &= "UPDATE"
                    SqlUpdate &= " 得点"
                    SqlUpdate &= " SET"
                    SqlUpdate &= " 得点 = '" & Maths1 & "'"
                    SqlUpdate &= " WHERE 人コード = '" & HitoCD & "'"
                    SqlUpdate &= " AND 科目コード = 02 "

                    'SQL実行
                    Dim command As New SqlCommand(SqlUpdate, sqlconn)
                    command.Connection = sqlconn
                    command.ExecuteNonQuery()

                Catch ex As Exception
                    Throw
                Finally
                    sqlconn.Close()
                End Try

            Catch ex As Exception
                MsgBox(ex.ToString)

            End Try
        End If

        If English <> "" Then
            Try
                'SQLServerの接続開始
                Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
                sqlconn.Open()
                Try
                    'SQL作成
                    Dim SqlUpdate As String = ""
                    SqlUpdate &= "UPDATE"
                    SqlUpdate &= " 得点"
                    SqlUpdate &= " SET"
                    SqlUpdate &= " 得点 = '" & English & "'"
                    SqlUpdate &= " WHERE 人コード = '" & HitoCD & "'"
                    SqlUpdate &= " AND 科目コード = 05 "

                    'SQL実行
                    Dim command As New SqlCommand(SqlUpdate, sqlconn)
                    command.Connection = sqlconn
                    command.ExecuteNonQuery()

                Catch ex As Exception
                    Throw
                Finally
                    sqlconn.Close()
                End Try

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If


        If Sience <> "" Then
            Try
                'SQLServerの接続開始
                Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
                sqlconn.Open()
                Try
                    'SQL作成
                    Dim SqlUpdate As String = ""
                    SqlUpdate &= "UPDATE"
                    SqlUpdate &= " 得点"
                    SqlUpdate &= " SET"
                    SqlUpdate &= " 得点 = '" & Sience & "'"
                    SqlUpdate &= " WHERE 人コード = '" & HitoCD & "'"
                    SqlUpdate &= " AND 科目コード = 03 "

                    'SQL実行
                    Dim command As New SqlCommand(SqlUpdate, sqlconn)
                    command.Connection = sqlconn
                    command.ExecuteNonQuery()

                Catch ex As Exception
                    Throw
                Finally
                    sqlconn.Close()
                End Try
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

        If SocialS <> "" Then
            Try
                'SQLServerの接続開始
                Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
                sqlconn.Open()
                Try
                    'SQL作成
                    Dim SqlUpdate As String = ""
                    SqlUpdate &= "UPDATE"
                    SqlUpdate &= " 得点"
                    SqlUpdate &= " SET"
                    SqlUpdate &= " 得点 = '" & SocialS & "'"
                    SqlUpdate &= " WHERE 人コード = '" & HitoCD & "'"
                    SqlUpdate &= " AND 科目コード = 04 "

                    'SQL実行
                    Dim command As New SqlCommand(SqlUpdate, sqlconn)
                    command.Connection = sqlconn
                    command.ExecuteNonQuery()

                Catch ex As Exception
                    Throw
                Finally
                    sqlconn.Close()
                End Try
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

        MessageBox.Show("登録しました。")

    End Sub

    Private Function CheckTxtbox() As Boolean
        If ComboBox1.Text = "" Then
            MessageBox.Show("氏名を選択してください。")
            Return False
            Exit Function
        End If

        If TextBox2.Text = "" Then

        ElseIf TextBox2.Text > 100 Then
            MsgBox("得点は'0～100'の数値で入力してください。")
            Return False
            Exit Function
        End If

        If TextBox3.Text = "" Then

        ElseIf TextBox3.Text > 100 Then
            MsgBox("得点は'0～100'の数値で入力してください。")
            Return False
            Exit Function
        End If

        If TextBox4.Text = "" Then

        ElseIf TextBox4.Text > 100 Then
            MsgBox("得点は'0～100'の数値で入力してください。")
            Return False
            Exit Function
        End If

        If TextBox5.Text = "" Then

        ElseIf TextBox5.Text > 100 Then
            MsgBox("得点は'0～100'の数値で入力してください。")
            Return False
            Exit Function
        End If

        If TextBox6.Text = "" Then

        ElseIf TextBox6.Text > 100 Then
            MsgBox("得点は'0～100'の数値で入力してください。")
            Return False
            Exit Function
        End If
        Return True
    End Function


End Class