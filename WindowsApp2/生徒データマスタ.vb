Imports System.Data.SqlClient

Public Class 生徒データマスタ
    Private dtNames As New DataTable
    Private Sub 生徒データマスタ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bRet As Boolean = False
        bRet = Disp_ListBox()

    End Sub

    '検索ボタン押下時
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim bRet As Boolean = False

        If TextBox1.Text = "" Then
            MessageBox.Show("学籍番号を入力してください。")
            Exit Sub
        End If

        bRet = Select_ID()
        If bRet = False Then
            MessageBox.Show("未使用の学籍番号です。")
            TextBox2.Text = ""
            TextBox3.Text = ""
        End If
    End Sub

    '学籍番号から氏名を出力するメソッド
    Private Function Select_ID() As Boolean
        'フィールド--------------------------------------------------------------
        Dim strId As String = ""

        '--------------------------------------------------------------

        strId = TextBox1.Text
        strId = Format(CInt(strId), "00")

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try
                'SQL作成
                Dim sql As New System.Text.StringBuilder
                sql.AppendLine("SELECT")
                sql.AppendLine(" 姓")
                sql.AppendLine(" ,名")
                sql.AppendLine(" FROM 氏名 ")
                sql.AppendLine(" WHERE 人コード = '" & strId & "'")
                'SQL実行
                Dim command As New SqlCommand(sql.ToString, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                Dim dtId As New DataTable
                adapter.Fill(dtId)

                TextBox2.Text = dtId.Rows(0).Item(0)
                TextBox3.Text = dtId.Rows(0).Item(1)

            Catch ex As Exception
                Return False
                Throw
            Finally
                sqlconn.Close()
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return True
    End Function

    '********************************************************
    '【名　称】  Button2
    '【機　能】  生徒データを新規登録するイベントを発生
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim bRet As Boolean = False
        Dim strId As String = ""
        If TextBox1.Text = "" Then
            MessageBox.Show("学籍番号が未入力です。")
            Exit Sub
        End If

        If TextBox2.Text = "" Then
            MessageBox.Show("姓が未入力です。")
            Exit Sub
        ElseIf TextBox3.Text = "" Then
            MessageBox.Show("名が未入力です。")
            Exit Sub
        End If

        bRet = Insert_Student()
        If bRet = False Then
            MessageBox.Show("データの登録に失敗しました。")
        Else

            MessageBox.Show("登録しました。")
        End If
    End Sub

    '********************************************************
    '【名　称】  Insert_Student
    '【機　能】  氏名テーブルに生徒データを新規登録する
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Function Insert_Student() As Boolean
        Dim Id1 As Integer = 0  '学籍番号
        Dim Fam_name As String = ""  '苗字
        Dim Name As String = ""  '名前
        Dim strSql As String = ""
        Dim strId As String = ""
        Dim bRet As Boolean = False
        Dim dtKamoku As New DataTable

        Id1 = CInt(TextBox1.Text)　'人コード
        Fam_name = TextBox2.Text　'苗字
        Name = TextBox3.Text　'名前

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim sql As New System.Text.StringBuilder

                strSql &= ""
                strSql &= " INSERT INTO"
                strSql &= " 氏名"
                strSql &= " VALUES" & "(" & "'" & Id1 & "'" & "," & "'" & Fam_name & "'" & "," & "'" & Name & "'" & ")"

                'SQL実行
                Dim command As New SqlCommand(strSql, sqlconn)
                'SQLコマンド設定
                command.Connection = sqlconn
                command.ExecuteNonQuery()

            Catch ex As Exception
                Return False
                Throw
            Finally
                sqlconn.Close()
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try

        bRet = Disp_ListBox()
        If bRet = False Then
            Return False
        End If

        bRet = Insert_Student_toTokuten()
        If bRet = False Then
            MessageBox.Show("データの登録に失敗しました。")
            Return False
        End If

        Return True

    End Function

    Private Function Insert_Student_toTokuten() As Boolean
        'Dim strId As String = ""
        Dim bRet As Boolean = False
        Dim dtKamoku As New DataTable
        Dim Id1 As Integer = 0  '学籍番号

        Id1 = CInt(TextBox1.Text) '人コード

        'Call Get_MaxID(strId)
        'If strId = "" Then
        '    Return False
        'End If

        bRet = Get_KamokuCode(dtKamoku)

        Dim array_Kcode(dtKamoku.Rows.Count - 1) As String

        For ii As Integer = 0 To dtKamoku.Rows.Count - 1
            array_Kcode(ii) = dtKamoku.Rows(ii).Item(0)
        Next

        For Each b As String In array_Kcode

            'strId = strId + 1

            Try
                'SQLServerの接続開始
                Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
                sqlconn.Open()
                Try
                    'SQL作成
                    Dim strSql As String = ""

                    strSql &= ""
                    strSql &= " INSERT INTO"
                    strSql &= " 得点"
                    strSql &= " VALUES" & "('" & b & "'" & "," & "'" & Id1 & "'" & "," & "NULL" & ")"

                    'SQL実行
                    Dim command As New SqlCommand(strSql, sqlconn)
                    'SQLコマンド設定
                    command.Connection = sqlconn
                    command.ExecuteNonQuery()

                Catch ex As Exception
                    Throw
                Finally
                    sqlconn.Close()
                End Try

            Catch ex As Exception
                MsgBox(ex.ToString)
                Return False
            End Try
        Next

        Return True
    End Function
    Private Sub Get_MaxID(ByRef strId As String)
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try
                'SQL作成
                Dim sql As New System.Text.StringBuilder
                sql.AppendLine("SELECT")
                sql.AppendLine(" MAX(Id)")
                sql.AppendLine(" FROM 得点")

                'SQL実行
                Dim command As New SqlCommand(sql.ToString, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                Dim dtMaxId As New DataTable
                adapter.Fill(dtMaxId)
                If dtMaxId.Rows(0).Item(0) <> 0 Then
                    strId = dtMaxId.Rows(0).Item(0)

                End If

            Catch ex As Exception

                Throw
            Finally
                sqlconn.Close()
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub


    Private Function Get_KamokuCode(ByRef dtKamokuCD As DataTable) As Boolean

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " 科目コード"
                SqlSelect &= " FROM 科目"
                SqlSelect &= " ORDER BY 科目コード asc"
                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                'Dim dtKamokuCD As New DataTable
                adapter.Fill(dtKamokuCD)

                If dtKamokuCD.Rows(0).Item(0) = "" Then
                    Return False
                End If
                ''順位表示処理
                'For ii As Integer = 0 To dtJapR.Rows.Count - 1

                '    strX = dtJapR.Rows(ii).Item(1)
                '    intRank1 += 1

                '    For i As Integer = 0 To dtJapR.Rows.Count - 1
                '        If dtJapR.Rows(i).Item(1) = strX Then
                '            dtJapR.Rows(i).Item(0) = intRank1
                '        End If
                '    Next (i)
                'Next (ii)
            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return True
    End Function

    ''********************************************************
    ''【名　称】  Insert_Student2
    ''【機　能】  得点テーブルに生徒データを新規登録する
    ''【備　考】
    ''【作成者】  氏名　水落　翔太
    ''【作成日】  2020/06/05
    ''【修正履歴】YYYY/MM/DD 修正内容　修正者
    ''******************************************************** 
    'Private Function Insert_Student2() As Boolean
    '    Dim Id1 As Integer = 0  '学籍番号
    '    Dim Fam_name As String = ""  '苗字
    '    Dim Name As String = ""  '名前
    '    Dim strSql As String = ""
    '    Dim strT_Id As String = ""
    '    Dim bRet As Boolean = ""


    '    Id1 = CInt(TextBox1.Text)　'人コード
    '    Fam_name = TextBox2.Text　'苗字
    '    Name = TextBox3.Text　'名前

    '    Try
    '        'SQLServerの接続開始
    '        Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
    '        sqlconn.Open()
    '        Try
    '            'SQL作成
    '            Dim sql As New System.Text.StringBuilder

    '            strSql &= ""
    '            strSql &= " INSERT INTO"
    '            strSql &= " 得点"
    '            strSql &= " VALUES" & "(" & "'" & Id1 & "'" & "," & "'" & Fam_name & "'" & "," & "'" & Name & "'" & ")"


    '            'SQL実行
    '            Dim command As New SqlCommand(strSql, sqlconn)
    '            'SQLコマンド設定
    '            command.Connection = sqlconn
    '            command.ExecuteNonQuery()

    '        Catch ex As Exception
    '            Return False
    '            Throw
    '        Finally
    '            sqlconn.Close()
    '        End Try

    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try
    '    bRet = Disp_ListBox()
    '    If bRet = False Then
    '        Return False
    '    End If

    '    Return True
    'End Function


    '********************************************************
    '【名　称】  Button4
    '【機　能】  生徒データを削除するイベントを発生
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim bRet As Boolean = False

        If TextBox1.Text = "" Then
            MessageBox.Show("学籍番号を入力してください。")
            Exit Sub
        End If

        bRet = Delete_Student()
        If bRet = False Then
            MessageBox.Show("データの削除に失敗しました。")
        Else
            MessageBox.Show("削除しました。")
            TextBox2.Text = ""
            TextBox3.Text = ""
        End If
    End Sub

    '********************************************************
    '【名　称】  Delete_Student
    '【機　能】  生徒データを削除するメソッド
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Function Delete_Student() As Boolean
        Dim Id1 As String = "" '学籍番号
        Dim Fam_name As String = ""  '苗字
        Dim Name As String = ""  '名前
        Dim strSql As String = ""
        Dim bRet As Boolean = False

        Id1 = TextBox1.Text
        Fam_name = TextBox2.Text
        Name = TextBox3.Text

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim sql As New System.Text.StringBuilder

                strSql &= ""
                strSql &= " DELETE FROM 氏名"
                strSql &= " WHERE 人コード = '" & Id1 & "'"

                'SQL実行
                Dim command As New SqlCommand(strSql, sqlconn)
                command.Connection = sqlconn
                command.ExecuteNonQuery()

            Catch ex As Exception
                Return False
                Throw
            Finally
                sqlconn.Close()
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        bRet = Delete_Student_fromTokuten()
        If bRet = False Then
            Return False
        End If

        bRet = Disp_ListBox()
        If bRet = False Then
            Return False
        End If

        Return True

    End Function

    '********************************************************
    '【名　称】  Delete_Student
    '【機　能】  得点テーブルから生徒データを削除するメソッド
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Function Delete_Student_fromTokuten() As Boolean
        Dim Id1 As Integer = 0  '学籍番号
        Dim Fam_name As String = ""  '苗字
        Dim Name As String = ""  '名前
        Dim strSql As String = ""
        Dim bRet As Boolean = False
        Id1 = CInt(TextBox1.Text)
        Fam_name = TextBox2.Text
        Name = TextBox3.Text

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成

                strSql &= ""
                strSql &= " DELETE FROM 得点"
                strSql &= " WHERE 人コード = '" & Id1 & "'"

                'SQL実行
                Dim command As New SqlCommand(strSql, sqlconn)
                command.Connection = sqlconn
                command.ExecuteNonQuery()

            Catch ex As Exception
                Return False
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try


        Return True
    End Function
    '********************************************************
    '【名　称】  Disp_ListBox
    '【機　能】  リストボックスに表示する
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Function Disp_ListBox() As Boolean
        Dim strName As String = ""

        dtNames.Clear()

        Try
        'SQLServerの接続開始
        Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try
                'SQL作成
                Dim sql As New System.Text.StringBuilder
                sql.AppendLine("SELECT")
                sql.AppendLine(" *")
                sql.AppendLine(" FROM 氏名 ")

                'SQL実行
                Dim command As New SqlCommand(sql.ToString, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dtNames)

                ListBox1.Items.Clear()

                For ii As Integer = 0 To dtNames.Rows.Count - 1
                    strName = dtNames.Rows(ii).Item(0)
                    strName &= "  " & dtNames.Rows(ii).Item(1)
                    strName &= " " & dtNames.Rows(ii).Item(2)
                    ListBox1.Items.Add(strName)
                Next

            Catch ex As Exception

                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
        Return True
    End Function

    '********************************************************
    '【名　称】  ListBox1_SelectedIndexChanged
    '【機　能】  リストボックス内で選択されたデータをテキストボックスに表示する
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim strIndx As String

        strIndx = ListBox1.SelectedIndex
        TextBox1.Text = dtNames.Rows(strIndx).Item(0).ToString
        TextBox2.Text = dtNames.Rows(strIndx).Item(1).ToString
        TextBox3.Text = dtNames.Rows(strIndx).Item(2).ToString
    End Sub
End Class