Imports System.Data.SqlClient

Public Class 科目データマスタ
    Private dtKamoku As New DataTable
    Private Sub 科目データマスタ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bRet As Boolean = False
        TextBox1.Text = ""
        TextBox2.Text = ""

        bRet = Make_ListBox() 'リストボックスを表示
        If bRet = False Then
            MessageBox.Show("リストを表示できませんでした。")
        End If
    End Sub

    '********************************************************
    '【名　称】  Button2
    '【機　能】  新規登録イベントの発生
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim bRet As Boolean = False

        bRet = Insert_Kamoku()
        If bRet = False Then
            Exit Sub
        Else
            MessageBox.Show("科目を登録しました。")
        End If

    End Sub

    '********************************************************
    '【名　称】  Make_ListBox
    '【機　能】  リストボックスにデータを格納する
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Function Make_ListBox() As Boolean
        Dim strKamoku As String = ""
        dtKamoku.Clear()

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try
                'SQL作成
                Dim sql As New System.Text.StringBuilder
                sql.AppendLine("SELECT")
                sql.AppendLine(" *")
                sql.AppendLine(" FROM 科目")

                'SQL実行
                Dim command As New SqlCommand(sql.ToString, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dtKamoku)

                ListBox1.Items.Clear()
                For ii As Integer = 0 To dtKamoku.Rows.Count - 1
                    strKamoku = dtKamoku.Rows(ii).Item(0)
                    strKamoku &= "  " & dtKamoku.Rows(ii).Item(1)
                    ListBox1.Items.Add(strKamoku)
                Next
            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
    End Function

    '********************************************************
    '【名　称】  ListBox1_SelectedIndexChanged
    '【機　能】  リストボックスで選択されたデータをテキストボックスに表示する
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim strIndx As String

        strIndx = ListBox1.SelectedIndex
        If strIndx <> -1 Then
            TextBox1.Text = dtKamoku.Rows(strIndx).Item(0).ToString
            TextBox2.Text = dtKamoku.Rows(strIndx).Item(1).ToString
        End If
    End Sub

    '********************************************************
    '【名　称】  Insert_Kamoku 
    '【機　能】  科目テーブルに科目を新規登録する
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Function Insert_Kamoku() As Boolean
        Dim bRet As Boolean = False
        Dim strMes As String = ""

        '入力チェック
        If TextBox1.Text = "" Then
            MessageBox.Show("科目コードを入力してください！")
            Return False
        End If
        If TextBox2.Text = "" Then
            MessageBox.Show("科目名を入力してください！")
            Return False
        End If

        '科目コードの重複チェック
        bRet = Check_Kcode()
        If bRet = True Then
            strMes = "すでに同じ科目コードが登録されています。" & vbCrLf & "科目コードを変更してください。"
            MessageBox.Show(strMes)
            Return False
        End If

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim strSql As String = ""

                strSql &= ""
                strSql &= " INSERT INTO"
                strSql &= " 科目"
                strSql &= " VALUES" & "(" & "'" & TextBox1.Text & "'" & "," & "'" & TextBox2.Text & "'" & ")"

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
        End Try

        ListBox1.Items.Clear()
        bRet = Make_ListBox() 'リストボックスを表示
        If bRet = False Then
            MessageBox.Show("リストを表示できませんでした。")
        End If

        bRet = Insert_Kamoku_toTokuten()

        Return True
    End Function

    '********************************************************
    '【名　称】  Check_Kcode 
    '【機　能】  科目コードに重複がないかチェックする
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Function Check_Kcode() As Boolean
        Dim strKamoku As String = ""
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try
                'SQL作成
                Dim sql As New System.Text.StringBuilder
                sql.AppendLine("SELECT")
                sql.AppendLine(" *")
                sql.AppendLine(" FROM 科目")

                'SQL実行
                Dim command As New SqlCommand(sql.ToString, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dtKamoku)

                For ii As Integer = 0 To dtKamoku.Rows.Count - 1
                    strKamoku = dtKamoku.Rows(ii).Item(0)
                    If TextBox1.Text = strKamoku Then
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
    End Function

    '********************************************************
    '【名　称】  Button4
    '【機　能】  科目データを削除するイベントの発生
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim bRet As Boolean
        Dim Answer As Long

        If TextBox1.Text = "" Then
            MessageBox.Show("科目コードを入力してください！")
            Exit Sub
        End If
        '科目コードの重複チェック
        bRet = Check_Kcode()
        If bRet = False Then
            MessageBox.Show("該当するデータが見つかりませんでした。")
            Exit Sub
        End If

        Answer = MsgBox("登録されている得点データも削除されます。よろしいですか？", vbYesNo, "確認")
        Select Case Answer

            Case vbYes

            Case vbNo
                Exit Sub
        End Select

        'データ削除メソッドを実行
        bRet = Delete_Kamoku()
        If bRet = False Then
            MessageBox.Show("データを削除できませんでした。")
        Else
            MessageBox.Show("データを削除しました。")
        End If



    End Sub

    Private Function Delete_Kamoku() As Boolean
        Dim bRet As Boolean = False
        Dim strKcode As String = ""
        strKcode = TextBox1.Text

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim strSql As String = ""
                strSql &= ""
                strSql &= " DELETE FROM 科目"
                strSql &= " WHERE 科目コード = '" & strKcode & "'"

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

        'リストボックスの再表示
        ListBox1.Items.Clear()
        bRet = Make_ListBox()
        If bRet = False Then
            MessageBox.Show("リストを表示できませんでした。")
            Return False
        End If

        bRet = Delete_Kamoku_fromTokuten()
        If bRet = False Then
            Return False
        End If

        Return True
    End Function

    '********************************************************
    '【名　称】  Delete_Kamoku_fromTokuten
    '【機　能】  得点テーブルから生徒データを削除するメソッド
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/05
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Function Delete_Kamoku_fromTokuten() As Boolean
        Dim Id1 As String = ""  '学籍番号
        Dim Fam_name As String = ""  '苗字
        Dim Name As String = ""  '名前
        Dim strSql As String = ""
        Dim bRet As Boolean = False
        Id1 = TextBox1.Text

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成

                strSql &= ""
                strSql &= " DELETE FROM 得点"
                strSql &= " WHERE 科目コード = '" & Id1 & "'"

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

    Private Function Get_HitoCode(ByRef array_HitoCodes As String()) As Boolean
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
                SqlSelect &= " ORDER BY 人コード　asc"
                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                Dim dtHcode As New DataTable
                adapter.Fill(dtHcode)

                ReDim array_HitoCodes(dtHcode.Rows.Count - 1)
                For ii As Integer = 0 To dtHcode.Rows.Count - 1
                    array_HitoCodes(ii) = dtHcode.Rows(ii).Item(0).ToString
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

    Private Function Insert_Kamoku_toTokuten() As Boolean
        Dim strId As String = ""
        Dim strKcode As String = ""
        Dim array_HitoCodes() As String
        Dim bRet As Boolean = False

        bRet = Get_HitoCode(array_HitoCodes)

        strKcode = TextBox1.Text
        'Call Get_MaxID(strId)

        For Each b As String In array_HitoCodes

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
                    strSql &= " VALUES" & "('" & strKcode & "'" & "," & "'" & b & "'" & "," & "NULL" & ")"

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
End Class