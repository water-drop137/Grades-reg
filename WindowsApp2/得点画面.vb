Imports System.Data.SqlClient

Public Class 得点編集
    Inherits System.Windows.Forms.Form
    Private Sub 得点画面_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Clear()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Dim strSql As String = ""
        Dim bRet As Boolean = False

        bRet = Get_Names()
        If bRet = False Then
            MessageBox.Show("データが取得できませんでした。")
        End If

    End Sub

    '生徒の氏名データを取得
    Private Function Get_Names() As Boolean
        Dim strSql As String = ""
        Dim strId As String = ""
        Dim strFamiName As String = ""
        Dim strFirstName As String = ""
        Dim strName As String

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
                    strFamiName = dtNameS.Rows(i).Item(1).ToString
                    strFirstName = dtNameS.Rows(i).Item(2).ToString
                    strName = strId & " " & strFamiName & " " & strFirstName

                    ComboBox1.Items.Add(strName)
                Next

            Catch ex As Exception
                Throw
                Return False
            Finally
                sqlconn.Close()
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
        Return True

    End Function

    '得点登録確認画面へ遷移
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim f1 As 得点登録確認 = New 得点登録確認(ComboBox1.Text)
        f1.ShowDialog()
        f1.Dispose()

    End Sub

    '得点入力画面へ遷移
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim f2 As 得点入力画面 = New 得点入力画面(ComboBox1.Text)
        f2.ShowDialog()
        f2.Dispose()

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim f3 As 得点削除 = New 得点削除(ComboBox1.Text)
        f3.ShowDialog()
        f3.Dispose()
    End Sub
End Class