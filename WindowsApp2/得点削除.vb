Imports System.Data.SqlClient

Public Class 得点削除
    Inherits System.Windows.Forms.Form
    Public Sub New(ByVal cmb1 As String)
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()
        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        ComboBox1.Text = cmb1
    End Sub
    Private Sub 得点削除_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Function Delete_data(ByVal strSubject As String, ByVal Hito_Code As String) As Boolean
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try
                Dim SqlUpdate As String = ""
                SqlUpdate &= "UPDATE"
                SqlUpdate &= " 得点"
                SqlUpdate &= " SET"
                SqlUpdate &= " 得点 = NULL"
                SqlUpdate &= " ,Flag = '" & 0 & "'"
                SqlUpdate &= " WHERE 科目コード = '" & strSubject & "'"
                SqlUpdate &= " AND 人コード = '" & Hito_Code & "'"


                'SQLコマンド設定
                Dim command As New SqlCommand(SqlUpdate, sqlconn)
                command.Connection = sqlconn
                command.ExecuteNonQuery()

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


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim array_SubjectCode(4) As String
        Dim strHito_Code As String = ""
        Dim bRet As Boolean = False
        strHito_Code = ComboBox1.Text
        strHito_Code = strHito_Code.Substring(0, 2)


        If CheckBox1.Checked = True Then
            array_SubjectCode(0) = "01"

            bRet = Delete_data(array_SubjectCode(0), strHito_Code)
            If bRet = False Then
                MessageBox.Show("データの削除ができませんでした。")
            End If
        End If
        If CheckBox2.Checked = True Then
                array_SubjectCode(1) = "02"
                bRet = Delete_data(array_SubjectCode(1), strHito_Code)
            If bRet = False Then
                MessageBox.Show("データの削除ができませんでした。")
            End If
        End If
        If CheckBox3.Checked = True Then
                array_SubjectCode(2) = "03"
                bRet = Delete_data(array_SubjectCode(2), strHito_Code)
            If bRet = False Then
                MessageBox.Show("データの削除ができませんでした。")
            End If
        End If
        If CheckBox4.Checked = True Then
                array_SubjectCode(3) = "04"
                bRet = Delete_data(array_SubjectCode(3), strHito_Code)
            If bRet = False Then
                MessageBox.Show("データの削除ができませんでした。")
            End If
        End If
        If CheckBox5.Checked = True Then
            array_SubjectCode(4) = "05"
            bRet = Delete_data(array_SubjectCode(4), strHito_Code)
            If bRet = False Then
                MessageBox.Show("データの削除ができませんでした。")
            End If
        End If

        MessageBox.Show("データを削除しました。")
    End Sub
End Class