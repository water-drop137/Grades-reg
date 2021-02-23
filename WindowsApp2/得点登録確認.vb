Imports System.Data.SqlClient

Public Class 得点登録確認
    Inherits System.Windows.Forms.Form
    Public Sub New(ByVal cmb1 As String)
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()
        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Label1.Text = cmb1
    End Sub


    Private Sub 得点登録確認_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dtScores As New DataTable
        Disp_Dgv(dtScores)

    End Sub

    Private Function Disp_Dgv(ByRef dtScores As DataTable) As Boolean
        Dim Sql_Select As String = ""
        Dim strId As String
        Dim strJap As String = ""
        Dim strSuugaku As String = ""
        Dim strEng As String = ""
        Dim strSie As String = ""
        Dim strSs As String = ""

        Dim arraySubject() As Integer

        'データグリッドビュー1に格納する値がNULLの時に空白を格納する設定
        DataGridView1.DefaultCellStyle.NullValue = " "
        strId = Label1.Text
        strId = strId.Substring(0, 2)

        dtScores.Columns.Add("国 語")
        dtScores.Columns.Add("数 学")
        dtScores.Columns.Add("英 語")
        dtScores.Columns.Add("理 科")
        dtScores.Columns.Add("社 会")
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成

                Sql_Select &= " SELECT"
                Sql_Select &= " 得点"
                Sql_Select &= " FROM"
                Sql_Select &= " 得点"
                Sql_Select &= " WHERE "
                Sql_Select &= " 人コード = '" & strId & "'"

                'SQL実行
                Dim command As New SqlCommand(Sql_Select, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                Dim dtNameS As New DataTable
                adapter.Fill(dtNameS)

                '実行結果をコンボボックスに格納
                'For i As Integer = 0 To dtNameS.Rows.Count - 1
                Dim row As DataRow
                strJap = dtNameS.Rows(0).Item(0).ToString
                strSuugaku = dtNameS.Rows(1).Item(0).ToString
                strEng = dtNameS.Rows(2).Item(0).ToString
                strSie = dtNameS.Rows(3).Item(0).ToString
                strSs = dtNameS.Rows(4).Item(0).ToString

                row = dtScores.NewRow
                row("国 語") = strJap
                row("数 学") = strSuugaku
                row("英 語") = strEng
                row("理 科") = strSie
                row("社 会") = strSs
                dtScores.Rows.Add(row)
                'Next
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
        'データテーブルをデータグリッドビューにセット
        Me.DataGridView1.DataSource = dtScores
        DataGridView1.ColumnHeadersHeight = 30
        DataGridView1.Rows(0).Height = 47
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("メイリオ", 15, FontStyle.Regular)
        ' メイリオ 9point 太字
        For ii As Integer = 0 To dtScores.Columns.Count - 1
            Me.DataGridView1.Columns(ii).DefaultCellStyle.Font = New Font("メイリオ", 15, FontStyle.Regular)
        Next
        'データグリッドビュー1のヘッダーの色を設定
        DataGridView1.Columns(0).HeaderCell.Style.BackColor = Color.Pink
        DataGridView1.Columns(1).HeaderCell.Style.BackColor = Color.Aqua
        DataGridView1.Columns(2).HeaderCell.Style.BackColor = Color.MediumOrchid
        DataGridView1.Columns(3).HeaderCell.Style.BackColor = Color.LimeGreen
        DataGridView1.Columns(4).HeaderCell.Style.BackColor = Color.Gold
        'ヘッダーの横の▲を消す処理
        For Each c As DataGridViewColumn In DataGridView1.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c
        Me.DataGridView1.CurrentCell = Nothing
        Return True

    End Function

    Private Sub Set_Column()


    End Sub
End Class