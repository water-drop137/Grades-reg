Imports System.Data.SqlClient

Public Class 順位
    Private strName As String = ""
    Dim strId As String = ""


    Private Sub Shukei_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Dim strFamiName As String = ""

        Dim Jap As Integer
        Dim Math1 As Integer
        Dim Eng As Integer
        Dim Sience As Integer
        Dim SocialS As Integer
        Dim strX As String = ""
        Dim intRank1 As Integer = 0
        Dim CodeOfSubject As String = ""
        Dim dtJapR As New DataTable '国語の順位
        Dim dtSuugakuR As New DataTable '数学の順位
        Dim dtEngR As New DataTable '英語の順位
        Dim dtSieR As New DataTable '理科の順位
        Dim dtSsR As New DataTable '社会の順位
        'Dim hito_cd As String
        Dim intjap_rank As Integer
        Dim intsuugaku_rank As Integer
        Dim inteng_rank As Integer
        Dim intsie_rank As Integer
        Dim intss_rank As Integer


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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim HCD() As String = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"} '配列
        Dim Jap As Integer
        Dim Math1 As Integer
        Dim Eng As Integer
        Dim Sience As Integer
        Dim SocialS As Integer
        Dim Total As Integer
        Dim Ave As Double
        Dim strX As String = ""
        Dim intRank1 As Integer = 0
        Dim CodeOfSubject As String = ""
        Dim dtJapR As New DataTable '国語の順位
        Dim dtSuugakuR As New DataTable '数学の順位
        Dim dtEngR As New DataTable '英語の順位
        Dim dtSieR As New DataTable '理科の順位
        Dim dtSsR As New DataTable '社会の順位
        'Dim hito_cd As String
        Dim intjap_rank As Integer
        Dim intsuugaku_rank As Integer
        Dim inteng_rank As Integer
        Dim intsie_rank As Integer
        Dim intss_rank As Integer
        Dim strhito_name As String

        Dim strhito_cd As String
        DataGridView1.DefaultCellStyle.NullValue = ""
        strhito_cd = ComboBox1.Text
        strhito_cd = strhito_cd.Substring(0, 2)

        Dim tbl As DataTable = New DataTable("Score")
        Dim row As DataRow
        tbl.Columns.Add("ID")
        tbl.Columns.Add("氏名")
        tbl.Columns.Add("国語順位")
        tbl.Columns.Add("数学順位")
        tbl.Columns.Add("英語順位")
        tbl.Columns.Add("理科順位")
        tbl.Columns.Add("社会順位")
        tbl.Columns.Add("総合順位")
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                'SqlSelect &= " 姓"
                'SqlSelect &= " ,名"
                SqlSelect &= " ISNULL(得点,0)"
                SqlSelect &= " FROM 得点 as t"
                'SqlSelect &= " INNER JOIN 氏名 as s"
                'SqlSelect &= " ON t.人コード = s.人コード"
                SqlSelect &= " WHERE t.人コード = '" & strhito_cd & "'"
                'SqlSelect &= " ORDER BY t.人コード asc"

                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)

                Dim dtNameS As New DataTable
                adapter.Fill(dtNameS)
                'For ii As Integer = 0 To dtNameS.Rows.Count - 1

                '    If dtNameS.Rows(ii).Item(0) Is DBNull.Value Then
                '        Exit Sub
                '    End If
                'Next
                Jap = dtNameS.Rows(0).Item(0)
                Math1 = dtNameS.Rows(1).Item(0)
                Eng = dtNameS.Rows(2).Item(0)
                Sience = dtNameS.Rows(3).Item(0)
                SocialS = dtNameS.Rows(4).Item(0)

            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        CodeOfSubject = "01" '国語

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " 人コード"
                SqlSelect &= " ,ISNULL(得点,0)"
                SqlSelect &= " FROM 得点"
                SqlSelect &= " WHERE 科目コード = '" & CodeOfSubject & "'"
                SqlSelect &= " ORDER BY 得点　desc"
                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dtJapR)

                '順位表示処理
                For ii As Integer = 0 To dtJapR.Rows.Count - 1
                    strX = dtJapR.Rows(ii).Item(1)
                    intRank1 += 1

                    For i As Integer = 0 To dtJapR.Rows.Count - 1

                        If dtJapR.Rows(i).Item(1) = strX Then
                            dtJapR.Rows(i).Item(1) = intRank1

                        End If
                    Next (i)
                Next (ii)
                For iii As Integer = 0 To dtJapR.Rows.Count - 1
                    If dtJapR.Rows(iii).Item(0) = strhito_cd Then
                        intjap_rank = dtJapR.Rows(iii).Item(1)
                    End If
                Next

            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        CodeOfSubject = "02"
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " 人コード"
                SqlSelect &= " ,ISNULL(得点,0)"
                SqlSelect &= " FROM 得点"
                SqlSelect &= " WHERE 科目コード = '" & CodeOfSubject & "'"
                SqlSelect &= " ORDER BY 得点　desc"
                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dtSuugakuR)

                intRank1 = 0
                '順位表示処理
                For ii As Integer = 0 To dtSuugakuR.Rows.Count - 1

                    strX = dtSuugakuR.Rows(ii).Item(1)

                    intRank1 += 1
                    For i As Integer = 0 To dtSuugakuR.Rows.Count - 1

                        If dtSuugakuR.Rows(i).Item(1) = strX Then
                            dtSuugakuR.Rows(i).Item(1) = intRank1
                        End If
                    Next (i)
                Next (ii)
                For iii As Integer = 0 To dtSuugakuR.Rows.Count - 1
                    If dtSuugakuR.Rows(iii).Item(0) = strhito_cd Then
                        intsuugaku_rank = dtSuugakuR.Rows(iii).Item(1)
                    End If
                Next
            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '英語
        CodeOfSubject = "03"
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " 人コード"
                SqlSelect &= " ,ISNULL(得点,0)"
                SqlSelect &= " FROM 得点"
                SqlSelect &= " WHERE 科目コード = '" & CodeOfSubject & "'"
                SqlSelect &= " ORDER BY 得点　desc"
                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dtEngR)

                intRank1 = 0
                '順位表示処理
                For ii As Integer = 0 To dtEngR.Rows.Count - 1

                    strX = dtEngR.Rows(ii).Item(1)

                    intRank1 += 1
                    For i As Integer = 0 To dtEngR.Rows.Count - 1

                        If dtEngR.Rows(i).Item(1) = strX Then
                            dtEngR.Rows(i).Item(1) = intRank1
                        End If

                    Next (i)
                Next (ii)

                For iii As Integer = 0 To dtEngR.Rows.Count - 1
                    If dtEngR.Rows(iii).Item(0) = strhito_cd Then
                        inteng_rank = dtEngR.Rows(iii).Item(1)
                    End If
                Next
            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '理科
        CodeOfSubject = "04"
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " 人コード"
                SqlSelect &= " ,ISNULL(得点,0)"
                SqlSelect &= " FROM 得点"
                SqlSelect &= " WHERE 科目コード = '" & CodeOfSubject & "'"
                SqlSelect &= " ORDER BY 得点　desc"
                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dtSieR)

                intRank1 = 0
                '順位表示処理
                For ii As Integer = 0 To dtSieR.Rows.Count - 1

                    strX = dtSieR.Rows(ii).Item(1)

                    intRank1 += 1
                    For i As Integer = 0 To dtSieR.Rows.Count - 1

                        If dtSieR.Rows(i).Item(1) = strX Then
                            dtSieR.Rows(i).Item(1) = intRank1
                        End If

                    Next (i)
                Next (ii)
                For iii As Integer = 0 To dtSieR.Rows.Count - 1
                    If dtSieR.Rows(iii).Item(0) = strhito_cd Then
                        intsie_rank = dtSieR.Rows(iii).Item(1)
                    End If
                Next
            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        '社会
        CodeOfSubject = "05"
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " 人コード"
                SqlSelect &= " ,ISNULL(得点,0)"
                SqlSelect &= " FROM 得点"
                SqlSelect &= " WHERE 科目コード = '" & CodeOfSubject & "'"
                SqlSelect &= " ORDER BY 得点　desc"
                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dtSsR)

                intRank1 = 0
                '順位表示処理
                For ii As Integer = 0 To dtSsR.Rows.Count - 1

                    strX = dtSsR.Rows(ii).Item(1)
                    intRank1 += 1
                    For i As Integer = 0 To dtSsR.Rows.Count - 1

                        If dtSsR.Rows(i).Item(1) = strX Then
                            dtSsR.Rows(i).Item(1) = intRank1
                        End If

                    Next (i)
                Next (ii)

                For iii As Integer = 0 To dtSsR.Rows.Count - 1
                    If dtSsR.Rows(iii).Item(0) = strhito_cd Then
                        intss_rank = dtSsR.Rows(iii).Item(1)
                    End If
                Next

            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        strhito_name = ComboBox1.Text
        strhito_name = strhito_name.Substring(3)
        '行を初期化
        row = tbl.NewRow

        row("ID") = strhito_cd
        row("氏名") = strhito_name
        row("国語順位") = intjap_rank
        row("数学順位") = intsuugaku_rank
        row("英語順位") = inteng_rank
        row("理科順位") = intsie_rank
        row("社会順位") = intss_rank

        tbl.Rows.Add(row)

        Me.DataGridView1.DataSource = tbl

        'データグリッドビューに表示するデータテーブル
        Dim tbl11 As DataTable = New DataTable("Score2")
        Dim row11 As DataRow
        tbl11.Columns.Add("ID")
        tbl11.Columns.Add("氏名")
        tbl11.Columns.Add("国語")
        tbl11.Columns.Add("数学")
        tbl11.Columns.Add("英語")
        tbl11.Columns.Add("理科")
        tbl11.Columns.Add("社会")
        tbl11.Columns.Add("合計")
        tbl11.Columns.Add("平均")
        tbl11.Columns.Add("総合順位")

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            For Each b As String In HCD
                Try
                    'SQL作成
                    Dim SqlSelect As String = ""
                    SqlSelect &= "SELECT"
                    SqlSelect &= " 姓"
                    SqlSelect &= " ,名"
                    SqlSelect &= " ,ISNULL(得点,0)"
                    SqlSelect &= " FROM 得点 as t"
                    SqlSelect &= " INNER JOIN 氏名 as s"
                    SqlSelect &= " ON t.人コード = s.人コード"
                    SqlSelect &= " WHERE t.人コード = '" & b & "'"
                    SqlSelect &= " ORDER BY t.人コード asc"

                    'SQL実行
                    Dim command As New SqlCommand(SqlSelect, sqlconn)
                    Dim adapter As New SqlDataAdapter(command)
                    Dim dtNameS As New DataTable
                    adapter.Fill(dtNameS)

                    Jap = dtNameS.Rows(0).Item(2)
                    Math1 = dtNameS.Rows(1).Item(2)
                    Eng = dtNameS.Rows(2).Item(2)
                    Sience = dtNameS.Rows(3).Item(2)
                    SocialS = dtNameS.Rows(4).Item(2)
                    Total = Jap + Math1 + Eng + Sience + SocialS
                    Ave = (Total) / 5
                    '平均点を少数第二で四捨五入
                    Ave = Math.Round(Ave, 1, MidpointRounding.AwayFromZero)

                    '行を初期化
                    row11 = tbl11.NewRow

                    row11("ID") = b
                    row11("氏名") = dtNameS.Rows(0).Item(0).ToString & dtNameS.Rows(0).Item(1).ToString
                    row11("国語") = Jap
                    row11("数学") = Math1
                    row11("英語") = Eng
                    row11("理科") = Sience
                    row11("社会") = SocialS
                    row11("合計") = Total
                    row11("平均") = Ave
                    tbl11.Rows.Add(row11)

                Catch ex As Exception
                    Throw
                Finally
                    sqlconn.Close()
                End Try
            Next
            'データテーブルをデータグリッドビューにセット
            Me.DataGridView2.DataSource = tbl11

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '順位を保持する
        Dim Rank1 As Integer = 0
        'データビュー
        Dim dv As New DataView(tbl11)
        'データビューの行
        Dim drv As DataRowView
        'ソート後の行データを格納するデータテーブル
        Dim tbl2 As DataTable

        'データビュー上で行を降順にソート
        dv.Sort = "平均 DESC"
        'tblのスキーマをtbl2としてインスタンス化
        tbl2 = tbl11.Clone()

        'データビューの行をDataViewRowとしてtbl2にインポート
        For Each drv In dv
            tbl2.ImportRow(drv.Row)
        Next
        'データグリッドビューにtbl2をセット
        Me.DataGridView2.DataSource = tbl2

        For ii As Integer = 0 To tbl2.Rows.Count - 1
            Rank1 += 1
            tbl2.Rows(ii).Item("総合順位") = Rank1
            For i As Integer = 0 To tbl2.Rows.Count - 1

                If tbl2.Rows(ii).Item("平均") = tbl2.Rows(i).Item("平均") Then
                    tbl2.Rows(i).Item("総合順位") = Rank1
                End If
            Next
        Next

        'データグリッドビューにtbl2をセット
        Me.DataGridView2.DataSource = tbl2
        tbl2.DefaultView.Sort = "ID"

        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        '総合順位の表示
        Dim Jap2 As Integer
        Dim Suu2 As Integer
        Dim Eng2 As Integer
        Dim Sie2 As Integer
        Dim Ss2 As Integer
        For ii As Integer = 0 To tbl2.Rows.Count - 1
            If tbl2.Rows(ii).Item(0) = tbl.Rows(0).Item(0) Then
                tbl.Rows(0).Item(7) = tbl2.Rows(ii).Item("総合順位")

                'Jap2 = tbl.Rows(0).Item(2)
                'tbl.Rows(0).Item(2) = Jap2 & "(" & tbl2.Rows(ii).Item("国語") & ")"

                'Suu2 = tbl.Rows(0).Item(3)
                'tbl.Rows(0).Item(3) = Suu2 & "(" & tbl2.Rows(ii).Item("数学") & ")"

                'Eng2 = tbl.Rows(0).Item(4)
                'tbl.Rows(0).Item(4) = Eng2 & "(" & tbl2.Rows(ii).Item("英語") & ")"

                'Sie2 = tbl.Rows(0).Item(5)
                'tbl.Rows(0).Item(5) = Sie2 & "(" & tbl2.Rows(ii).Item("理科") & ")"

                'Ss2 = tbl.Rows(0).Item(6)
                'tbl.Rows(0).Item(6) = Ss2 & "(" & tbl2.Rows(ii).Item("社会") & ")"
            End If

        Next

        'カラムの横の▲を消す処理
        For Each c As DataGridViewColumn In DataGridView1.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        'カラムの幅を調整
        DataGridView1.Columns(0).Width = 20 'ID
        DataGridView1.Columns(1).Width = 90 '氏名
        DataGridView1.Columns(2).Width = 100 '国語
        DataGridView1.Columns(3).Width = 100 '数学
        DataGridView1.Columns(4).Width = 100 '英語
        DataGridView1.Columns(5).Width = 100 '理科
        DataGridView1.Columns(6).Width = 100 '社会
        DataGridView1.Columns(7).Width = 80
        'DataGridView1.Columns(8).Width = 50
        'DataGridView1.Columns(9).Width = 60
        DataGridView1.ColumnHeadersHeight = 25
        DataGridView1.Columns(2).HeaderCell.Style.BackColor = Color.Pink
        DataGridView1.Columns(3).HeaderCell.Style.BackColor = Color.Aqua
        DataGridView1.Columns(4).HeaderCell.Style.BackColor = Color.MediumOrchid
        DataGridView1.Columns(5).HeaderCell.Style.BackColor = Color.LimeGreen
        DataGridView1.Columns(6).HeaderCell.Style.BackColor = Color.Gold
        DataGridView1.Columns(7).HeaderCell.Style.BackColor = Color.Crimson

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class