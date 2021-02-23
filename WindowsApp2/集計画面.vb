Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Xml

Public Class 集計画面
    Private tbl2 As DataTable

    Private Sub 集計画面_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bRet As Boolean = False

        '初期状態でID順にチェック
        RadioButton1.Checked = True

        '生徒数をラベルに表示する処理
        bRet = Count_Students_Data()
        If bRet = False Then
            MessageBox.Show("データを取得できませんでした。")
            Label1.Text = 0
        End If
        'データグリッドビュー2にを表示する関数を実行
        Call Disp_DGV2()

        'データグリッドビュー1に表示するデータを取得
        '= {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"} '配列

        Dim Jap As String = ""
        Dim Math1 As String = ""
        Dim Eng As String = ""
        Dim Sience As String = ""
        Dim SocialS As String = ""
        Dim Total As Integer = 0
        Dim Ave As Decimal = 0

        Dim strKname As String = ""
        Dim dtJapR As New DataTable '国語の順位
        Dim dtSuugakuR As New DataTable '数学の順位
        Dim dtEngR As New DataTable '英語の順位
        Dim dtSieR As New DataTable '理科の順位
        Dim dtSsR As New DataTable '社会の順位
        Dim Flag1 As Integer = 0
        Dim dtKamokuNm As New DataTable
        Dim HCD() As String
        Dim strKamokuNames() As String
        '人コードを配列に取得
        bRet = Get_HitoCode(HCD)

        'データグリッドビュー1に表示するデータテーブル
        Dim tbl As DataTable = New DataTable("Score")
        Dim row As DataRow

        bRet = Get_KamokuNm(dtKamokuNm)

        tbl.Columns.Add("ID")
        tbl.Columns.Add("氏名")
        For ii As Integer = 0 To dtKamokuNm.Rows.Count - 1
            strKname = dtKamokuNm.Rows(ii).Item(0).ToString
            tbl.Columns.Add(strKname)
        Next

        tbl.Columns.Add("合計", Type.GetType("System.Int32"))
        tbl.Columns.Add("平均")
        tbl.Columns.Add("総合順位")

        ' Call S_Rank(dtJapR, dtSuugakuR, dtEngR, dtSieR, dtSsR)

        'データグリッドビュー1に格納する値がNULLの時に空白を格納する設定
        DataGridView1.DefaultCellStyle.NullValue = " "

        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            '生徒数分、繰り返し
            For Each b As String In HCD
                Try
                    'SQL作成
                    Dim SqlSelect As String = ""
                    SqlSelect &= "SELECT"
                    SqlSelect &= " 姓"
                    SqlSelect &= " ,名"
                    SqlSelect &= " ,得点"
                    SqlSelect &= " ,t.科目コード"
                    SqlSelect &= " ,科目名"
                    SqlSelect &= " FROM 得点 as t"
                    SqlSelect &= " INNER JOIN 氏名 as s"
                    SqlSelect &= " ON t.人コード = s.人コード"
                    SqlSelect &= " INNER JOIN 科目 as k"
                    SqlSelect &= " ON t.科目コード = k.科目コード"
                    SqlSelect &= " WHERE t.人コード = '" & b & "'"
                    SqlSelect &= " ORDER BY t.人コード asc"

                    'SQL実行
                    Dim command As New SqlCommand(SqlSelect, sqlconn)
                    Dim adapter As New SqlDataAdapter(command)
                    Dim dtNameS As New DataTable
                    adapter.Fill(dtNameS)

                    ReDim strKamokuNames(dtNameS.Rows.Count - 1)
                    For ii As Integer = 0 To dtNameS.Rows.Count - 1
                        strKamokuNames(ii) = dtNameS.Rows(ii).Item(2).ToString

                    Next
                    Jap = dtNameS.Rows(0).Item(2).ToString
                    Math1 = dtNameS.Rows(1).Item(2).ToString
                    Eng = dtNameS.Rows(2).Item(2).ToString
                    Sience = dtNameS.Rows(3).Item(2).ToString
                    SocialS = dtNameS.Rows(4).Item(2).ToString

                    '行を初期化
                    row = tbl.NewRow

                    row("ID") = b
                    row("氏名") = dtNameS.Rows(0).Item(0).ToString & dtNameS.Rows(0).Item(1).ToString
                    For ii As Integer = 0 To dtNameS.Rows.Count - 1
                        strKamokuNames(ii) = dtNameS.Rows(ii).Item(2).ToString
                    Next
                    row("国語") = Jap
                    row("数学") = Math1
                    row("英語") = Eng
                    row("理科") = Sience
                    row("社会") = SocialS
                    'row("合計") = Total
                    'row("平均") = Ave
                    tbl.Rows.Add(row)

                Catch ex As Exception
                    Throw
                Finally
                    sqlconn.Close()
                End Try
            Next
            'データテーブルをデータグリッドビューにセット
            Me.DataGridView1.DataSource = tbl

            'データグリッドビューの表示を揃える
            DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridView1.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        '合計点と平均点を表示するためのメソッド
        bRet = Total_ave_score(tbl)
        If bRet = False Then
            MessageBox.Show("一部データの表示ができません。")
        End If
        '順位を保持する
        Dim Rank1 As Integer = 0
        'データビュー
        Dim dv As New DataView(tbl)
        'データビューの行
        Dim drv As DataRowView
        'データビュー上で行を降順にソート
        dv.Sort = "合計 DESC"
        'tblのスキーマをtbl2としてインスタンス化
        tbl2 = tbl.Clone()

        'データビューの行をDataViewRowとしてtbl2にインポート
        For Each drv In dv
            tbl2.ImportRow(drv.Row)
        Next
        'データグリッドビューにtbl2()をセット
        Me.DataGridView1.DataSource = tbl2

        For ii As Integer = 0 To tbl2.Rows.Count - 1
            Rank1 += 1
            'tbl2.Rows(ii).Item("総合順位") = Rank1
            For i As Integer = 0 To tbl2.Rows.Count - 1
                If tbl2.Rows(i).Item("合計") IsNot DBNull.Value And tbl2.Rows(ii).Item("合計") IsNot DBNull.Value Then

                    If tbl2.Rows(ii).Item("合計") = tbl2.Rows(i).Item("合計") Then
                        tbl2.Rows(i).Item("総合順位") = Rank1
                    End If
                End If
            Next
            'If tbl2.Rows(ii).Item("合計") = 0 Or tbl2.Rows(ii).Item("平均") Is DBNull.Value Then
            '    tbl2.Rows(ii).Item("総合順位") = tbl2.Rows.Count
            'End If
        Next
        'データグリッドビューにtbl2をセット
        Me.DataGridView1.DataSource = tbl2
        tbl2.DefaultView.Sort = "ID"

        'データグリッドビュー1のヘッダーのサイズを設定
        DataGridView1.Columns(0).Width = 20 'ID
        DataGridView1.Columns(1).Width = 90 '氏名
        DataGridView1.Columns(2).Width = 50 '国語
        DataGridView1.Columns(3).Width = 50 '数学
        DataGridView1.Columns(4).Width = 50 '英語
        DataGridView1.Columns(5).Width = 50 '理科
        DataGridView1.Columns(6).Width = 50 '社会
        DataGridView1.Columns(7).Width = 50
        DataGridView1.Columns(8).Width = 50
        DataGridView1.Columns(9).Width = 60
        DataGridView1.ColumnHeadersHeight = 25
        'データグリッドビュー1のヘッダーの色を設定
        DataGridView1.Columns(2).HeaderCell.Style.BackColor = Color.Pink
        DataGridView1.Columns(3).HeaderCell.Style.BackColor = Color.Aqua
        DataGridView1.Columns(4).HeaderCell.Style.BackColor = Color.MediumOrchid
        DataGridView1.Columns(5).HeaderCell.Style.BackColor = Color.LimeGreen
        DataGridView1.Columns(6).HeaderCell.Style.BackColor = Color.Gold
        'データグリッドビュー1のセルの色を設定
        DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.LavenderBlush
        DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.PaleTurquoise
        DataGridView1.Columns(4).DefaultCellStyle.BackColor = Color.Plum
        DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.PaleGreen
        DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.LemonChiffon

        'データグリッドビュー2のヘッダーのサイズを設定
        DataGridView2.Columns(0).Width = 50 '
        DataGridView2.Columns(1).Width = 70 '氏名
        DataGridView2.Columns(2).Width = 70 '国語
        DataGridView2.Columns(3).Width = 70 '数学
        DataGridView2.Columns(4).Width = 70 '英語
        DataGridView2.Columns(5).Width = 70 '理科
        DataGridView2.ColumnHeadersHeight = 25
        'データグリッドビュー2のヘッダーの色を設定
        DataGridView2.Columns(1).HeaderCell.Style.BackColor = Color.Pink
        DataGridView2.Columns(2).HeaderCell.Style.BackColor = Color.Aqua
        DataGridView2.Columns(3).HeaderCell.Style.BackColor = Color.MediumOrchid
        DataGridView2.Columns(4).HeaderCell.Style.BackColor = Color.LimeGreen
        DataGridView2.Columns(5).HeaderCell.Style.BackColor = Color.Gold


        'ヘッダーの横の▲を消す処理
        For Each c As DataGridViewColumn In DataGridView1.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c
        Me.DataGridView1.CurrentCell = Nothing

    End Sub

    Private Function Get_KamokuNm(ByRef dtKamokuNm As DataTable) As Boolean
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " 科目名"
                SqlSelect &= " FROM 科目"
                SqlSelect &= " ORDER BY 科目コード asc"
                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                'Dim dtHcode As New DataTable
                adapter.Fill(dtKamokuNm)


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

    '********************************************************
    '【名　称】  S_Rank
    '【機　能】  データグリッドビュー2に順位を表示する関数
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/08
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
    Private Sub S_Rank(ByRef dtJapR As DataTable, ByRef dtSuugakuR As DataTable, ByRef dtEngR As DataTable, ByRef dtSieR As DataTable, ByRef dtSsR As DataTable)
        Dim strX As String = ""
        Dim intRank1 As Integer = 0
        Dim CodeOfSubject As String = ""
        Dim bRet As Boolean = False
        Dim dtKamoku As New DataTable

        bRet = Get_KamokuCode(dtKamoku)
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
                            dtJapR.Rows(i).Item(0) = intRank1
                        End If
                    Next (i)
                Next (ii)
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
                            dtSuugakuR.Rows(i).Item(0) = intRank1
                        End If

                    Next (i)
                Next (ii)

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
                            dtEngR.Rows(i).Item(0) = intRank1
                        End If

                    Next (i)
                Next (ii)

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
                            dtSieR.Rows(i).Item(0) = intRank1
                        End If

                    Next (i)
                Next (ii)
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
                            dtSsR.Rows(i).Item(0) = intRank1
                        End If

                    Next (i)
                Next (ii)

            Catch ex As Exception
                Throw
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    '********************************************************
    '【名　称】  Get_KamokuCode
    '【機　能】  科目コードを取得
    '【備　考】
    '【作成者】  氏名　水落　翔太
    '【作成日】  2020/06/08
    '【修正履歴】YYYY/MM/DD 修正内容　修正者
    '******************************************************** 
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


    'データグリッドビュー１のソート順を変更するメソッド
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton2.Checked = True Then
            tbl2.DefaultView.Sort = "合計 desc"
        Else
            'RadioButton1.Checked = True
            tbl2.DefaultView.Sort = "ID"
        End If
    End Sub


    Private Function Total_ave_score(ByRef tbl2 As DataTable) As Boolean
        Dim CodeOfSubject As String = ""
        Dim Hito_code1 As Integer = "01"
        Dim Hito_code2 As String
        Dim bRet As Boolean = False
        Dim dtSuugakuR As New DataTable '数学の順位
        Dim dtEngR As New DataTable '英語の順位
        Dim dtSieR As New DataTable '理科の順位
        Dim dtSsR As New DataTable '社会の順位
        Dim intMax_score As Integer
        Dim intMin_score As Integer
        Dim Id As String
        Dim HCD() As String = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"} '配列
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()
            Try
                For Each b As String In HCD
                    'SQL作成
                    Dim SqlSelect As String = ""
                    SqlSelect &= " SELECT"
                    SqlSelect &= " 人コード"
                    SqlSelect &= " ,得点"
                    SqlSelect &= " FROM 得点"
                    SqlSelect &= " WHERE 人コード = '" & b & "'"

                    'SQL実行
                    Dim command As New SqlCommand(SqlSelect, sqlconn)
                    Dim adapter As New SqlDataAdapter(command)
                    Dim dtSum As New DataTable '合計値
                    adapter.Fill(dtSum)
                    bRet = Check_nulls(dtSum)

                    Dim Cnt1 As Integer = 0
                    Dim Total1 As Integer = 0
                    Dim Avrg1 As Decimal = 0

                    For ii As Integer = 0 To dtSum.Rows.Count - 1
                        If dtSum.Rows(ii).Item(1) Is DBNull.Value Then '得点データがNULLの場合、0を代入する
                            dtSum.Rows(ii).Item(1) = 0
                        Else
                            Cnt1 += 1　'データがNULLでない科目数のカウント
                        End If
                        Total1 += dtSum.Rows(ii).Item(1)
                    Next



                    Hito_code2 = dtSum.Rows(0).Item(0).ToString
                    For i As Integer = 0 To tbl2.Rows.Count - 1
                        Id = tbl2.Rows(i).Item("ID").ToString

                        If Hito_code2 = Id Then
                            tbl2.Rows(i).Item(7) = Total1 '人コードが一致したら合計点を代入
                            If Cnt1 > 0 Then '得点データがNULLのものが一つもない場合、平均点を算出する
                                Avrg1 = Total1 / Cnt1
                                Avrg1 = Format(Avrg1, "0.0")
                                tbl2.Rows(i).Item(8) = Avrg1
                            End If
                            Exit For
                        End If
                    Next
                Next
            Catch ex As Exception
                Throw
                Return False
            Finally
                sqlconn.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return True
    End Function

    Private Function Check_nulls(ByRef dtSum As DataTable) As Boolean
        Dim cnt_Number As Integer = 0

        For ii As Integer = 0 To dtSum.Rows.Count - 1
            If dtSum.Rows(ii).Item(1) Is DBNull.Value Then
                cnt_Number += 1
            End If
        Next
        If cnt_Number = 0 Then　'NULLが0の場合にtrueを返す
            Return True
        End If
        Return False
    End Function

    ''データグリッドビュー2を表示する
    'Private Sub Disp_DGV2()
    '    'データグリッドビュー2に表示するデータテーブル
    '    Dim dtScores2 As DataTable = New DataTable("Score2")
    '    Dim row As DataRow
    '    dtScores2.Columns.Add("点数")
    '    dtScores2.Columns.Add("国語")
    '    dtScores2.Columns.Add("数学")
    '    dtScores2.Columns.Add("英語")
    '    dtScores2.Columns.Add("理科")
    '    dtScores2.Columns.Add("社会")

    '    Dim ScoreMAX_ALL As Integer = 0
    '    Dim ScoreMINI_ALL As Integer = 0
    '    Dim ScoreAverage As Decimal = 0
    '    Dim CodeOfSubject As String = ""
    '    Dim NameS As String = ""
    '    Dim Jap1(2) As Decimal
    '    Dim Suugaku1(2) As Decimal
    '    Dim Eng1(2) As Decimal
    '    Dim Sie1(2) As Decimal
    '    Dim SocialS1(2) As Decimal
    '    Dim Title1(2) As String

    '    Title1(0) = "最高点"
    '    Title1(1) = "最低点"
    '    Title1(2) = "平均点"

    '    'データグリッドビュー1に格納する値がNULLの時に空白を格納する設定
    '    DataGridView1.DefaultCellStyle.NullValue = " "
    '    Try
    '        'SQLServerの接続開始
    '        Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
    '        sqlconn.Open()

    '        Try
    '            '教科数で繰り返し
    '            For ii As Integer = 0 To 4
    '                If ii = 0 Then
    '                    CodeOfSubject = "01"
    '                ElseIf ii = 1 Then
    '                    CodeOfSubject = "02"
    '                ElseIf ii = 2 Then
    '                    CodeOfSubject = "03"
    '                ElseIf ii = 3 Then
    '                    CodeOfSubject = "04"
    '                ElseIf ii = 4 Then
    '                    CodeOfSubject = "05"
    '                End If

    '                'SQL作成
    '                Dim SqlSelect As String = ""
    '                SqlSelect &= "SELECT"
    '                SqlSelect &= " 得点"
    '                'SqlSelect &= " ,ISNULL(得点,0)"
    '                SqlSelect &= " FROM 得点"
    '                SqlSelect &= " WHERE 科目コード = '" & CodeOfSubject & "'"
    '                SqlSelect &= " ORDER BY 得点　desc"

    '                'SQL実行
    '                Dim command As New SqlCommand(SqlSelect, sqlconn)
    '                Dim adapter As New SqlDataAdapter(command)
    '                Dim dtScores As New DataTable
    '                adapter.Fill(dtScores)

    '                ScoreMAX_ALL = 0
    '                ScoreMINI_ALL = 0

    '                For i As Integer = 0 To dtScores.Rows.Count - 1
    '                    If dtScores.Rows(i).Item(0) IsNot DBNull.Value Then
    '                        ScoreAverage = ScoreAverage + dtScores.Rows(i).Item(0)
    '                    Else

    '                    End If

    '                    If i = 0 Then
    '                        ScoreMAX_ALL = dtScores.Rows(i).Item(0)

    '                    ElseIf i = dtScores.Rows.Count - 1 And dtScores.Rows(i).Item(0) Is DBNull.Value Then
    '                        'ScoreMINI_ALL = ""
    '                    ElseIf i = dtScores.Rows.Count - 1 And dtScores.Rows(i).Item(0) IsNot DBNull.Value Then
    '                        ScoreMINI_ALL = dtScores.Rows(i).Item(0)
    '                    End If
    '                Next

    '                '平均点を算出
    '                ScoreAverage = ScoreAverage / 20
    '                ScoreAverage = Format(ScoreAverage, "0.0")

    '                If ii = 0 Then
    '                    Jap1(0) = ScoreMAX_ALL
    '                    Jap1(1) = ScoreMINI_ALL
    '                    Jap1(2) = ScoreAverage

    '                ElseIf ii = 1 Then
    '                    Suugaku1(0) = ScoreMAX_ALL
    '                    Suugaku1(1) = ScoreMINI_ALL
    '                    Suugaku1(2) = ScoreAverage

    '                ElseIf ii = 2 Then
    '                    Eng1(0) = ScoreMAX_ALL
    '                    Eng1(1) = ScoreMINI_ALL
    '                    Eng1(2) = ScoreAverage

    '                ElseIf ii = 3 Then
    '                    Sie1(0) = ScoreMAX_ALL
    '                    Sie1(1) = ScoreMINI_ALL
    '                    Sie1(2) = ScoreAverage

    '                ElseIf ii = 4 Then
    '                    SocialS1(0) = ScoreMAX_ALL
    '                    SocialS1(1) = ScoreMINI_ALL
    '                    SocialS1(2) = ScoreAverage
    '                End If

    '            Next
    '        Catch ex As Exception
    '            Throw
    '        Finally
    '            sqlconn.Close()
    '        End Try

    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try

    '    For ii As Integer = 0 To 2
    '        '行を初期化
    '        row = dtScores2.NewRow
    '        row("点数") = Title1(ii)
    '        row("国語") = Jap1(ii)
    '        row("数学") = Suugaku1(ii)
    '        row("英語") = Eng1(ii)
    '        row("理科") = Sie1(ii)
    '        row("社会") = SocialS1(ii)

    '        dtScores2.Rows.Add(row)
    '    Next

    '    Me.DataGridView2.DataSource = dtScores2

    '    'データグリッドビューのデータの表示を揃える
    '    DataGridView2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '    DataGridView2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '    DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '    DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '    DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '    DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

    '    'カラムの横の▲を消す処理
    '    For Each c As DataGridViewColumn In DataGridView2.Columns
    '        c.SortMode = DataGridViewColumnSortMode.NotSortable
    '    Next c
    '    '初期表示でセルにフォーカスが当たっていない状態にする
    '    Me.DataGridView2.CurrentCell = Nothing
    'End Sub


    'Private Sub Create_dgv1()

    '    DataGridView1.ColumnCount = 10　'列数

    '    DataGridView1.Columns(0).HeaderText = "ID"
    '    DataGridView1.Columns(1).HeaderText = "氏名"
    '    DataGridView1.Columns(2).HeaderText = "国語"
    '    DataGridView1.Columns(3).HeaderText = "数学"
    '    DataGridView1.Columns(4).HeaderText = "英語"
    '    DataGridView1.Columns(5).HeaderText = "理科"
    '    DataGridView1.Columns(6).HeaderText = "社会"
    '    DataGridView1.Columns(7).HeaderText = "合計"
    '    DataGridView1.Columns(8).HeaderText = "平均"
    '    DataGridView1.Columns(9).HeaderText = "総合順位"

    '    'データグリッドビュー1のヘッダーの色を設定
    '    DataGridView1.Columns(2).HeaderCell.Style.BackColor = Color.Pink '国語
    '    DataGridView1.Columns(3).HeaderCell.Style.BackColor = Color.Aqua '数学
    '    DataGridView1.Columns(4).HeaderCell.Style.BackColor = Color.MediumOrchid '英語
    '    DataGridView1.Columns(5).HeaderCell.Style.BackColor = Color.LimeGreen '理科
    '    DataGridView1.Columns(6).HeaderCell.Style.BackColor = Color.Gold '社会

    '    DataGridView1.Columns(0).Width = 20 'ID
    '    DataGridView1.Columns(1).Width = 90 '氏名
    '    DataGridView1.Columns(2).Width = 50 '国語
    '    DataGridView1.Columns(3).Width = 50 '数学
    '    DataGridView1.Columns(4).Width = 50 '英語
    '    DataGridView1.Columns(5).Width = 50 '理科
    '    DataGridView1.Columns(6).Width = 50 '社会
    '    DataGridView1.Columns(7).Width = 50 '合計
    '    DataGridView1.Columns(8).Width = 50 '平均
    '    DataGridView1.Columns(9).Width = 60 '総合順位
    '    DataGridView1.ColumnHeadersHeight = 25

    '    'ヘッダーの横の▲を消す処理
    '    For Each c As DataGridViewColumn In DataGridView1.Columns
    '        c.SortMode = DataGridViewColumnSortMode.NotSortable
    '    Next c
    '    Me.DataGridView1.CurrentCell = Nothing
    'End Sub
    'データグリッドビュー2を表示する
    Private Sub Disp_DGV2()
        'データグリッドビュー2に表示するデータテーブル
        Dim dtScores2 As DataTable = New DataTable("Score2")
        Dim row As DataRow
        dtScores2.Columns.Add("点数")
        dtScores2.Columns.Add("国語")
        dtScores2.Columns.Add("数学")
        dtScores2.Columns.Add("英語")
        dtScores2.Columns.Add("理科")
        dtScores2.Columns.Add("社会")

        Dim ScoreMAX_ALL As String = ""
        Dim ScoreMINI_ALL As String = ""

        Dim CodeOfSubject As String = ""
        Dim NameS As String = ""
        Dim Jap1(2) As String
        Dim Suugaku1(2) As String
        Dim Eng1(2) As String
        Dim Sie1(2) As String
        Dim SocialS1(2) As String
        Dim Title1(2) As String

        Dim J_ave As Decimal
        Dim Su_ave As Decimal
        Dim E_ave As Decimal
        Dim Si_ave As Decimal
        Dim Ss_ave As Decimal

        Title1(0) = "最高点"
        Title1(1) = "最低点"
        Title1(2) = "平均点"

        'データグリッドビュー1に格納する値がNULLの時に空白を格納する設定
        DataGridView1.DefaultCellStyle.NullValue = " "
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try
                '教科数で繰り返し
                For ii As Integer = 0 To 4
                    If ii = 0 Then
                        CodeOfSubject = "01"
                    ElseIf ii = 1 Then
                        CodeOfSubject = "02"
                    ElseIf ii = 2 Then
                        CodeOfSubject = "03"
                    ElseIf ii = 3 Then
                        CodeOfSubject = "04"
                    ElseIf ii = 4 Then
                        CodeOfSubject = "05"
                    End If

                    Dim ScoreAverage As Decimal = 0
                    'SQL作成
                    Dim SqlSelect As String = ""
                    SqlSelect &= "SELECT"
                    SqlSelect &= " 得点"
                    'SqlSelect &= " ,ISNULL(得点,0)"
                    SqlSelect &= " FROM 得点"
                    SqlSelect &= " WHERE 科目コード = '" & CodeOfSubject & "'"
                    SqlSelect &= " AND 得点 IS NOT NULL"
                    SqlSelect &= " ORDER BY 得点　desc"

                    'SQL実行
                    Dim command As New SqlCommand(SqlSelect, sqlconn)
                    Dim adapter As New SqlDataAdapter(command)
                    Dim dtScores As New DataTable
                    adapter.Fill(dtScores)

                    For i As Integer = 0 To dtScores.Rows.Count - 1
                        If dtScores.Rows(i).Item(0) IsNot DBNull.Value Then
                            ScoreAverage = ScoreAverage + dtScores.Rows(i).Item(0)
                        Else

                        End If

                        If i = 0 And dtScores.Rows(i).Item(0) IsNot DBNull.Value Then
                            ScoreMAX_ALL = dtScores.Rows(i).Item(0)
                        ElseIf i = 0 And dtScores.Rows(i).Item(0) Is DBNull.Value Then
                            ScoreMAX_ALL = ""
                        End If

                        If i = dtScores.Rows.Count - 1 And dtScores.Rows(i).Item(0) Is DBNull.Value Then
                            ScoreMINI_ALL = ""
                        ElseIf i = dtScores.Rows.Count - 1 Then
                            ScoreMINI_ALL = dtScores.Rows(i).Item(0)
                        End If

                    Next

                    '平均点を算出
                    If ScoreAverage = 0 Then
                    Else
                        ScoreAverage = ScoreAverage / dtScores.Rows.Count
                        ScoreAverage = Format(ScoreAverage, "0.0")
                    End If

                    If ii = 0 Then
                        Jap1(0) = ScoreMAX_ALL
                        Jap1(1) = ScoreMINI_ALL
                        J_ave = ScoreAverage

                    ElseIf ii = 1 Then
                        Suugaku1(0) = ScoreMAX_ALL
                        Suugaku1(1) = ScoreMINI_ALL
                        Su_ave = ScoreAverage

                    ElseIf ii = 2 Then
                        Eng1(0) = ScoreMAX_ALL
                        Eng1(1) = ScoreMINI_ALL
                        E_ave = ScoreAverage

                    ElseIf ii = 3 Then
                        Sie1(0) = ScoreMAX_ALL
                        Sie1(1) = ScoreMINI_ALL
                        Si_ave = ScoreAverage

                    ElseIf ii = 4 Then
                        SocialS1(0) = ScoreMAX_ALL
                        SocialS1(1) = ScoreMINI_ALL
                        Ss_ave = ScoreAverage
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

        For ii As Integer = 0 To 1
            '行を初期化
            row = dtScores2.NewRow
            row("点数") = Title1(ii)
            row("国語") = Jap1(ii)
            row("数学") = Suugaku1(ii)
            row("英語") = Eng1(ii)
            row("理科") = Sie1(ii)
            row("社会") = SocialS1(ii)

            dtScores2.Rows.Add(row)
        Next

        row = dtScores2.NewRow
        row("点数") = Title1(2)
        row("国語") = J_ave
        row("数学") = Su_ave
        row("英語") = E_ave
        row("理科") = Si_ave
        row("社会") = Ss_ave
        dtScores2.Rows.Add(row)
        'dtScores2.Rows(2).Item("国語") = J_ave
        'dtScores2.Rows(2).Item("数学") = Su_ave
        'dtScores2.Rows(2).Item("英語") = E_ave
        'dtScores2.Rows(2).Item("理科") = Si_ave
        'dtScores2.Rows(2).Item("社会") = Ss_ave

        Me.DataGridView2.DataSource = dtScores2

        'データグリッドビューのデータの表示を揃える
        DataGridView2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        'カラムの横の▲を消す処理
        For Each c As DataGridViewColumn In DataGridView2.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c
        '初期表示でセルにフォーカスが当たっていない状態にする
        Me.DataGridView2.CurrentCell = Nothing
    End Sub

    Private Function Count_Students_Data() As Boolean
        Try
            'SQLServerの接続開始
            Dim sqlconn As New SqlConnection(My.Settings.sqlServer)
            sqlconn.Open()

            Try

                'SQL作成
                Dim SqlSelect As String = ""
                SqlSelect &= "SELECT"
                SqlSelect &= " COUNT(人コード)"
                SqlSelect &= " FROM 氏名"

                'SQL実行
                Dim command As New SqlCommand(SqlSelect, sqlconn)
                Dim adapter As New SqlDataAdapter(command)
                Dim dtStudent As New DataTable
                adapter.Fill(dtStudent)

                Label1.Text = dtStudent.Rows(0).Item(0).ToString
                Return True
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
End Class