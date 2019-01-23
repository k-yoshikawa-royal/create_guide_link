Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''INIファイルを読み込む。
        Dim dbini As IO.StreamReader
        Dim stCurrentDir As String = System.IO.Directory.GetCurrentDirectory()
        CuDr = stCurrentDir

        If IO.File.Exists(CuDr & "\config.ini") = True Then
            dbini = New IO.StreamReader(CuDr & "\config.ini", System.Text.Encoding.Default)

            For lp1 As Integer = 0 To 21
                Dim tbxn1 As String = "Cf_TextBox" & lp1.ToString

                Dim cs As Control() = Me.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    CType(cs(0), TextBox).Text = dbini.ReadLine
                End If
            Next

            ''メイン作業タブへ
            Me.TabControl1.SelectedTab = TabPage1

            dbini.Close()
            dbini.Dispose()
        Else
            MessageBox.Show("設定ファイルが見つからないか壊れています。", "通知")
            Me.TabControl1.SelectedTab = TabPage2
        End If

    End Sub

    Private Sub ConfigButton1_Click(sender As Object, e As EventArgs) Handles ConfigButton1.Click, ConfigButton2.Click, ConfigButton3.Click, ConfigButton4.Click, ConfigButton5.Click, Button18.Click
        close_save()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        excelRead01()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        excelRead02()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        excelRead03()
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        excelRead04()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        ''データベースと接続
        Call sql_st()

        Dim sql1 As String = "SELECT"
        sql1 &= " `商品ID`"          '0
        sql1 &= ",`類似関連`"          '1
        sql1 &= ",`サイズ関連`"          '2
        sql1 &= ",`セット関連`"          '3
        sql1 &= " FROM `"
        sql1 &= tmptbn01
        sql1 &= "`;"

        Dim dTb1 As DataTable = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                If DRow.Item(0) = Nothing Then
                Else


                End If


            Next
        End If


        ''データベース切断
        Call sql_cl()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        excelRead4()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        initializ_work01("rr")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        excelRead11()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        excelRead12()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        excelRead13()
    End Sub


    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        initializ_work02("ry")
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        initializ_work01("sr")
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        excelRead21()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        excelRead22()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        excelRead23()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        initializ_work02("sy")
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        excelRead41()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        excelRead42()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        excelRead43()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        initializ_work02("my")
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        excelRead31()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        excelRead32()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        excelRead33()
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        excelRead14()
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        excelRead24()
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        excelRead34()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        excelRead44()
    End Sub
End Class
