Imports System.IO
Imports MySql.Data.MySqlClient
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Module Module1

    Public CuDr As String       ''exeの動くカレントディレクトリを格納
    Public mysqlCon As New MySqlConnection
    Public sqlCommand As New MySqlCommand
    Public tmptbn01 As String     ''テンポラリテーブルを保持
    Public tmptbn02 As String     ''テンポラリテーブルを保持
    Public tmptbn03 As String     ''テンポラリテーブルを保持
    Public tmptbn04 As String     ''テンポラリテーブルを保持

    Sub sql_st()
        ''データベースに接続

        Dim Builder = New MySqlConnectionStringBuilder()
        ' データベースに接続するために必要な情報をBuilderに与える。データベース情報はGitに乗せないこと。
        Builder.Server = ""
        Builder.Port =
        Builder.UserID = ""
        Builder.Password = ""
        Builder.Database = ""
        Dim ConStr = Builder.ToString()

        mysqlCon.ConnectionString = ConStr
        mysqlCon.Open()

    End Sub

    Sub sql_cl()
        ' データベースの切断
        mysqlCon.Close()
    End Sub

    Function sql_result_return(ByVal query As String) As DataTable
        ''データセットを返すSELECT系のSQLを処理するコード

        Dim dt As New DataTable()

        Try
            ' 4.データ取得のためのアダプタの設定
            Dim Adapter = New MySqlDataAdapter(query, mysqlCon)

            ' 5.データを取得
            Dim Ds As New DataSet
            Adapter.Fill(dt)

            Return dt
        Catch ex As Exception

            Return dt
        End Try

    End Function

    Function sql_result_no(ByVal query As String)
        ''データセットを返さない、DELETE、UPDATE、INSERT系のSQLを処理するコード

        Try
            sqlCommand.Connection = mysqlCon
            sqlCommand.CommandText = query
            sqlCommand.ExecuteNonQuery()

            Return "Complete"
        Catch ex As Exception

            Return ex.Message
        End Try

    End Function

    Sub close_save()
        ''設定用ファイルの保存

        Dim dtx1 As String = ""

        For lp1 As Integer = 0 To 21
            Dim tbxn1 As String = "Cf_TextBox" & lp1.ToString

            Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
            If cs.Length > 0 Then
                dtx1 &= CType(cs(0), TextBox).Text
                dtx1 &= vbCrLf
            End If
        Next


        Dim stCurrentDir As String = System.IO.Directory.GetCurrentDirectory()
        CuDr = stCurrentDir

        Dim excsv1 As IO.StreamWriter
        excsv1 = New IO.StreamWriter(CuDr & "\config.ini", False, System.Text.Encoding.GetEncoding("shift_jis"))
        excsv1.Write(dtx1)
        excsv1.Close()
        excsv1.Dispose()

    End Sub

    Function dt2unepocht(ByVal vbdate As DateTime) As Long
        ''VBで使用出来る日付を入力すると、UNIX エポック秒に変換する
        vbdate = vbdate.ToUniversalTime()

        Dim dt1 As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        Dim elapsedTime As TimeSpan = vbdate - dt1

        Return CType(elapsedTime.TotalSeconds, Long)

    End Function

    Function imgsamp(ByVal ur As String)

        ''新RMSの[商品画像URL]から、先頭の画像URLを抽出する自作関数

        Dim rl01 As Integer

        If IsDBNull(ur) = True Then
            Return ""
        Else
            rl01 = ur.IndexOf(" ")

            If rl01 < 1 Then
                Return ur
            Else
                Return ur.Substring(0, rl01)
            End If
        End If
    End Function

    Function pnameshp(ByVal pname As String)
        '商品名の整形をする
        Dim rl01 As Integer

        If IsDBNull(pname) = True Then
            Return ""
        Else
            rl01 = pname.IndexOf("|")

            If rl01 < 1 Then

                Dim rl02 As Integer = pname.IndexOf("｜")

                If rl02 < 1 Then
                    Return pname.Replace("代金引換不可", "").Replace("【】", "")
                Else
                    Return pname.Substring(0, rl02).Replace("代金引換不可", "").Replace("【】", "")
                End If

            Else
                Return pname.Substring(0, rl01).Replace("代金引換不可", "").Replace("【】", "")
            End If
        End If
    End Function

    Function idc01(ByVal vo As String)

        ''入力されたデータを小文字に変換する変数
        idc01 = vo.ToLower()

    End Function

    Function idc02(ByVal vo As String)

        ''入力されたデータを大文字に変換する変数
        idc02 = vo.ToUpper()

    End Function


    Function create_temporary_table(ByVal tmptbn As String, ByVal t_name As String) As String
        '作業用一時テーブルを作成(データ名称＋100フィールド)

        Dim sql1 As String = "CREATE TABLE "
        sql1 &= tmptbn
        sql1 &= " ("
        sql1 &= "serial INTEGER UNSIGNED auto_increment primary key"
        sql1 &= ",`"
        sql1 &= t_name
        sql1 &= "` VARCHAR(120)"

        For lpc1 As Integer = 1 To 100 Step 1

            sql1 &= ",`pid"
            sql1 &= lpc1.ToString("000")
            sql1 &= "` VARCHAR(25)"

        Next

        sql1 &= ");"

        Dim rs As String = sql_result_no(sql1)

        If rs = "Complete" Then
            Return "Complete"
        Else
            Return sql1
        End If

    End Function

    Function create_temporary_table02(ByVal tmptbn As String, ByVal t_name As String) As String
        '作業用一時テーブルを作成(データ名称＋100フィールド)

        Dim sql1 As String = "CREATE TABLE "
        sql1 &= tmptbn
        sql1 &= " ("
        sql1 &= "serial INTEGER UNSIGNED auto_increment primary key"
        sql1 &= ",`"
        sql1 &= t_name
        sql1 &= "` VARCHAR(120)"

        For lpc1 As Integer = 1 To 8 Step 1

            sql1 &= ",`pid"
            sql1 &= lpc1.ToString("000")
            sql1 &= "` VARCHAR(25)"
            sql1 &= ",`商品名"
            sql1 &= lpc1.ToString("000")
            sql1 &= "` VARCHAR(120)"
            sql1 &= ",`コメント"
            sql1 &= lpc1.ToString("000")
            sql1 &= "` MEDIUMTEXT"

        Next

        sql1 &= ");"

        Dim rs As String = sql_result_no(sql1)

        If rs = "Complete" Then
            Return "Complete"
        Else
            Return sql1
        End If

    End Function

    Function create_temporary_table03(ByVal tmptbn As String) As String
        '作業用一時テーブルを作成(データ名称＋100フィールド)

        Dim sql1 As String = "CREATE TABLE "
        sql1 &= tmptbn
        sql1 &= " ("
        sql1 &= " `serial` INTEGER UNSIGNED auto_increment primary key"
        sql1 &= ",`pid01` VARCHAR(25)"
        sql1 &= ",`group01` VARCHAR(50)"
        sql1 &= ",`gid01` SMALLINT UNSIGNED"

        sql1 &= ");"

        Dim rs As String = sql_result_no(sql1)

        If rs = "Complete" Then
            Return "Complete"
        Else
            Return sql1
        End If

    End Function


    Function temporarytable_selecti(ByVal tmptbn As String, ByVal t_name As String) As DataTable

        Dim sql1 As String = "SELECT"
        sql1 &= " `serial`"

        For lpc1 As Integer = 1 To 100 Step 1

            sql1 &= ",`pid"
            sql1 &= lpc1.ToString("000")
            sql1 &= "`"
        Next

        sql1 &= ",`"
        sql1 &= t_name
        sql1 &= "`"
        sql1 &= " FROM `"
        sql1 &= tmptbn
        sql1 &= "`;"

        Return sql_result_return(sql1)

    End Function

    Function temporarytable_selecti02(ByVal tmptbn As String, ByVal t_name As String) As DataTable

        Dim sql1 As String = "SELECT"
        sql1 &= " `serial`"

        For lpc1 As Integer = 1 To 8 Step 1

            sql1 &= ",`pid"
            sql1 &= lpc1.ToString("000")
            sql1 &= "`"

            sql1 &= ",`商品名"
            sql1 &= lpc1.ToString("000")
            sql1 &= "`"

            sql1 &= ",`コメント"
            sql1 &= lpc1.ToString("000")
            sql1 &= "`"


        Next

        sql1 &= ",`"
        sql1 &= t_name
        sql1 &= "`"
        sql1 &= " FROM `"
        sql1 &= tmptbn
        sql1 &= "`;"

        Return sql_result_return(sql1)

    End Function

    Function html_template01(ByVal t_name As String) As String

        Dim html1 As String = "<!DOCTYPE html>" & vbCrLf
        html1 &= "<html lang=""ja"">" & vbCrLf
        html1 &= vbCrLf
        html1 &= "	<head>" & vbCrLf
        html1 &= "		<meta charset=""UTF-8"">" & vbCrLf
        html1 &= "		<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & vbCrLf
        html1 &= "		<meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0"">" & vbCrLf
        html1 &= "<title>"
        html1 &= t_name
        html1 &= "誘導リンク</title>" & vbCrLf
        html1 &= "		<link rel=""stylesheet"" href=""css/slick.css"">" & vbCrLf
        html1 &= "		<link rel=""stylesheet"" href=""css/royal.css"">" & vbCrLf
        html1 &= "		<script src=""js/jquery-1.9.1.min.js""></script>" & vbCrLf
        html1 &= "		<script src=""js/slick.min.js""></script>" & vbCrLf
        html1 &= "		<script src=""js/jquery.js""></script>" & vbCrLf
        html1 &= "	</head>" & vbCrLf
        html1 &= vbCrLf
        html1 &= "	<body>" & vbCrLf
        html1 &= "		<div id=""container"">" & vbCrLf
        html1 &= "			<div class=""slider-title"">" & vbCrLf
        html1 &= "				<h3>"
        html1 &= t_name
        html1 &= "</h3>" & vbCrLf
        html1 &= "			</div>" & vbCrLf
        html1 &= "			<div class=""slider""> " & vbCrLf
        html1 &= "				<ul class=""slick"">" & vbCrLf

        Return html1

    End Function

    Function html_template02(ByVal t_name As String) As String

        ''生活空間専用ヘッダ

        Dim html1 As String = "<!DOCTYPE html>" & vbCrLf
        html1 &= "<html lang=""ja"">" & vbCrLf
        html1 &= vbCrLf
        html1 &= "	<head>" & vbCrLf
        html1 &= "		<meta charset=""UTF-8"">" & vbCrLf
        html1 &= "		<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & vbCrLf
        html1 &= "		<meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0"">" & vbCrLf
        html1 &= "<title>"
        html1 &= t_name
        html1 &= "誘導リンク</title>" & vbCrLf
        html1 &= "		<link rel=""stylesheet"" href=""css/slick.css"">" & vbCrLf
        html1 &= "		<link rel=""stylesheet"" href=""css/seikatsu.css"">" & vbCrLf
        html1 &= "		<script src=""js/jquery-1.9.1.min.js""></script>" & vbCrLf
        html1 &= "		<script src=""js/slick.min.js""></script>" & vbCrLf
        html1 &= "		<script src=""js/jquery.js""></script>" & vbCrLf
        html1 &= "	</head>" & vbCrLf
        html1 &= vbCrLf
        html1 &= "	<body>" & vbCrLf
        html1 &= "		<div id=""container"">" & vbCrLf
        html1 &= "			<div class=""slider-title"">" & vbCrLf
        html1 &= "				<h3>"
        html1 &= t_name
        html1 &= "</h3>" & vbCrLf
        html1 &= "			</div>" & vbCrLf
        html1 &= "			<div class=""slider""> " & vbCrLf
        html1 &= "				<ul class=""slick"">" & vbCrLf

        Return html1

    End Function

    Function html_template02(ByVal shop As String, ByVal lpc As Integer, ByVal url As String, ByVal imge As String, ByVal name As String) As String

        'レコメンドHTMLの上部(全店舗共通。URLブロックは店舗ごと)


        Dim html1 As String = "					<!-- " & lpc & "商品 -->" & vbCrLf
        html1 &= "					<li>" & vbCrLf

        Select Case shop
            Case "rr"
                'ロイヤル楽天用
                html1 &= "						<a href=""https://item.rakuten.co.jp/royal3000/"
                html1 &= url
                html1 &= "/"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src="""
                html1 &= imgsamp(imge)
                html1 &= """>" & vbCrLf

            Case "sr"
                '生活空間楽天用
                html1 &= "						<a href=""https://item.rakuten.co.jp/seikatsukukan/"
                html1 &= url
                html1 &= "/"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src="""
                html1 &= imgsamp(imge)
                html1 &= """>" & vbCrLf

            Case "ry"
                ''ロイヤルYahoo用
                html1 &= "						<a href=""https://store.shopping.yahoo.co.jp/royal3000/"
                html1 &= idc01(url)
                html1 &= ".html"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src=""https://item-shopping.c.yimg.jp/i/j/royal3000_"
                html1 &= idc01(url)
                html1 &= """>" & vbCrLf

            Case "sy"
                ''生活空間Yahoo用
                html1 &= "						<a href=""https://store.shopping.yahoo.co.jp/seikatsukukan/"
                html1 &= idc01(url)
                html1 &= ".html"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src=""https://item-shopping.c.yimg.jp/i/j/seikatsukukan_"
                html1 &= idc01(url)
                html1 &= """>" & vbCrLf

            Case "my"
                ''生活空間Yahoo用
                html1 &= "						<a href=""https://store.shopping.yahoo.co.jp/motherplusstore/"
                html1 &= idc01(url)
                html1 &= ".html"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src=""https://item-shopping.c.yimg.jp/i/j/motherplusstore_"
                html1 &= idc01(url)
                html1 &= """>" & vbCrLf


            Case Else
        End Select

        html1 &= "							<p>"

        If IsDBNull(name) = True Then
        Else
            html1 &= name

        End If

        html1 &= "</p>" & vbCrLf
        html1 &= "						</a>" & vbCrLf
        html1 &= "					</li>" & vbCrLf


        Return html1

    End Function

    Function html_template02b(ByVal shop As String, ByVal lpc As Integer, ByVal url As String, ByVal imge As String, ByVal name As String, ByVal com As String) As String

        'レコメンドHTMLの上部(全店舗共通。URLブロックは店舗ごと)


        Dim html1 As String = "					<!-- " & lpc & "商品 -->" & vbCrLf
        html1 &= "					<li>" & vbCrLf

        Select Case shop
            Case "rr"
                'ロイヤル楽天用
                html1 &= "						<a href=""https://item.rakuten.co.jp/royal3000/"
                html1 &= url
                html1 &= "/"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src="""
                html1 &= imgsamp(imge)
                html1 &= """>" & vbCrLf

            Case "sr"
                '生活空間用
                html1 &= "						<a href=""https://item.rakuten.co.jp/seikatsukukan/"
                html1 &= url
                html1 &= "/"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src="""
                html1 &= imgsamp(imge)
                html1 &= """>" & vbCrLf

            Case "sy"
                ''生活空間Yahoo用
                html1 &= "						<a href=""https://store.shopping.yahoo.co.jp/seikatsukukan/"
                html1 &= idc01(url)
                html1 &= ".html"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src=""https://item-shopping.c.yimg.jp/i/j/seikatsukukan_"
                html1 &= idc01(url)
                html1 &= """>" & vbCrLf

            Case "ry"
                ''ロイヤルYahoo用
                html1 &= "						<a href=""https://store.shopping.yahoo.co.jp/royal3000/"
                html1 &= idc01(url)
                html1 &= ".html"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src=""https://item-shopping.c.yimg.jp/i/j/royal3000_"
                html1 &= idc01(url)
                html1 &= """>" & vbCrLf

            Case "my"
                ''生活空間Yahoo用
                html1 &= "						<a href=""https://store.shopping.yahoo.co.jp/motherplusstore/"
                html1 &= idc01(url)
                html1 &= ".html"" target=""_parent"">" & vbCrLf
                html1 &= "							<img src=""https://item-shopping.c.yimg.jp/i/j/motherplusstore_"
                html1 &= idc01(url)
                html1 &= """>" & vbCrLf

            Case Else
        End Select
        html1 &= "							<div Class=""product-desc"">" & vbCrLf
        html1 &= "							  <div Class=""back_box""></div>" & vbCrLf
        html1 &= "							  <div Class=""text_box"">" & vbCrLf
        html1 &= "							    <p>"

        If IsDBNull(name) = True Then
        Else
            html1 &= name

        End If

        html1 &= "</p>" & vbCrLf
        html1 &= "							    <p>"

        If IsDBNull(com) = True Then
        Else
            html1 &= com

        End If

        html1 &= "</p>" & vbCrLf

        html1 &= "							  </div>" & vbCrLf
        html1 &= "							</div>" & vbCrLf

        html1 &= "						</a>" & vbCrLf
        html1 &= "					</li>" & vbCrLf


        Return html1

    End Function

    Function html_template03() As String

        'レコメンドHTMLの上部(全店舗共通)

        Dim html1 As String = "				</ul>" & vbCrLf
        html1 &= "			</div>" & vbCrLf
        html1 &= "			<!-- .slider End -->" & vbCrLf
        html1 &= "		</div>" & vbCrLf
        html1 &= "		<!-- #container End -->" & vbCrLf
        html1 &= "	</body>" & vbCrLf
        html1 &= "" & vbCrLf
        html1 &= "</html>" & vbCrLf


        Return html1

    End Function

    Function r_pc_csv_template01(ByVal shop As String, ByVal t_name As String, ByVal url As String, ByVal pid As String, ByVal pro_des As String, ByVal pro_des_sp As String, ByVal htmln As Integer) As String

        'csv
        Dim des0 As String = ""
        Dim des1 As String = ""

        Dim flnam As String = ""

        Dim epls As Integer = 0

        Select Case t_name
            Case "類似関連"
                flnam = "iframe/similar01/"
                epls = 13

            Case "サイズ関連"
                flnam = "iframe/sizelist01/"
                epls = 14

            Case "セット関連"
                flnam = "iframe/setitem01/"
                epls = 14

            Case Else

        End Select


        Dim csv1 As String = """u"","""
        csv1 &= url
        csv1 &= ""","""
        csv1 &= pid
        csv1 &= ""","""

        If IsDBNull(pro_des) Then
            des0 = ""
        Else
            des0 = pro_des
        End If

        Dim idxst As String = "<!--" & t_name & "開始-->"
        Dim idxsp As String = "<!--" & t_name & "終了-->"

        Dim po0 As Integer = des0.IndexOf(idxst)
        If 0 <= po0 Then
            '含まれています
            Dim po1 As Integer = des0.IndexOf(idxsp) + epls
            des1 = des0.Replace(des0.Substring(po0, po1 - po0), "[RECO]")

        Else
            '含まれていません
            des1 = des0 & "[RECO]"
        End If

        des1 = des1.Replace("""", """""")

        Dim cont1 As String = idxst

        Select Case shop
            Case "rr"
                cont1 &= "<IFRAME src=""""https://www.rakuten.ne.jp/gold/royal3000/"

            Case "sr"
                cont1 &= "<IFRAME src=""""https://www.rakuten.ne.jp/gold/seikatsukukan/"

            Case Else

        End Select

        cont1 &= flnam

        Dim fn2 As Integer = htmln
        cont1 &= fn2.ToString("0000") & ".html"

        cont1 &= """"" frameborder=""""0"""" scrolling=""""no"""" width=""""1000"""" height="""""
        cont1 &= "380"
        cont1 &= """""></IFRAME>"
        cont1 &= idxsp

        des1 = des1.Replace("[RECO]", cont1)
        csv1 &= des1
        csv1 &= """"
        csv1 &= ","
        csv1 &= """"


        ''スマホ用
        Dim des2 As String = ""
        Dim des3 As String = ""

        If IsDBNull(pro_des_sp) Then
            des2 = ""
        Else
            des2 = pro_des_sp
        End If

        Dim po2 As Integer = des2.IndexOf(idxst)
        If 0 <= po2 Then
            '含まれています
            Dim po3 As Integer = des2.IndexOf(idxsp) + epls
            des3 = des2.Replace(des2.Substring(po2, po3 - po2), "[RECO]")

        Else
            '含まれていません
            des3 = des2 & "[RECO]"
        End If

        des3 = des3.Replace("""", """""")

        Dim cont2 As String = idxst
        Select Case shop
            Case "rr"
                cont2 &= "<IFRAME ="""""""" src=""""https://www.rakuten.ne.jp/gold/royal3000/"

            Case "sr"
                cont2 &= "<IFRAME ="""""""" src=""""https://www.rakuten.ne.jp/gold/seikatsukukan/"

            Case Else

        End Select
        cont2 &= flnam

        Dim fn3 As Integer = htmln
        cont2 &= fn2.ToString("0000") & ".html"

        cont2 &= """"" frameborder=""""0"""" scrolling=""""no"""" width=""""100%"""" height=""""360"""" style=""""min-width: 100%; width: 100px;""""></IFRAME ="""""""">"
        cont2 &= idxsp

        des3 = des3.Replace("[RECO]", cont2)
        csv1 &= des3
        csv1 &= """"
        csv1 &= vbCrLf

        Select Case shop
            Case "rr"
                r_reco_link_update("nRms_item_newest", url, des1, des3)

            Case "sr"
                r_reco_link_update("nRms_seikatsukukan_item_newest", url, des1, des3)

            Case Else
        End Select

        Return csv1


    End Function

    Function r_pc_csv_template02(ByVal shop As String, ByVal url As String, ByVal pid As String, ByVal pro_des As String, ByVal pro_des_sp As String, ByVal htmln As Integer) As String

        'csv
        Dim des0 As String = ""
        Dim des1 As String = ""

        Dim flnam As String = "iframe/recommend/"

        Dim epls As Integer = 15

        Dim csv1 As String = """u"","""
        csv1 &= url
        csv1 &= ""","""
        csv1 &= pid
        csv1 &= ""","""

        If IsDBNull(pro_des) Then
            des0 = ""
        Else
            des0 = pro_des
        End If

        Dim idxst As String = "<!--おすすめ商品開始-->"
        Dim idxsp As String = "<!--おすすめ商品終了-->"

        Dim po0 As Integer = des0.IndexOf(idxst)
        If 0 <= po0 Then
            '含まれています
            Dim po1 As Integer = des0.IndexOf(idxsp) + epls
            des1 = des0.Replace(des0.Substring(po0, po1 - po0), "[RECO]")

        Else
            '含まれていません
            des1 = des0 & "[RECO]"
        End If

        des1 = des1.Replace("""", """""")

        Dim cont1 As String = idxst

        Select Case shop
            Case "rr"
                cont1 &= "<IFRAME src=""""https://www.rakuten.ne.jp/gold/royal3000/"

            Case "sr"
                cont1 &= "<IFRAME src=""""https://www.rakuten.ne.jp/gold/seikatsukukan/"

            Case Else

        End Select

        cont1 &= flnam

        Dim fn2 As Integer = htmln
        cont1 &= fn2.ToString("0000") & ".html"

        cont1 &= """"" frameborder=""""0"""" scrolling=""""no"""" width=""""1000"""" height="""""
        cont1 &= "340"
        cont1 &= """""></IFRAME>"
        cont1 &= idxsp

        des1 = des1.Replace("[RECO]", cont1)
        csv1 &= des1
        csv1 &= """"
        csv1 &= ","
        csv1 &= """"


        ''スマホ用
        Dim des2 As String = ""
        Dim des3 As String = ""

        If IsDBNull(pro_des_sp) Then
            des2 = ""
        Else
            des2 = pro_des_sp
        End If

        Dim po2 As Integer = des2.IndexOf(idxst)
        If 0 <= po2 Then
            '含まれています
            Dim po3 As Integer = des2.IndexOf(idxsp) + epls
            des3 = des2.Replace(des2.Substring(po2, po3 - po2), "[RECO]")

        Else
            '含まれていません
            des3 = des2 & "[RECO]"
        End If

        des3 = des3.Replace("""", """""")

        Dim cont2 As String = idxst
        Select Case shop
            Case "rr"
                cont2 &= "<IFRAME ="""""""" src=""""https://www.rakuten.ne.jp/gold/royal3000/"

            Case "sr"
                cont2 &= "<IFRAME ="""""""" src=""""https://www.rakuten.ne.jp/gold/seikatsukukan/"

            Case Else

        End Select
        cont2 &= flnam

        Dim fn3 As Integer = htmln
        cont2 &= fn2.ToString("0000") & ".html"

        cont2 &= """"" frameborder=""""0"""" scrolling=""""no"""" width=""""100%"""" height=""""370"""" style=""""min-width: 100%; width: 100px;""""></IFRAME ="""""""">"
        cont2 &= idxsp

        des3 = des3.Replace("[RECO]", cont2)
        csv1 &= des3
        csv1 &= """"
        csv1 &= vbCrLf

        Select Case shop
            Case "rr"
                r_reco_link_update("nRms_item_newest", url, des1, des3)

            Case "sr"
                r_reco_link_update("nRms_seikatsukukan_item_newest", url, des1, des3)

            Case Else
        End Select

        Return csv1

    End Function


    Function y_pc_csv_template01(ByVal shop As String, ByVal t_name As String, ByVal url As String, ByVal pro_des As String, ByVal htmln As Integer) As String

        'csv
        Dim des0 As String = ""
        Dim des1 As String = ""

        Dim flnam As String = ""

        Dim epls As Integer = 0

        Select Case t_name
            Case "類似関連"
                flnam = "iframe/similar01/"
                epls = 13

            Case "サイズ関連"
                flnam = "iframe/sizelist01/"
                epls = 14

            Case "セット関連"
                flnam = "iframe/setitem01/"
                epls = 14

            Case Else

        End Select


        Dim csv1 As String = """"
        csv1 &= url
        csv1 &= ""","""

        If IsDBNull(pro_des) Then
            des0 = ""
        Else
            des0 = pro_des
        End If

        Dim idxst1 As String = "<!--" & t_name & "開始-->"
        Dim idxsp1 As String = "<!--" & t_name & "終了-->"
        Dim idxst2 As String = "<!--おすすめ商品開始-->"
        Dim idxsp2 As String = "<!--おすすめ商品終了-->"


        Dim po0 As Integer = des0.IndexOf(idxst1)
        If 0 <= po0 Then
            '含まれています
            Dim po1 As Integer = des0.IndexOf(idxsp1) + epls
            des1 = des0.Replace(des0.Substring(po0, po1 - po0), "[RECO]")

        Else
            '含まれていません
            Dim po4 As Integer = des0.IndexOf(idxst2)
            If 0 <= po4 Then
                '含まれています
                des1 = des0.Substring(0, po4)
                des1 &= "[RECO]"
                des1 &= des0.Substring(po4)

            Else
                '含まれていません
                des1 = des0 & "[RECO]"
            End If

        End If

        des1 = des1.Replace("""", """""")

        Dim cont1 As String = idxst1

        Select Case shop
            Case "ry"
                'ロイヤルYahoo!用
                cont1 &= "<IFRAME src=""""https://shopping.geocities.jp/royal3000/"

            Case "sy"
                '生活空間Yahoo!用
                cont1 &= "<IFRAME src=""""https://shopping.geocities.jp/seikatsukukan/"

            Case "my"
                'マザープラスYahoo!用
                cont1 &= "<IFRAME src=""""https://shopping.geocities.jp/motherplusstore/"

            Case Else

        End Select

        cont1 &= flnam

        Dim fn2 As Integer = htmln
        cont1 &= fn2.ToString("0000") & ".html"
        cont1 &= """"""
        cont1 &= " frameborder=""""0"""""
        cont1 &= " scrolling=""""no"""""


        Select Case shop
            Case "ry"
                'ロイヤルYahoo!用
                cont1 &= " width = """"1000"""""
                cont1 &= " height =""""380"""""

            Case "sy"
                '生活空間Yahoo!用
                cont1 &= " width = """"1000"""""
                cont1 &= " height =""""380"""""

            Case "my"
                'マザープラスYahoo!用
                cont1 &= " width = """"740"""""
                cont1 &= " height =""""380"""""

            Case Else

        End Select

        cont1 &= "></IFRAME>"
        cont1 &= idxsp1

        des1 = des1.Replace("[RECO]", cont1)
        csv1 &= des1
        csv1 &= """"
        csv1 &= vbCrLf

        Select Case shop
            Case "ry"
                'ロイヤルYahoo!用
                y_reco_link_update("shopping_yahoo_data_newest", url, des1)

            Case "sy"
                '生活空間Yahoo!用
                y_reco_link_update("seikatsukukan_data_newest", url, des1)

            Case "my"
                'マザープラスYahoo!用
                y_reco_link_update("motherplusstore_data_newest", url, des1)
            Case Else
        End Select

        Return csv1


    End Function

    Function y_pc_csv_template02(ByVal shop As String, ByVal url As String, ByVal pro_des As String, ByVal htmln As Integer) As String

        'csv
        Dim des0 As String = ""
        Dim des1 As String = ""

        Dim flnam As String = "iframe/recommend/"

        Dim epls As Integer = 15

        Dim csv1 As String = """"
        csv1 &= url
        csv1 &= ""","""

        If IsDBNull(pro_des) Then
            des0 = ""
        Else
            des0 = pro_des
        End If

        Dim idxst As String = "<!--おすすめ商品開始-->"
        Dim idxsp As String = "<!--おすすめ商品終了-->"

        Dim po0 As Integer = des0.IndexOf(idxst)
        If 0 <= po0 Then
            '含まれています
            Dim po1 As Integer = des0.IndexOf(idxsp) + epls
            des1 = des0.Replace(des0.Substring(po0, po1 - po0), "[RECO]")

        Else
            '含まれていません
            des1 = des0 & "[RECO]"

        End If

        des1 = des1.Replace("""", """""")

        Dim cont1 As String = idxst

        Select Case shop
            Case "ry"
                'ロイヤルYahoo!用
                cont1 &= "<IFRAME src=""""https://shopping.geocities.jp/royal3000/"

            Case "sy"
                '生活空間Yahoo!用
                cont1 &= "<IFRAME src=""""https://shopping.geocities.jp/seikatsukukan/"

            Case "my"
                'マザープラスYahoo!用
                cont1 &= "<IFRAME src=""""https://shopping.geocities.jp/motherplusstore/"

            Case Else

        End Select

        cont1 &= flnam

        Dim fn2 As Integer = htmln
        cont1 &= fn2.ToString("0000") & ".html"
        cont1 &= """"""
        cont1 &= " frameborder=""""0"""""
        cont1 &= " scrolling=""""no"""""


        Select Case shop
            Case "ry"
                'ロイヤルYahoo!用
                cont1 &= " width = """"1000"""""
                cont1 &= " height =""""340"""""

            Case "sy"
                '生活空間Yahoo!用
                cont1 &= " width = """"1000"""""
                cont1 &= " height =""""340"""""

            Case "my"
                'マザープラスYahoo!用
                cont1 &= " width = """"740"""""
                cont1 &= " height =""""340"""""

            Case Else

        End Select

        cont1 &= "></IFRAME>"
        cont1 &= idxsp

        des1 = des1.Replace("[RECO]", cont1)
        csv1 &= des1
        csv1 &= """"
        csv1 &= vbCrLf


        Select Case shop
            Case "ry"
                'ロイヤルYahoo!用
                y_reco_link_update("shopping_yahoo_data_newest", url, des1)

            Case "sy"
                '生活空間Yahoo!用
                y_reco_link_update("seikatsukukan_data_newest", url, des1)

            Case "my"
                'マザープラスYahoo!用
                y_reco_link_update("motherplusstore_data_newest", url, des1)
            Case Else
        End Select

        Return csv1


    End Function

    Function r_reco_link_update(ByVal tmptbn As String, ByVal url As String, ByVal pro_des As String, ByVal pro_des_sp As String) As String

        Dim sql1 As String = "UPDATE `"
        sql1 &= tmptbn
        sql1 &= "` Set `PC用商品説明文` = "
        sql1 &= "'"
        sql1 &= pro_des.Replace("""""", """")
        sql1 &= "'"

        sql1 &= ","

        sql1 &= " `スマートフォン用商品説明文` = "
        sql1 &= "'"
        sql1 &= pro_des_sp.Replace("""""", """")
        sql1 &= "'"

        sql1 &= " WHERE `商品管理番号（商品URL）` = '"
        sql1 &= url
        sql1 &= "';"

        Dim rs As String = sql_result_no(sql1)

        If rs = "Complete" Then
            Return "Complete"
        Else
            Return sql1
        End If

    End Function

    Function y_reco_link_update(ByVal tmptbn As String, ByVal url As String, ByVal pro_des As String) As String

        Dim sql1 As String = "UPDATE `"
        sql1 &= tmptbn
        sql1 &= "` Set `caption` = "
        sql1 &= "'"
        sql1 &= pro_des.Replace("""""", """")
        sql1 &= "'"

        sql1 &= " WHERE `code` = '"
        sql1 &= url
        sql1 &= "';"

        Dim rs As String = sql_result_no(sql1)

        If rs = "Complete" Then
            Return "Complete"
        Else
            Return sql1
        End If

    End Function

    Sub excelRead01()
        '楽天(royal)の類似関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "類似関連")
        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("類似関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `類似関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100

                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If

                    End If

                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "類似関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/similar01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）`"
                        sql3 &= ",`nRms_item_newest`.`商品番号`"
                        sql3 &= ",`nRms_item_newest`.`商品名`"
                        sql3 &= ",`nRms_item_newest`.`商品画像URL`"
                        sql3 &= ",`nRms_item_newest`.`PC用商品説明文`"
                        sql3 &= ",`nRms_item_newest`.`スマートフォン用商品説明文`"
                        sql3 &= " FROM `nRms_item_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `nRms_item_newest`.`倉庫指定` = 0"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("rr", lpc1, DRow3.Item(0), DRow3.Item(3), pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(4)) = False Then
                                    pro_des1 = DRow3.Item(4)
                                End If

                                csv1 &= r_pc_csv_template01("rr", "類似関連", DRow3.Item(0), DRow3.Item(1), pro_des1, DRow3.Item(5), DRow.Item(0))


                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox1.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox2.Text, Form1.Cf_TextBox3.Text)

                'CSVの出力(テスト)

            Next
            'CSVの出力(本番)
            Dim lofn2 As String = CuDr & "\item.csv"

            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead11()

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "類似関連")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("類似関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `類似関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100

                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If

                    End If

                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "類似関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/similar01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品番号`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品名`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品画像URL`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`PC用商品説明文`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`スマートフォン用商品説明文`"
                        sql3 &= " FROM `nRms_seikatsukukan_item_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `nRms_seikatsukukan_item_newest`.`倉庫指定` = 0"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("sr", lpc1, DRow3.Item(0), DRow3.Item(3), pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(4)) = False Then
                                    pro_des1 = DRow3.Item(4)
                                End If

                                Dim pro_des2 As String = ""
                                If IsDBNull(DRow3.Item(5)) = False Then
                                    pro_des2 = DRow3.Item(5)
                                End If

                                csv1 &= r_pc_csv_template01("sr", "類似関連", DRow3.Item(0), DRow3.Item(1), pro_des1, pro_des2, DRow.Item(0))

                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()


                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox10.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox11.Text, Form1.Cf_TextBox12.Text)

            Next

            'CSVの出力テスト
            Dim lofn2 As String = CuDr & "\item.csv"

            '書き込むファイルが既に存在している場合は、上書きする
            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If

        'CSVの出力ここへ戻す


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead21()
        'yahoo!(royal)の類似関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "類似関連")
        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("類似関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `類似関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100

                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If

                    End If

                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "類似関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/similar01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `shopping_yahoo_data_newest`.`code`"
                        sql3 &= ",`shopping_yahoo_data_newest`.`caption`"
                        sql3 &= ",`shopping_yahoo_data_newest`.`name`"
                        sql3 &= " FROM `shopping_yahoo_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `shopping_yahoo_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `shopping_yahoo_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("ry", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("ry", "類似関連", DRow3.Item(0), pro_des1, DRow.Item(0))


                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox4.Text & "/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox5.Text, Form1.Cf_TextBox6.Text)

                'CSVの出力(テスト)

            Next
            'CSVの出力(本番)
            Dim lofn2 As String = CuDr & "\data_spy.csv"

            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead02()
        '楽天(royal)のサイズレコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table(tmptbn02, "サイズ関連")


        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("サイズ関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else
                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `サイズ関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "サイズ関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/sizelist01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）`"
                        sql3 &= ",`nRms_item_newest`.`商品番号`"
                        sql3 &= ",`nRms_item_newest`.`商品名`"
                        sql3 &= ",`nRms_item_newest`.`商品画像URL`"
                        sql3 &= ",`nRms_item_newest`.`PC用商品説明文`"
                        sql3 &= ",`nRms_item_newest`.`スマートフォン用商品説明文`"
                        sql3 &= " FROM `nRms_item_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `nRms_item_newest`.`倉庫指定` = 0"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html

                                html1 &= html_template02("rr", lpc1, DRow3.Item(0), DRow3.Item(3), pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(4)) = False Then
                                    pro_des1 = DRow3.Item(4)
                                End If

                                Dim pro_des2 As String = ""
                                If IsDBNull(DRow3.Item(5)) = False Then
                                    pro_des2 = DRow3.Item(5)
                                End If

                                csv1 &= r_pc_csv_template01("rr", "サイズ関連", DRow3.Item(0), DRow3.Item(1), pro_des1, DRow3.Item(5), DRow.Item(0))


                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox1.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox2.Text, Form1.Cf_TextBox3.Text)

                'CSVの出力テスト

            Next

            'CSVの出力ここへ戻す
            Dim lofn2 As String = CuDr & "\item.csv"

            '書き込むファイルが既に存在している場合は、上書きする
            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead12()

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table(tmptbn02, "サイズ関連")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("サイズ関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else
                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `サイズ関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成
        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "サイズ関連")


        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/sizelist01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品番号`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品名`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品画像URL`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`PC用商品説明文`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`スマートフォン用商品説明文`"
                        sql3 &= " FROM `nRms_seikatsukukan_item_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `nRms_seikatsukukan_item_newest`.`倉庫指定` = 0"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("sr", lpc1, DRow3.Item(0), DRow3.Item(3), pnameshp(DRow3.Item(2)))

                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(4)) = False Then
                                    pro_des1 = DRow3.Item(4)
                                End If

                                Dim pro_des2 As String = ""
                                If IsDBNull(DRow3.Item(5)) = False Then
                                    pro_des2 = DRow3.Item(5)
                                End If

                                csv1 &= r_pc_csv_template01("sr", "サイズ関連", DRow3.Item(0), DRow3.Item(1), pro_des1, DRow3.Item(5), DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox10.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox11.Text, Form1.Cf_TextBox12.Text)

                'CSVの出力テスト

            Next

            'CSVの出力ここへ戻す
            Dim lofn2 As String = CuDr & "\item.csv"

            '書き込むファイルが既に存在している場合は、上書きする
            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead22()
        'Yahoo!(royal)のサイズレコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table(tmptbn02, "サイズ関連")


        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("サイズ関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else
                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `サイズ関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "サイズ関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/sizelist01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `shopping_yahoo_data_newest`.`code`"
                        sql3 &= ",`shopping_yahoo_data_newest`.`caption`"
                        sql3 &= ",`shopping_yahoo_data_newest`.`name`"
                        sql3 &= " FROM `shopping_yahoo_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `shopping_yahoo_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `shopping_yahoo_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows

                                'html
                                html1 &= html_template02("ry", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("ry", "サイズ関連", DRow3.Item(0), pro_des1, DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox4.Text & "/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox5.Text, Form1.Cf_TextBox6.Text)

                'CSVの出力テスト

            Next

            'CSVの出力ここへ戻す
            Dim lofn2 As String = CuDr & "\data_spy.csv"

            '書き込むファイルが既に存在している場合は、上書きする
            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub


    Sub excelRead23()
        'Yahoo!(royal)セット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "セット関連")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("セット関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `セット関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"
                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成
        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "セット関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/setitem01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `shopping_yahoo_data_newest`.`code`"
                        sql3 &= ",`shopping_yahoo_data_newest`.`caption`"
                        sql3 &= ",`shopping_yahoo_data_newest`.`name`"
                        sql3 &= " FROM `shopping_yahoo_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `shopping_yahoo_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `shopping_yahoo_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("ry", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("ry", "セット関連", DRow3.Item(0), pro_des1, DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox4.Text & "/" & flnam & fn.ToString("0000") & ".html"

                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox5.Text, Form1.Cf_TextBox6.Text)

            Next

        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\data_spy.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead24()
        'Yahoo!(royal)セット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table02(tmptbn02, "グループ")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("レコメンドグループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `グループ`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"

                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 22 Step 3
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ", "
                                sql2b &= ", "

                                sql2h &= "`pid"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= "'" & nrow1.GetCell(c1).ToString & "'"

                                sql2h &= ",`商品名"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 1).ToString & "'"


                                sql2h &= ",`コメント"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 2).ToString & "'"

                            End If

                            num2 += 1

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        tmptbn03 = "summary" & dt2unepocht(noto)
        rs = create_temporary_table03(tmptbn03)

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book02 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book02.GetSheetIndex("レコメンド商品")
            Dim sheet2 As ISheet = book02.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet2.LastRowNum


            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet2.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("

                        sql2h &= tmptbn03
                        sql2h &= "` ("
                        sql2h &= " `pid01`"
                        sql2h &= ",`group01`"


                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"
                        sql2b &= ",'" & nrow1.GetCell(1).ToString & "'"
                        sql2b &= ");"

                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If


        'HTML作成
        Dim dTb1 As DataTable = temporarytable_selecti02(tmptbn02, "グループ")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "/iframe/recommend/"

                Dim html1 As String = html_template01("スタッフのオススメ商品")

                Dim lpc2 As Integer = 1

                For lpc1 As Integer = 1 To 23 Step 3

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `shopping_yahoo_data_newest`.`code`"
                        sql3 &= " FROM `shopping_yahoo_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `shopping_yahoo_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `shopping_yahoo_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02b("ry", lpc1, DRow3.Item(0), "", DRow.Item(lpc1 + 1), DRow.Item(lpc1 + 2))

                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox4.Text & "/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox5.Text, Form1.Cf_TextBox6.Text)

                'UPDATE
                Dim sql4 As String = "UPDATE `"
                sql4 &= tmptbn03
                sql4 &= "` Set `gid01` = "
                sql4 &= lpc3
                sql4 &= " WHERE `group01` = '"
                sql4 &= DRow.Item(25)
                sql4 &= "';"

                rs = sql_result_no(sql4)
                If rs = "Complete" Then
                Else
                    Debug.Print(rs)
                End If

                lpc3 += 1

            Next

        End If

        'csv作成
        Dim csv1 As String = """code"",""caption""" & vbCrLf

        sql1 = "SELECT"
        sql1 &= " `serial`" '0
        sql1 &= ",`pid01`"  '1
        sql1 &= ",`group01`"    '2
        sql1 &= ",`gid01`"  '3
        sql1 &= " FROM `"
        sql1 &= tmptbn03
        sql1 &= "`"
        dTb1 = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim sql3 As String = "SELECT"
                sql3 &= " `shopping_yahoo_data_newest`.`code`"
                sql3 &= ",`shopping_yahoo_data_newest`.`caption`"
                sql3 &= ",`shopping_yahoo_data_newest`.`name`"
                sql3 &= " FROM `shopping_yahoo_data_newest`"
                sql3 &= " WHERE"
                sql3 &= " `shopping_yahoo_data_newest`.`code` = '"
                sql3 &= DRow.Item(1)
                sql3 &= "' AND `shopping_yahoo_data_newest`.`display` = 1"
                sql3 &= " LIMIT 1;"

                Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                If dTb3.Rows.Count = 0 Then
                    'Debug.Print(DRow.Item(c1))
                Else
                    For Each DRow3 As DataRow In dTb3.Rows

                        'CSV
                        Dim pro_des1 As String = ""
                        If IsDBNull(DRow3.Item(1)) = False Then
                            pro_des1 = DRow3.Item(1)
                        End If

                        csv1 &= y_pc_csv_template02("ry", DRow3.Item(0), pro_des1, DRow.Item(0))

                    Next
                End If

            Next

        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\data_spy.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")


    End Sub


    Sub excelRead31()
        'yahoo!(マザープラス)の類似関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "類似関連")
        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("類似関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `類似関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100

                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If

                    End If

                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "類似関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/similar01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `motherplusstore_data_newest`.`code`"
                        sql3 &= ",`motherplusstore_data_newest`.`caption`"
                        sql3 &= ",`motherplusstore_data_newest`.`name`"
                        sql3 &= " FROM `motherplusstore_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `motherplusstore_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `motherplusstore_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("my", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("my", "類似関連", DRow3.Item(0), pro_des1, DRow.Item(0))


                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox19.Text & "/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox20.Text, Form1.Cf_TextBox21.Text)

                'CSVの出力(テスト)

            Next
            'CSVの出力(本番)
            Dim lofn2 As String = CuDr & "\data_spy.csv"

            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead32()
        ''yahoo!(マザープラス)のサイズレコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table(tmptbn02, "サイズ関連")


        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("サイズ関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else
                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `サイズ関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "サイズ関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/sizelist01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `motherplusstore_data_newest`.`code`"
                        sql3 &= ",`motherplusstore_data_newest`.`caption`"
                        sql3 &= ",`motherplusstore_data_newest`.`name`"
                        sql3 &= " FROM `motherplusstore_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `motherplusstore_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `motherplusstore_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows

                                'html
                                html1 &= html_template02("my", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("my", "サイズ関連", DRow3.Item(0), pro_des1, DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox19.Text & "/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox20.Text, Form1.Cf_TextBox21.Text)

                'CSVの出力テスト

            Next

            'CSVの出力ここへ戻す
            Dim lofn2 As String = CuDr & "\data_spy.csv"

            '書き込むファイルが既に存在している場合は、上書きする
            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead33()
        ''yahoo!(マザープラス)のセット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "セット関連")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("セット関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `セット関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"
                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成
        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "セット関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/setitem01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `motherplusstore_data_newest`.`code`"
                        sql3 &= ",`motherplusstore_data_newest`.`caption`"
                        sql3 &= ",`motherplusstore_data_newest`.`name`"
                        sql3 &= " FROM `motherplusstore_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `motherplusstore_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `motherplusstore_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("my", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("my", "セット関連", DRow3.Item(0), pro_des1, DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox19.Text & "/" & flnam & fn.ToString("0000") & ".html"

                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox20.Text, Form1.Cf_TextBox21.Text)
            Next

        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\data_spy.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead34()
        ''yahoo!(マザープラス)のセット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table02(tmptbn02, "グループ")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("レコメンドグループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `グループ`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"

                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 22 Step 3
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ", "
                                sql2b &= ", "

                                sql2h &= "`pid"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= "'" & nrow1.GetCell(c1).ToString & "'"

                                sql2h &= ",`商品名"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 1).ToString & "'"


                                sql2h &= ",`コメント"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 2).ToString & "'"

                            End If

                            num2 += 1

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        tmptbn03 = "summary" & dt2unepocht(noto)
        rs = create_temporary_table03(tmptbn03)

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book02 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book02.GetSheetIndex("レコメンド商品")
            Dim sheet2 As ISheet = book02.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet2.LastRowNum


            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet2.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("

                        sql2h &= tmptbn03
                        sql2h &= "` ("
                        sql2h &= " `pid01`"
                        sql2h &= ",`group01`"


                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"
                        sql2b &= ",'" & nrow1.GetCell(1).ToString & "'"
                        sql2b &= ");"

                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If


        'HTML作成
        Dim dTb1 As DataTable = temporarytable_selecti02(tmptbn02, "グループ")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "/iframe/recommend/"

                Dim html1 As String = html_template01("スタッフのオススメ商品")

                Dim lpc2 As Integer = 1

                For lpc1 As Integer = 1 To 23 Step 3

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `motherplusstore_data_newest`.`code`"
                        sql3 &= " FROM `motherplusstore_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `motherplusstore_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `motherplusstore_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02b("my", lpc1, DRow3.Item(0), "", DRow.Item(lpc1 + 1), DRow.Item(lpc1 + 2))

                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox19.Text & "/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox20.Text, Form1.Cf_TextBox21.Text)

                'UPDATE
                Dim sql4 As String = "UPDATE `"
                sql4 &= tmptbn03
                sql4 &= "` Set `gid01` = "
                sql4 &= lpc3
                sql4 &= " WHERE `group01` = '"
                sql4 &= DRow.Item(25)
                sql4 &= "';"

                rs = sql_result_no(sql4)
                If rs = "Complete" Then
                Else
                    Debug.Print(rs)
                End If

                lpc3 += 1

            Next

        End If

        'csv作成
        Dim csv1 As String = """code"",""caption""" & vbCrLf

        sql1 = "SELECT"
        sql1 &= " `serial`" '0
        sql1 &= ",`pid01`"  '1
        sql1 &= ",`group01`"    '2
        sql1 &= ",`gid01`"  '3
        sql1 &= " FROM `"
        sql1 &= tmptbn03
        sql1 &= "`"
        dTb1 = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim sql3 As String = "SELECT"
                sql3 &= " `motherplusstore_data_newest`.`code`"
                sql3 &= ",`motherplusstore_data_newest`.`caption`"
                sql3 &= ",`motherplusstore_data_newest`.`name`"
                sql3 &= " FROM `motherplusstore_data_newest`"
                sql3 &= " WHERE"
                sql3 &= " `motherplusstore_data_newest`.`code` = '"
                sql3 &= DRow.Item(1)
                sql3 &= "' AND `motherplusstore_data_newest`.`display` = 1"
                sql3 &= " LIMIT 1;"

                Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                If dTb3.Rows.Count = 0 Then
                    'Debug.Print(DRow.Item(c1))
                Else
                    For Each DRow3 As DataRow In dTb3.Rows

                        'CSV
                        Dim pro_des1 As String = ""
                        If IsDBNull(DRow3.Item(1)) = False Then
                            pro_des1 = DRow3.Item(1)
                        End If

                        csv1 &= y_pc_csv_template02("my", DRow3.Item(0), pro_des1, DRow.Item(0))

                    Next
                End If

            Next

        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\data_spy.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")


    End Sub

    Sub excelRead41()
        'yahoo!(生活空間)の類似関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "類似関連")
        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("類似関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `類似関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100

                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If

                    End If

                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "類似関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/similar01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `seikatsukukan_data_newest`.`code`"
                        sql3 &= ",`seikatsukukan_data_newest`.`caption`"
                        sql3 &= ",`seikatsukukan_data_newest`.`name`"
                        sql3 &= " FROM `seikatsukukan_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `seikatsukukan_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `seikatsukukan_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("sy", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("sy", "類似関連", DRow3.Item(0), pro_des1, DRow.Item(0))


                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox16.Text & "/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox17.Text, Form1.Cf_TextBox18.Text)

                'CSVの出力(テスト)

            Next
            'CSVの出力(本番)
            Dim lofn2 As String = CuDr & "\data_spy.csv"

            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead42()
        'yahoo!(生活空間)サイズレコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table(tmptbn02, "サイズ関連")


        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("サイズ関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else
                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else
                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `サイズ関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"

                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "サイズ関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/sizelist01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `seikatsukukan_data_newest`.`code`"
                        sql3 &= ",`seikatsukukan_data_newest`.`caption`"
                        sql3 &= ",`seikatsukukan_data_newest`.`name`"
                        sql3 &= " FROM `seikatsukukan_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `seikatsukukan_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `seikatsukukan_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows

                                'html
                                html1 &= html_template02("sy", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("sy", "サイズ関連", DRow3.Item(0), pro_des1, DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox16.Text & "/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox17.Text, Form1.Cf_TextBox18.Text)

                'CSVの出力テスト

            Next

            'CSVの出力ここへ戻す
            Dim lofn2 As String = CuDr & "\data_spy.csv"

            '書き込むファイルが既に存在している場合は、上書きする
            Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
            'TextBox1.Textの内容を書き込む
            sw2.Write(csv1)
            '閉じる
            sw2.Close()


        End If


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub


    Sub excelRead43()
        ''yahoo!(生活空間)セット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "セット関連")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("セット関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `セット関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"
                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成
        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "セット関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/setitem01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `seikatsukukan_data_newest`.`code`"
                        sql3 &= ",`seikatsukukan_data_newest`.`caption`"
                        sql3 &= ",`seikatsukukan_data_newest`.`name`"
                        sql3 &= " FROM `seikatsukukan_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `seikatsukukan_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `seikatsukukan_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("sy", lpc1, DRow3.Item(0), "", pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(1)) = False Then
                                    pro_des1 = DRow3.Item(1)
                                End If

                                csv1 &= y_pc_csv_template01("sy", "セット関連", DRow3.Item(0), pro_des1, DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox16.Text & "/" & flnam & fn.ToString("0000") & ".html"

                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox17.Text, Form1.Cf_TextBox18.Text)
            Next

        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\data_spy.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead44()
        ''yahoo!(生活空間)セット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table02(tmptbn02, "グループ")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("レコメンドグループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `グループ`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"

                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 22 Step 3
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ", "
                                sql2b &= ", "

                                sql2h &= "`pid"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= "'" & nrow1.GetCell(c1).ToString & "'"

                                sql2h &= ",`商品名"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 1).ToString & "'"


                                sql2h &= ",`コメント"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 2).ToString & "'"

                            End If

                            num2 += 1

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        tmptbn03 = "summary" & dt2unepocht(noto)
        rs = create_temporary_table03(tmptbn03)

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book02 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book02.GetSheetIndex("レコメンド商品")
            Dim sheet2 As ISheet = book02.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet2.LastRowNum


            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet2.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("

                        sql2h &= tmptbn03
                        sql2h &= "` ("
                        sql2h &= " `pid01`"
                        sql2h &= ",`group01`"


                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"
                        sql2b &= ",'" & nrow1.GetCell(1).ToString & "'"
                        sql2b &= ");"

                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If


        'HTML作成
        Dim dTb1 As DataTable = temporarytable_selecti02(tmptbn02, "グループ")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "/iframe/recommend/"

                Dim html1 As String = html_template02("スタッフのオススメ商品")

                Dim lpc2 As Integer = 1

                For lpc1 As Integer = 1 To 23 Step 3

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `seikatsukukan_data_newest`.`code`"
                        sql3 &= " FROM `seikatsukukan_data_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `seikatsukukan_data_newest`.`code` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `seikatsukukan_data_newest`.`display` = 1"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02b("sy", lpc1, DRow3.Item(0), "", DRow.Item(lpc1 + 1), DRow.Item(lpc1 + 2))

                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox16.Text & "/" & flnam & fn.ToString("0000") & ".html"

                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox17.Text, Form1.Cf_TextBox18.Text)

                'UPDATE
                Dim sql4 As String = "UPDATE `"
                sql4 &= tmptbn03
                sql4 &= "` Set `gid01` = "
                sql4 &= lpc3
                sql4 &= " WHERE `group01` = '"
                sql4 &= DRow.Item(25)
                sql4 &= "';"

                rs = sql_result_no(sql4)
                If rs = "Complete" Then
                Else
                    Debug.Print(rs)
                End If

                lpc3 += 1

            Next

        End If

        'csv作成
        Dim csv1 As String = """code"",""caption""" & vbCrLf

        sql1 = "SELECT"
        sql1 &= " `serial`" '0
        sql1 &= ",`pid01`"  '1
        sql1 &= ",`group01`"    '2
        sql1 &= ",`gid01`"  '3
        sql1 &= " FROM `"
        sql1 &= tmptbn03
        sql1 &= "`"
        dTb1 = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim sql3 As String = "SELECT"
                sql3 &= " `seikatsukukan_data_newest`.`code`"
                sql3 &= ",`seikatsukukan_data_newest`.`caption`"
                sql3 &= ",`seikatsukukan_data_newest`.`name`"
                sql3 &= " FROM `seikatsukukan_data_newest`"
                sql3 &= " WHERE"
                sql3 &= " `seikatsukukan_data_newest`.`code` = '"
                sql3 &= DRow.Item(1)
                sql3 &= "' AND `seikatsukukan_data_newest`.`display` = 1"
                sql3 &= " LIMIT 1;"

                Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                If dTb3.Rows.Count = 0 Then
                    'Debug.Print(DRow.Item(c1))
                Else
                    For Each DRow3 As DataRow In dTb3.Rows

                        'CSV
                        Dim pro_des1 As String = ""
                        If IsDBNull(DRow3.Item(1)) = False Then
                            pro_des1 = DRow3.Item(1)
                        End If

                        csv1 &= y_pc_csv_template02("sy", DRow3.Item(0), pro_des1, DRow.Item(0))

                    Next
                End If

            Next

        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\data_spy.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")


    End Sub


    Sub excelRead03()
        '楽天(royal)のセット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "セット関連")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("セット関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `セット関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"
                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成

        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "セット関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/setitem01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）`"
                        sql3 &= ",`nRms_item_newest`.`商品番号`"
                        sql3 &= ",`nRms_item_newest`.`商品名`"
                        sql3 &= ",`nRms_item_newest`.`商品画像URL`"
                        sql3 &= ",`nRms_item_newest`.`PC用商品説明文`"
                        sql3 &= ",`nRms_item_newest`.`スマートフォン用商品説明文`"
                        sql3 &= " FROM `nRms_item_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `nRms_item_newest`.`倉庫指定` = 0"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02("rr", lpc1, DRow3.Item(0), DRow3.Item(3), pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(4)) = False Then
                                    pro_des1 = DRow3.Item(4)
                                End If

                                Dim pro_des2 As String = ""
                                If IsDBNull(DRow3.Item(5)) = False Then
                                    pro_des2 = DRow3.Item(5)
                                End If


                                csv1 &= r_pc_csv_template01("rr", "セット関連", DRow3.Item(0), DRow3.Item(1), pro_des1, DRow3.Item(5), DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox1.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox2.Text, Form1.Cf_TextBox3.Text)

            Next

        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\item.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        '
        'SizeVar01
        'SetVar01



        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead04()
        '楽天(royal)のおすすめ関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table02(tmptbn02, "グループ")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("レコメンドグループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `グループ`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"

                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 22 Step 3
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ", "
                                sql2b &= ", "

                                sql2h &= "`pid"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= "'" & nrow1.GetCell(c1).ToString & "'"

                                sql2h &= ",`商品名"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 1).ToString & "'"


                                sql2h &= ",`コメント"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 2).ToString & "'"

                            End If

                            num2 += 1

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        tmptbn03 = "summary" & dt2unepocht(noto)
        rs = create_temporary_table03(tmptbn03)

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book02 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book02.GetSheetIndex("レコメンド商品")
            Dim sheet2 As ISheet = book02.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet2.LastRowNum


            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet2.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("

                        sql2h &= tmptbn03
                        sql2h &= "` ("
                        sql2h &= " `pid01`"
                        sql2h &= ",`group01`"


                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"
                        sql2b &= ",'" & nrow1.GetCell(1).ToString & "'"
                        sql2b &= ");"

                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If


        'HTML作成
        Dim dTb1 As DataTable = temporarytable_selecti02(tmptbn02, "グループ")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "/iframe/recommend/"

                Dim html1 As String = html_template01("スタッフのオススメ商品")

                Dim lpc2 As Integer = 1

                For lpc1 As Integer = 1 To 23 Step 3

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）`"
                        sql3 &= ",`nRms_item_newest`.`商品番号`"
                        sql3 &= ",`nRms_item_newest`.`商品名`"
                        sql3 &= ",`nRms_item_newest`.`商品画像URL`"
                        sql3 &= " FROM `nRms_item_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `nRms_item_newest`.`倉庫指定` = 0"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02b("rr", lpc1, DRow3.Item(0), DRow3.Item(3), pnameshp(DRow.Item(lpc1 + 1)), DRow.Item(lpc1 + 2))

                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox1.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox2.Text, Form1.Cf_TextBox3.Text)

                'UPDATE
                Dim sql4 As String = "UPDATE `"
                sql4 &= tmptbn03
                sql4 &= "` Set `gid01` = "
                sql4 &= lpc3
                sql4 &= " WHERE `group01` = '"
                sql4 &= DRow.Item(25)
                sql4 &= "';"

                rs = sql_result_no(sql4)
                If rs = "Complete" Then
                Else
                    Debug.Print(rs)
                End If

                lpc3 += 1

            Next

        End If

        'csv作成
        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        sql1 = "SELECT"
        sql1 &= " `serial`" '0
        sql1 &= ",`pid01`"  '1
        sql1 &= ",`group01`"    '2
        sql1 &= ",`gid01`"  '3
        sql1 &= " FROM `"
        sql1 &= tmptbn03
        sql1 &= "`"
        dTb1 = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim sql3 As String = "SELECT"
                sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）`"
                sql3 &= ",`nRms_item_newest`.`商品番号`"
                sql3 &= ",`nRms_item_newest`.`PC用商品説明文`"
                sql3 &= ",`nRms_item_newest`.`スマートフォン用商品説明文`"
                sql3 &= " FROM `nRms_item_newest`"
                sql3 &= " WHERE"
                sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）` = '"
                sql3 &= DRow.Item(1)
                sql3 &= "' AND `nRms_item_newest`.`倉庫指定` = 0"
                sql3 &= " LIMIT 1;"

                Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                If dTb3.Rows.Count = 0 Then
                    'Debug.Print(DRow.Item(c1))
                Else
                    For Each DRow3 As DataRow In dTb3.Rows
                        'CSV

                        Dim pro_des1 As String = ""
                        If IsDBNull(DRow3.Item(2)) = False Then
                            pro_des1 = DRow3.Item(2)
                        End If

                        Dim pro_des2 As String = ""
                        If IsDBNull(DRow3.Item(3)) = False Then
                            pro_des2 = DRow3.Item(3)
                        End If
                        csv1 &= r_pc_csv_template02("rr", DRow3.Item(0), DRow3.Item(1), pro_des1, pro_des2, DRow.Item(3))

                    Next
                End If

            Next

        End If

        'CSVの出力(koko)
        Dim lofn2 As String = CuDr & "\item.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead13()
        '楽天(生活空間)のセット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        Dim rs As String = create_temporary_table(tmptbn02, "セット関連")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("セット関連グループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `セット関連`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"


                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 100
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て
                                sql2h &= ",`pid"
                                sql2h &= c1.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1).ToString & "'"
                            End If

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        'HTML作成
        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        Dim dTb1 As DataTable = temporarytable_selecti(tmptbn02, "セット関連")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "iframe/setitem01/"

                Dim html1 As String = html_template01(DRow.Item(101))

                Dim lpc2 As Integer = 1
                For lpc1 As Integer = 1 To 100

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品番号`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品名`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品画像URL`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`PC用商品説明文`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`スマートフォン用商品説明文`"
                        sql3 &= " FROM `nRms_seikatsukukan_item_newest` LEFT JOIN `product_ledger`"
                        sql3 &= " ON `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）` = `product_ledger`.`ID`"
                        sql3 &= " WHERE"
                        sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `nRms_seikatsukukan_item_newest`.`倉庫指定` = 0"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html

                                html1 &= html_template02("sr", lpc1, DRow3.Item(0), DRow3.Item(3), pnameshp(DRow3.Item(2)))

                                'CSV
                                Dim pro_des1 As String = ""
                                If IsDBNull(DRow3.Item(4)) = False Then
                                    pro_des1 = DRow3.Item(4)
                                End If

                                Dim pro_des2 As String = ""
                                If IsDBNull(DRow3.Item(5)) = False Then
                                    pro_des2 = DRow3.Item(5)
                                End If


                                csv1 &= r_pc_csv_template01("sr", "セット関連", DRow3.Item(0), DRow3.Item(1), pro_des1, pro_des2, DRow.Item(0))

                            Next
                        End If

                    End If


                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox10.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox11.Text, Form1.Cf_TextBox12.Text)

            Next

        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\item.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub excelRead14()
        '楽天(生活空間)のセット関連レコメンドのファイル／ＣＳＶ作成

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)
        Dim rs As String = create_temporary_table02(tmptbn02, "グループ")

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book01 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book01.GetSheetIndex("レコメンドグループ")
            Dim sheet1 As ISheet = book01.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet1.LastRowNum

            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet1.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("
                        sql2h &= tmptbn02
                        sql2h &= "` ("

                        sql2h &= " `グループ`"
                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"

                        Dim num2 As Integer = 1     'テーブルのナンバリング

                        For c1 As Integer = 1 To 22 Step 3
                            Dim cel03 As ICell = nrow1.GetCell(c1)

                            If IsNothing(cel03) = True Then
                                'セルは空白⇒なにもしない
                            Else
                                'セルは空白ではない⇒SQL組み立て

                                sql2h &= ", "
                                sql2b &= ", "

                                sql2h &= "`pid"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= "'" & nrow1.GetCell(c1).ToString & "'"

                                sql2h &= ",`商品名"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 1).ToString & "'"


                                sql2h &= ",`コメント"
                                sql2h &= num2.ToString("000")
                                sql2h &= "`"

                                sql2b &= ",'" & nrow1.GetCell(c1 + 2).ToString & "'"

                            End If

                            num2 += 1

                        Next
                        sql2b &= ");"
                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        tmptbn03 = "summary" & dt2unepocht(noto)
        rs = create_temporary_table03(tmptbn03)

        If rs = "Complete" Then
            ''テンポラリ作成成功

            ''Excelを開く
            Dim opfname01 As String = Form1.Cf_TextBox0.Text
            Dim rfs As FileStream = File.OpenRead(opfname01)
            Dim book02 As IWorkbook = New XSSFWorkbook(rfs)

            rfs.Close()

            Dim sheetNo As Integer = book02.GetSheetIndex("レコメンド商品")
            Dim sheet2 As ISheet = book02.GetSheetAt(sheetNo)

            Dim lpe As Integer = sheet2.LastRowNum


            For lps As Integer = 1 To lpe Step 1

                Dim nrow1 As IRow = sheet2.GetRow(lps)
                If nrow1 Is Nothing Then
                    ''rowは空である
                    Debug.Print("")
                Else

                    If nrow1.GetCell(0) Is Nothing Or nrow1.GetCell(1) Is Nothing Then
                        ''1番目に値がないのはおかしいのでスキップ
                        Debug.Print("")
                    Else

                        Dim sql2h As String = "INSERT INTO `"
                        Dim sql2b As String = " ) VALUE ("

                        sql2h &= tmptbn03
                        sql2h &= "` ("
                        sql2h &= " `pid01`"
                        sql2h &= ",`group01`"


                        sql2b &= " '" & nrow1.GetCell(0).ToString & "'"
                        sql2b &= ",'" & nrow1.GetCell(1).ToString & "'"
                        sql2b &= ");"

                        rs = sql_result_no(sql2h & sql2b)
                        If rs = "Complete" Then
                        Else
                            Debug.Print(rs)
                        End If
                    End If
                End If
            Next
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If


        'HTML作成
        Dim dTb1 As DataTable = temporarytable_selecti02(tmptbn02, "グループ")

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim flnam As String = "/iframe/recommend/"

                Dim html1 As String = html_template02("スタッフのオススメ商品")

                Dim lpc2 As Integer = 1

                For lpc1 As Integer = 1 To 23 Step 3

                    If IsDBNull(DRow.Item(lpc1)) = True Then
                    Else

                        Dim sql3 As String = "SELECT"
                        sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品番号`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品名`"
                        sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品画像URL`"
                        sql3 &= " FROM `nRms_seikatsukukan_item_newest`"
                        sql3 &= " WHERE"
                        sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）` = '"
                        sql3 &= DRow.Item(lpc1)
                        sql3 &= "' AND `nRms_seikatsukukan_item_newest`.`倉庫指定` = 0"
                        sql3 &= " LIMIT 1;"

                        Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                        If dTb3.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow3 As DataRow In dTb3.Rows
                                'html
                                html1 &= html_template02b("sr", lpc1, DRow3.Item(0), DRow3.Item(3), pnameshp(DRow.Item(lpc1 + 1)), DRow.Item(lpc1 + 2))

                            Next
                        End If

                    End If

                    lpc2 += 1

                Next

                html1 &= html_template03()

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim fn As Integer = DRow.Item(0)
                Dim hofn2 As String = Form1.Cf_TextBox10.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox11.Text, Form1.Cf_TextBox12.Text)

                'UPDATE
                Dim sql4 As String = "UPDATE `"
                sql4 &= tmptbn03
                sql4 &= "` Set `gid01` = "
                sql4 &= lpc3
                sql4 &= " WHERE `group01` = '"
                sql4 &= DRow.Item(25)
                sql4 &= "';"

                rs = sql_result_no(sql4)
                If rs = "Complete" Then
                Else
                    Debug.Print(rs)
                End If

                lpc3 += 1

            Next

        End If

        'csv作成
        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        sql1 = "SELECT"
        sql1 &= " `serial`" '0
        sql1 &= ",`pid01`"  '1
        sql1 &= ",`group01`"    '2
        sql1 &= ",`gid01`"  '3
        sql1 &= " FROM `"
        sql1 &= tmptbn03
        sql1 &= "`"
        dTb1 = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            Dim lpc3 As Integer = 1

            For Each DRow As DataRow In dTb1.Rows

                Dim sql3 As String = "SELECT"
                sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）`"
                sql3 &= ",`nRms_seikatsukukan_item_newest`.`商品番号`"
                sql3 &= ",`nRms_seikatsukukan_item_newest`.`PC用商品説明文`"
                sql3 &= ",`nRms_seikatsukukan_item_newest`.`スマートフォン用商品説明文`"
                sql3 &= " FROM `nRms_seikatsukukan_item_newest`"
                sql3 &= " WHERE"
                sql3 &= " `nRms_seikatsukukan_item_newest`.`商品管理番号（商品URL）` = '"
                sql3 &= DRow.Item(1)
                sql3 &= "' AND `nRms_seikatsukukan_item_newest`.`倉庫指定` = 0"
                sql3 &= " LIMIT 1;"

                Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                If dTb3.Rows.Count = 0 Then
                    'Debug.Print(DRow.Item(c1))
                Else
                    For Each DRow3 As DataRow In dTb3.Rows
                        'CSV

                        Dim pro_des1 As String = ""
                        If IsDBNull(DRow3.Item(2)) = False Then
                            pro_des1 = DRow3.Item(2)
                        End If

                        Dim pro_des2 As String = ""
                        If IsDBNull(DRow3.Item(3)) = False Then
                            pro_des2 = DRow3.Item(3)
                        End If
                        csv1 &= r_pc_csv_template02("sr", DRow3.Item(0), DRow3.Item(1), pro_des1, pro_des2, DRow.Item(3))

                    Next
                End If

            Next

        End If

        'CSVの出力(koko)
        Dim lofn2 As String = CuDr & "\item.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub


    Sub excelRead4()

        Dim sql1 As String = ""

        ''データベースと接続
        Call sql_st()

        ''テンポラリ格納テーブルの作成
        Dim noto As DateTime = Now

        tmptbn02 = "cost" & dt2unepocht(noto)

        sql1 = "CREATE TABLE `"
        sql1 &= tmptbn02
        sql1 &= "` ("
        sql1 &= " `serial` INTEGER UNSIGNED auto_increment primary key"
        sql1 &= ",`scate` VARCHAR(50)"
        sql1 &= ",`pid` VARCHAR(25)"
        sql1 &= ",`cleng` SMALLINT UNSIGNED"
        sql1 &= ",`dispn` VARCHAR(5)"

        sql1 &= ");"
        Dim rs As String = sql_result_no(sql1)

        If rs = "Complete" Then
            ''テンポラリ作成成功
        Else
            ''テンポラリ作成失敗
            Debug.Print(rs)
        End If

        '登録用SQL
        Dim sql2 As String = "INSERT INTO `"
        sql2 &= tmptbn02
        sql2 &= "` ("
        sql2 &= " `scate`"
        sql2 &= ",`pid`"
        sql2 &= ",`cleng`"
        sql2 &= ") VALUE "

        'カテゴリ
        sql1 = "SELECT"
        sql1 &= " `nRms_itemcat_newest`.`商品管理番号（商品URL）`"
        sql1 &= ",`nRms_itemcat_newest`.`表示先カテゴリ`"
        sql1 &= ",CHAR_LENGTH(`表示先カテゴリ`)"
        sql1 &= " FROM `nRms_item_newest` INNER JOIN `nRms_itemcat_newest`"
        sql1 &= " ON `nRms_item_newest`.`商品管理番号（商品URL）` = `nRms_itemcat_newest`.`商品管理番号（商品URL）`"
        sql1 &= " WHERE `nRms_item_newest`.`倉庫指定` = 0"
        sql1 &= " ORDER BY `nRms_itemcat_newest`.`表示先カテゴリ`"
        sql1 &= ";"

        Dim dTb1 As DataTable = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim vcat1 As String = DRow.Item(1)
                Dim iFind1 As Integer = vcat1.LastIndexOf("\"c)

                sql2 &= "("

                If iFind1 < 1 Then
                    ''\が見つからない⇒そのままデータにする
                    sql2 &= "'"
                    sql2 &= vcat1
                    sql2 &= "'"
                Else
                    ''\が見つかった⇒\で切断してデータとする
                    sql2 &= "'"
                    sql2 &= vcat1.Substring(iFind1 + 1)
                    sql2 &= "'"

                End If

                sql2 &= ","
                sql2 &= "'"
                sql2 &= DRow.Item(0)
                sql2 &= "'"
                sql2 &= ","
                sql2 &= DRow.Item(2)
                sql2 &= ")"
                sql2 &= ","

            Next
        End If

        Dim ln1 As Long = sql2.Length
        sql2 = sql2.Substring(0, ln1 - 1)
        sql2 &= ";"
        Dim rst As String = sql_result_no(sql2)
        If rst <> "Complete" Then
            MsgBox("SQLが正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
            Debug.Print(sql1)
        End If

        Dim flnam As String = "iframe/divisio01/"

        ''グループ抽出
        Dim fn As Integer = 1

        sql1 = "SELECT"
        sql1 &= " `scate`"
        sql1 &= "FROM `"
        sql1 &= tmptbn02
        sql1 &= "`"
        sql1 &= "GROUP BY `scate`"
        sql1 &= ";"
        dTb1 = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows


                Dim html1 As String = "<!DOCTYPE html>" & vbCrLf
                html1 &= "<html lang=""ja"">" & vbCrLf
                html1 &= vbCrLf
                html1 &= "	<head>" & vbCrLf
                html1 &= "		<meta charset=""UTF-8"">" & vbCrLf
                html1 &= "		<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & vbCrLf
                html1 &= "		<meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0"">" & vbCrLf
                html1 &= "<title>同一カテゴリ商品</title>" & vbCrLf
                html1 &= "		<link rel=""stylesheet"" href=""css/slick.css"">" & vbCrLf
                html1 &= "		<link rel=""stylesheet"" href=""css/royal.css"">" & vbCrLf
                html1 &= "		<script src=""js/jquery-1.9.1.min.js""></script>" & vbCrLf
                html1 &= "		<script src=""js/slick.min.js""></script>" & vbCrLf
                html1 &= "		<script src=""js/jquery.js""></script>" & vbCrLf
                html1 &= "	</head>" & vbCrLf
                html1 &= vbCrLf
                html1 &= "	<body>" & vbCrLf
                html1 &= "		<div id=""container"">" & vbCrLf
                html1 &= "			<div class=""slider-title"">" & vbCrLf
                html1 &= "				<h3>"
                html1 &= DRow.Item(0)
                html1 &= "</h3>" & vbCrLf
                html1 &= "			</div>" & vbCrLf
                html1 &= "			<div class=""slider""> " & vbCrLf
                html1 &= "				<ul class=""slick"">" & vbCrLf

                Dim lpc2 As Integer = 1

                sql2 = "SELECT "
                sql2 &= " pid"
                sql2 &= ",`scate`"
                sql2 &= ",`serial`"
                sql2 &= " FROM `"
                sql2 &= tmptbn02
                sql2 &= "` WHERE `scate` = '"
                sql2 &= DRow.Item(0)
                sql2 &= "'"

                Dim dTb2 As DataTable = sql_result_return(sql2)
                If dTb2.Rows.Count = 0 Then
                    MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
                Else
                    For Each DRow2 As DataRow In dTb2.Rows

                        Dim sql4 As String = "SELECT"
                        sql4 &= " `nRms_item_newest`.`商品管理番号（商品URL）`"
                        sql4 &= ",`nRms_item_newest`.`商品番号`"
                        sql4 &= ",`product_ledger`.`商品名`"
                        sql4 &= ",`nRms_item_newest`.`商品画像URL`"
                        sql4 &= ",`nRms_item_newest`.`倉庫指定`"
                        sql4 &= " FROM `nRms_item_newest` LEFT JOIN `product_ledger`"
                        sql4 &= " ON `nRms_item_newest`.`商品管理番号（商品URL）` = `product_ledger`.`ID`"
                        sql4 &= " WHERE"
                        sql4 &= " `nRms_item_newest`.`商品管理番号（商品URL）` = '"
                        sql4 &= DRow2.Item(0)
                        sql4 &= "' LIMIT 1;"

                        Dim dTb4 As DataTable = sql_result_return(sql4)

                        If dTb4.Rows.Count = 0 Then
                            'Debug.Print(DRow.Item(c1))
                        Else
                            For Each DRow4 As DataRow In dTb4.Rows
                                'html

                                If DRow4.Item(4) = 1 Then
                                    Debug.Print("")
                                Else

                                    If lpc2 < 20 Then

                                        html1 &= "					<!-- " & lpc2 & "商品 -->" & vbCrLf
                                        html1 &= "					<li>" & vbCrLf
                                        html1 &= "						<a href=""https://item.rakuten.co.jp/royal3000/"
                                        html1 &= DRow4.Item(0)
                                        html1 &= "/"" target=""_parent"">" & vbCrLf
                                        html1 &= "							<img src="""
                                        html1 &= imgsamp(DRow4.Item(3))
                                        html1 &= """>" & vbCrLf
                                        html1 &= "							<p>"

                                        If IsDBNull(DRow4.Item(2)) = True Then
                                        Else
                                            html1 &= DRow4.Item(2)

                                        End If

                                        html1 &= "</p>" & vbCrLf
                                        html1 &= "						</a>" & vbCrLf
                                        html1 &= "					</li>" & vbCrLf

                                    End If

                                    lpc2 += 1

                                End If

                            Next
                        End If

                        Dim sql5 As String = "UPDATE `"
                        sql5 &= tmptbn02
                        sql5 &= "` SET `dispn` = '"
                        sql5 &= fn.ToString("0000")
                        sql5 &= "' WHERE `serial` = "
                        sql5 &= DRow2.Item(2)
                        sql5 &= ";"

                        rs = sql_result_no(sql5)

                        If rs = "Complete" Then
                            ''テンポラリ作成成功
                        Else
                            ''テンポラリ作成失敗
                            Debug.Print(rs)
                        End If

                    Next
                End If

                html1 &= "				</ul>" & vbCrLf
                html1 &= "			</div>" & vbCrLf
                html1 &= "			<!-- .slider END -->" & vbCrLf
                html1 &= "		</div>" & vbCrLf
                html1 &= "		<!-- #container END -->" & vbCrLf
                html1 &= "	</body>" & vbCrLf
                html1 &= "" & vbCrLf
                html1 &= "</html>" & vbCrLf

                Dim lofn1 As String = CuDr & "\temp.html"

                Dim hofn2 As String = Form1.Cf_TextBox1.Text & ":16910/" & flnam & fn.ToString("0000") & ".html"


                '書き込むファイルが既に存在している場合は、上書きする
                Dim sw As New System.IO.StreamWriter(lofn1)
                'TextBox1.Textの内容を書き込む
                sw.Write(html1)
                '閉じる
                sw.Close()

                'FTPへ上記ファイルをアップ
                fil_ftp_up(lofn1, "ftp://" & hofn2, Form1.Cf_TextBox2.Text, Form1.Cf_TextBox3.Text)

                fn += 1

            Next
        End If


        'CSV作成ブロック
        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        sql1 = "SELECT"
        sql1 &= " `pid`"
        sql1 &= ",Max(`cleng`)"
        sql1 &= ",Max(`dispn`)"
        sql1 &= "FROM `"
        sql1 &= tmptbn02
        sql1 &= "`"
        sql1 &= " GROUP BY `pid`"
        sql1 &= " HAVING Max(`dispn`) Is Not Null"
        sql1 &= ";"

        dTb1 = sql_result_return(sql1)
        If dTb1.Rows.Count = 0 Then
            MsgBox("データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim sql3 As String = "SELECT"
                sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）`"
                sql3 &= ",`nRms_item_newest`.`商品番号`"
                sql3 &= ",`product_ledger`.`商品名`"
                sql3 &= ",`nRms_item_newest`.`商品画像URL`"
                sql3 &= ",`nRms_item_newest`.`PC用商品説明文`"
                sql3 &= ",`nRms_item_newest`.`スマートフォン用商品説明文`"
                sql3 &= " FROM `nRms_item_newest` LEFT JOIN `product_ledger`"
                sql3 &= " ON `nRms_item_newest`.`商品管理番号（商品URL）` = `product_ledger`.`ID`"
                sql3 &= " WHERE"
                sql3 &= " `nRms_item_newest`.`商品管理番号（商品URL）` = '"
                sql3 &= DRow.Item(0)
                sql3 &= "' AND `nRms_item_newest`.`倉庫指定` = 0"
                sql3 &= " LIMIT 1;"

                Dim dTb3 As DataTable = sql_result_return(sql3.ToString)

                If dTb3.Rows.Count = 0 Then
                    'Debug.Print(DRow.Item(c1))
                Else
                    For Each DRow3 As DataRow In dTb3.Rows
                        'csv
                        Dim des0 As String = ""
                        Dim des1 As String = ""

                        csv1 &= """u"","""
                        csv1 &= DRow3.Item(0)
                        csv1 &= ""","""
                        csv1 &= DRow3.Item(1)
                        csv1 &= ""","""

                        If IsDBNull(DRow3.Item(4)) Then
                            des0 = ""
                        Else
                            des0 = DRow3.Item(4)
                        End If

                        Dim po0 As Integer = des0.IndexOf("<!--同一カテ開始-->")
                        If 0 <= po0 Then
                            '含まれています
                            Dim po1 As Integer = des0.IndexOf("<!--同一カテ終了-->") + 13
                            des1 = des0.Replace(des0.Substring(po0, po1 - po0), "[RECO]")

                        Else
                            '含まれていません
                            des1 = des0 & "[RECO]"
                        End If

                        des1 = des1.Replace("""", """""")

                        Dim cont1 As String = "<!--同一カテ開始-->"
                        cont1 &= "<IFRAME src=""""https://www.rakuten.ne.jp/gold/royal3000/"
                        cont1 &= flnam

                        cont1 &= DRow.Item(2)
                        cont1 &= ".html"

                        cont1 &= """"" frameborder=""""0"""" scrolling=""""no"""" width=""""1000"""" height="""""
                        cont1 &= "330"
                        cont1 &= """""></IFRAME>"
                        cont1 &= "<!--同一カテ終了-->"

                        des1 = des1.Replace("[RECO]", cont1)
                        csv1 &= des1
                        csv1 &= """"
                        csv1 &= ","
                        csv1 &= """"

                        ''スマホ用
                        Dim des2 As String = ""
                        Dim des3 As String = ""

                        If IsDBNull(DRow3.Item(5)) Then
                            des2 = ""
                        Else
                            des2 = DRow3.Item(5)
                        End If

                        Dim po2 As Integer = des2.IndexOf("<!--同一カテ開始-->")
                        If 0 <= po2 Then
                            '含まれています
                            Dim po3 As Integer = des2.IndexOf("<!--同一カテ終了-->") + 13
                            des3 = des2.Replace(des2.Substring(po2, po3 - po2), "[RECO]")

                        Else
                            '含まれていません
                            des3 = des2 & "[RECO]"
                        End If

                        des3 = des3.Replace("""", """""")

                        Dim cont2 As String = "<!--同一カテ開始-->"
                        cont2 &= "<IFRAME ="""""""" src=""""https://www.rakuten.ne.jp/gold/royal3000/"
                        cont2 &= flnam

                        cont2 &= DRow.Item(2)
                        cont2 &= ".html"

                        cont2 &= """"" frameborder=""""0"""" scrolling=""""no"""" width=""""1000"""" height=""""330""""></IFRAME ="""""""">"
                        cont2 &= "<!--同一カテ終了-->"

                        des3 = des3.Replace("[RECO]", cont2)
                        csv1 &= des3
                        csv1 &= """"
                        csv1 &= vbCrLf


                        Dim sql4 As String = "UPDATE `nRms_item_newest` SET `PC用商品説明文` = "
                        sql4 &= "'"
                        sql4 &= des1.Replace("""""", """")
                        sql4 &= "'"
                        sql4 &= ","
                        sql4 &= " `スマートフォン用商品説明文` = "
                        sql4 &= "'"
                        sql4 &= des3.Replace("""""", """")
                        sql4 &= "'"

                        sql4 &= " WHERE `商品管理番号（商品URL）` = '"
                        sql4 &= DRow3.Item(0)
                        sql4 &= "';"

                        rs = sql_result_no(sql4)

                    Next
                End If

            Next
        End If

        'CSVの出力
        Dim lofn2 As String = CuDr & "\item.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()


        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub initializ_work01(ByVal shop As String)

        ''レコメンドを初期化する

        ''データベースと接続
        Call sql_st()

        Dim csv1 As String = """コントロールカラム"",""商品管理番号（商品URL）"",""商品番号"",""PC用商品説明文"",""スマートフォン用商品説明文""" & vbCrLf

        Dim sql1 As String = "SELECT"
        sql1 &= " `商品管理番号（商品URL）`"
        sql1 &= ",`商品番号`"
        sql1 &= ",`商品画像URL`"
        sql1 &= ",`PC用商品説明文`"
        sql1 &= ",`スマートフォン用商品説明文`"

        Select Case shop
            Case "rr"
                '楽天ロイヤル用
                sql1 &= " FROM `nRms_item_newest` "

            Case "sr"
                '生活空間楽天用
                sql1 &= " FROM `nRms_seikatsukukan_item_newest` "

            Case Else
        End Select




        sql1 &= " WHERE"
        sql1 &= " `倉庫指定` = 0"
        sql1 &= ";"

        Dim dTb1 As DataTable = sql_result_return(sql1)
        If dTb1.Rows.Count = 0 Then
            ''レコードはない
        Else
            For Each DRow1 As DataRow In dTb1.Rows

                Dim des0 As String = ""
                Dim po0 As Integer = 0
                Dim po1 As Integer = 0
                Dim chk1 As Integer = 0

                If IsDBNull(DRow1.Item(3)) Then
                    des0 = ""
                Else
                    des0 = DRow1.Item(3)
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--同一カテ開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--同一カテ終了-->") + 13
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--サイズ関連開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--サイズ関連終了-->") + 14
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--類似関連開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--類似関連終了-->") + 13
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--セット関連開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--セット関連終了-->") + 14
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--おすすめ商品開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--おすすめ商品終了-->") + 15
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                ''スマホ用
                Dim des2 As String = ""
                Dim po2 As Integer = 0
                Dim po3 As Integer = 0

                If IsDBNull(DRow1.Item(4)) Then
                    des2 = ""
                Else
                    des2 = DRow1.Item(4)
                End If

                po2 = 0
                po2 = des2.IndexOf("<!--同一カテ開始-->")
                If 0 <= po2 Then
                    '含まれています
                    po3 = 0
                    po3 = des2.IndexOf("<!--同一カテ終了-->") + 13
                    des2 = des2.Replace(des2.Substring(po2, po3 - po2), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po2 = 0
                po2 = des2.IndexOf("<!--サイズ関連開始-->")
                If 0 <= po2 Then
                    '含まれています
                    po3 = 0
                    po3 = des2.IndexOf("<!--サイズ関連終了-->") + 14
                    des2 = des2.Replace(des2.Substring(po2, po3 - po2), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po2 = 0
                po2 = des2.IndexOf("<!--類似関連開始-->")
                If 0 <= po2 Then
                    '含まれています
                    po3 = 0
                    po3 = des2.IndexOf("<!--類似関連終了-->") + 13
                    des2 = des2.Replace(des2.Substring(po2, po3 - po2), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po2 = 0
                po2 = des2.IndexOf("<!--セット関連開始-->")
                If 0 <= po2 Then
                    '含まれています
                    po3 = 0
                    po3 = des2.IndexOf("<!--セット関連終了-->") + 14
                    des2 = des2.Replace(des2.Substring(po2, po3 - po2), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po2 = 0
                po2 = des2.IndexOf("<!--おすすめ商品開始-->")
                If 0 <= po2 Then
                    '含まれています
                    po3 = 0
                    po3 = des2.IndexOf("<!--おすすめ商品終了-->") + 15
                    des2 = des2.Replace(des2.Substring(po2, po3 - po2), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If



                If chk1 > 0 Then

                    Dim sql4 As String = "UPDATE "
                    Select Case shop
                        Case "rr"
                            '楽天ロイヤル用
                            sql4 &= "`nRms_item_newest`"

                        Case "sr"
                            '生活空間楽天用
                            sql4 &= "`nRms_seikatsukukan_item_newest`"

                        Case Else
                    End Select


                    sql4 &= " Set `PC用商品説明文` = "
                    sql4 &= "'"
                    sql4 &= des0
                    sql4 &= "'"
                    sql4 &= ","
                    sql4 &= " `スマートフォン用商品説明文` = "
                    sql4 &= "'"
                    sql4 &= des2
                    sql4 &= "'"

                    sql4 &= " WHERE `商品管理番号（商品URL）` = '"
                    sql4 &= DRow1.Item(0)
                    sql4 &= "';"

                    Dim rs As String = sql_result_no(sql4)



                    csv1 &= """u"","""
                    csv1 &= DRow1.Item(0)
                    csv1 &= ""","""
                    csv1 &= DRow1.Item(1)
                    csv1 &= ""","""
                    csv1 &= des0.Replace("""", """""")
                    csv1 &= ""","""
                    csv1 &= des2.Replace("""", """""")
                    csv1 &= """"
                    csv1 &= vbCrLf


                End If


            Next
        End If

        Dim lofn2 As String = CuDr & "\item.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub initializ_work02(ByVal shop As String)

        ''レコメンドを初期化する(yahoo用)

        ''データベースと接続
        Call sql_st()

        Dim csv1 As String = """code"",""caption""" & vbCrLf

        Dim sql1 As String = "SELECT"
        sql1 &= " `code`"
        sql1 &= ",`caption`"

        Select Case shop
            Case "ry"
                'ロイヤルYahoo!用
                sql1 &= " FROM `shopping_yahoo_data_newest` "

            Case "sy"
                '生活空間Yahoo!用
                sql1 &= " FROM `seikatsukukan_data_newest` "

            Case "my"
                'マザープラスYahoo!用
                sql1 &= " FROM `motherplusstore_data_newest` "

            Case Else
        End Select

        sql1 &= " WHERE"
        sql1 &= " `display` = 1"
        sql1 &= ";"

        Dim dTb1 As DataTable = sql_result_return(sql1)
        If dTb1.Rows.Count = 0 Then
            ''レコードはない
        Else
            For Each DRow1 As DataRow In dTb1.Rows

                Dim des0 As String = ""
                Dim po0 As Integer = 0
                Dim po1 As Integer = 0
                Dim chk1 As Integer = 0

                If IsDBNull(DRow1.Item(1)) Then
                    des0 = ""
                Else
                    des0 = DRow1.Item(1)
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--同一カテ開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--同一カテ終了-->") + 13
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--サイズ関連開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--サイズ関連終了-->") + 14
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--類似関連開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--類似関連終了-->") + 13
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If

                po0 = 0
                po0 = des0.IndexOf("<!--セット関連開始-->")
                If 0 <= po0 Then
                    '含まれています
                    po1 = 0
                    po1 = des0.IndexOf("<!--セット関連終了-->") + 14
                    des0 = des0.Replace(des0.Substring(po0, po1 - po0), "")
                    chk1 += 1

                Else
                    '含まれていません
                End If


                If chk1 > 0 Then

                    Dim sql4 As String = "UPDATE "
                    Select Case shop
                        Case "ry"
                            'ロイヤルYahoo!用
                            sql4 &= " `shopping_yahoo_data_newest` "

                        Case "sy"
                            '生活空間Yahoo!用
                            sql4 &= " `seikatsukukan_data_newest` "

                        Case "my"
                            'マザープラスYahoo!用
                            sql4 &= " `motherplusstore_data_newest` "

                        Case Else
                    End Select

                    sql4 &= " Set `caption` = "
                    sql4 &= "'"
                    sql4 &= des0
                    sql4 &= "'"

                    sql4 &= " WHERE `code` = '"
                    sql4 &= DRow1.Item(0)
                    sql4 &= "';"

                    Dim rs As String = sql_result_no(sql4)



                    csv1 &= """"
                    csv1 &= DRow1.Item(0)
                    csv1 &= ""","""
                    csv1 &= des0.Replace("""", """""")
                    csv1 &= """"
                    csv1 &= vbCrLf

                End If


            Next
        End If

        Dim lofn2 As String = CuDr & "\data_spy.csv"

        '書き込むファイルが既に存在している場合は、上書きする
        Dim sw2 As New System.IO.StreamWriter(lofn2, False, System.Text.Encoding.GetEncoding("shift_jis"))
        'TextBox1.Textの内容を書き込む
        sw2.Write(csv1)
        '閉じる
        sw2.Close()

        ''データベース切断
        Call sql_cl()

        MsgBox("完了しました。", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "情報")

    End Sub

    Sub fil_ftp_up(ByVal upFile As String, ByVal upUrl As String, ByVal uid As String, ByVal ups As String)

        'アップロード先のURI
        Dim u As New Uri(upUrl)

        'FtpWebRequestの作成
        Dim ftpReq As System.Net.FtpWebRequest =
            CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
        'ログインユーザー名とパスワードを設定
        ftpReq.Credentials = New System.Net.NetworkCredential(uid, ups)
        'MethodにWebRequestMethods.Ftp.UploadFile("STOR")を設定
        ftpReq.Method = System.Net.WebRequestMethods.Ftp.UploadFile
        '要求の完了後に接続を閉じる
        ftpReq.KeepAlive = False
        'ASCIIモードで転送する
        ftpReq.UseBinary = False
        'PASVモードを無効にする
        ftpReq.UsePassive = True

        'ファイルをアップロードするためのStreamを取得
        Dim reqStrm As System.IO.Stream = ftpReq.GetRequestStream()
        'アップロードするファイルを開く
        Dim fs As New System.IO.FileStream(
            upFile, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        'アップロードStreamに書き込む
        Dim buffer(1023) As Byte
        While True
            Dim readSize As Integer = fs.Read(buffer, 0, buffer.Length)
            If readSize = 0 Then
                Exit While
            End If
            reqStrm.Write(buffer, 0, readSize)
        End While
        fs.Close()
        reqStrm.Close()

        'FtpWebResponseを取得
        Dim ftpRes As System.Net.FtpWebResponse =
            CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
        'FTPサーバーから送信されたステータスを表示
        Console.WriteLine("{0}: {1}", ftpRes.StatusCode, ftpRes.StatusDescription)
        '閉じる
        ftpRes.Close()


    End Sub



End Module
