Imports TODO
Imports System.Xml

Partial Class CH_GQ_GQ0590
    Inherits TODO.clsCHTodoWeb
    Private WithEvents Grid As clsSimpleGridPage
    Private mCurrentRowNumber As Long

    WithEvents oSimpleGrid As clsSimpleGridPage
    Dim pagepos, sqlaa, headstr, headaa, sqluse, cmd, pagetext
    Dim errmsg As String


    Sub Main()

        XMLDoc.loadXML(xmlData)

        cmd = getField("CMD")
        ' 功能流程控制  
        Select Case cmd
            Case "btn_excel"                'Excel
                Call QueryExcel()
            Case "INQ_TITLE"
                Call INQ_TITLE()

        End Select

        Call WriteXMLData()
        Response.End()

    End Sub
    Sub INQ_TITLE()
        'Dim GQXX_COMP_CD As String = Trim(getField("GQXX_COMP_CD"))
        Dim GQXX_SEASON_CD As String = Trim(getField("GQXX_SEASON_CD"))


        errmsg = ""
        Dim objDataservice As TODO.clsDataService = New TODO.clsDataService(clsDataService.DsDBName.GQ02)
        Dim objDataReader As OleDb.OleDbDataReader
        Dim sqlStr As String
        'Dim GQ01_NM_CA As String
        Dim GQ05_SEASON_NM As String


        ''*******讀取GQ01,關係企業代碼檔
        'sqlStr = " select GQ01_NM_CA from GQ01 where GQ01_COMP_CD = '" & GQXX_COMP_CD & "'"

        'objDataReader = objDataservice.ExecuteReader(sqlStr)

        'If objDataReader.Read() Then
        '    If Not IsDBNull(objDataReader(0)) Then
        '        GQ01_NM_CA = objDataReader(0)
        '    Else
        '        GQ01_NM_CA = ""
        '    End If
        'Else
        '    GQ01_NM_CA = ""
        'End If
        'objDataReader.Close()


        '*******讀取GQ05,季節月份代碼檔
        sqlStr = " select GQ05_SEASON_NM from GQ05 where GQ05_SEASON_CD = '" & GQXX_SEASON_CD & "'"

        objDataReader = objDataservice.ExecuteReader(sqlStr)

        If objDataReader.Read() Then
            If Not IsDBNull(objDataReader(0)) Then
                GQ05_SEASON_NM = objDataReader(0)
            Else
                GQ05_SEASON_NM = ""
            End If
        Else
            GQ05_SEASON_NM = ""
        End If
        objDataReader.Close()


        xmlData = "<Collection>"
        xmlData = xmlData & "<MSG></MSG>"
        'xmlData = xmlData & "<GQ01_NM_CA>" & GQ01_NM_CA & "</GQ01_NM_CA>"
        xmlData = xmlData & "<GQ05_SEASON_NM>" & GQ05_SEASON_NM & "</GQ05_SEASON_NM>"
        xmlData = xmlData & "</Collection>"
        objDataReader.Close()

    End Sub
    Sub QueryExcel()

        Dim GQXX_COMP_CD As String = Trim(getField("GQXX_COMP_CD"))
        Dim GQXX_DATA_YY As String = Trim(getField("GQXX_DATA_YY"))
        Dim GQXX_SEASON_CD As String = Trim(getField("GQXX_SEASON_CD"))

        Dim i As Integer
        Dim title As Object
        Dim Titletext As String
        Dim sqlStr As String
        Dim oRootNode, table_node, tbody_node, tr_node, td_node, thead1_node

        Dim objDataservice As TODO.clsDataService = New TODO.clsDataService(clsDataService.DsDBName.GQ02)
        Dim objDataReader As OleDb.OleDbDataReader

        oRootNode = XMLDoc.createElement("Collection")
        XMLDoc.documentElement = oRootNode

        '==================================================
        '******************* 金融資產 *********************
        '==================================================

        sqlStr = "select GQ22_ITEM_NM, GQ24_ITEM_NM "
        'sqlStr = sqlStr & ", GQ24_START_TWD, GQ24_ADD_TWD "
        'sqlStr = sqlStr & ",GQ24_INVT_TWD,GQ24_LESS_TWD ,GQ24_DIFF_TWD"    
        sqlStr = sqlStr & ",sum(nvl(GQ24_FINAL_TWD,0)) ,sum(nvl(GQ24_DIVIDEND_TWD,0)) ,sum(nvl(GQ24_UNIT_CNT,0)) ,sum(nvl(GQ24_PRICE_TWD,0)) ,GQ01_NM_CA "
        sqlStr = sqlStr & ",GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " from GQ24, GQ22, GQ01"
        sqlStr = sqlStr & " where GQ24_ITEM_CD = GQ22_ITEM_CD"
        sqlStr = sqlStr & "   and GQ24_COMP_CD = GQ01_COMP_CD"
        sqlStr = sqlStr & "   and GQ22_DATA_CD ='D'"

        If GQXX_DATA_YY <> "" Then
            sqlStr = sqlStr & " AND GQ24_DATA_YY = '" & GQXX_DATA_YY & "'"
        End If
        If GQXX_SEASON_CD <> "BLANK" Then
            sqlStr = sqlStr & " AND GQ24_SEASON_CD = '" & GQXX_SEASON_CD & "'"
        End If
        sqlStr = sqlStr & " group by GQ22_ITEM_NM,GQ24_ITEM_NM,GQ01_NM_CA,GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " order by GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP,GQ01_NM_CA"

        ' 建立 TABLE 
        table_node = XMLDoc.createElement("TABLE")
        oRootNode.appendChild(table_node)
        table_node.setAttribute("name", "金融資產")
        'table_node.SetAttribute("border", "1")

        '***************** 建立 表頭
        thead1_node = XMLDoc.createElement("THEAD")
        table_node.appendChild(thead1_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = "金融資產"
        tr_node.appendChild(td_node)


        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)

        'headstr = " 項目類別@1_被投資標的@1_期初@1_本期增加@1_本期減少@3_期末@3_"
        'headstr = headstr & " @1_ @1_ @1_ @1_ @1_轉列長期投資@1_本期處分@2_ @1_　@1_ @1_"
        'headstr = headstr & " @1_ @1_投資金額@1_投資金額@1_投資金額@1_投資金額@1_處分(損)益@1_投資金額@1_單位數@1_市價@1_"

        headstr = " 項目類別@1_被投資標的@1_期末投資金額@1_股利收入@1_單位數@1_市價@1_投資公司@1_"

        Titletext = (Replace(headstr, "@1", ""))
        'Titletext = (Replace(Titletext, "@2", "_"))
        'Titletext = (Replace(Titletext, "@3", "__"))
        title = Split(Titletext, "_")

        For i = LBound(title) To UBound(title)
            td_node = XMLDoc.createElement("TD")
            td_node.text = title(i)
            tr_node.appendChild(td_node)

            '' 斷行
            'If i = 10 Or i = 20 Then
            '    tr_node = XMLDoc.createElement("TR")
            '    thead1_node.appendChild(tr_node)

            'End If

        Next

        '***************** 建立 表身體
        'tbody_node = XMLDoc.createElement("TBODY")
        'table_node.appendChild(tbody_node)

        objDataReader = objDataservice.ExecuteReader(sqlStr)

        Dim OLD_GROUP_SUB, TEMP_GROUP_SUB As String
        Dim OLD_GROUP, TEMP_GROUP As String
        Dim OLD_ITEM_NM, TEMP_ITEM_NM As String
        'Dim S_START_TWD, S_ADD_TWD, S_INVT_TWD, S_LESS_TWD, S_DIFF_TWD, S_FINAL_TWD, S_UNIT_CNT, S_PRICE_TWD As Decimal
        'Dim T_START_TWD, T_ADD_TWD, T_INVT_TWD, T_LESS_TWD, T_DIFF_TWD, T_FINAL_TWD, T_UNIT_CNT, T_PRICE_TWD As Decimal
        Dim S_FINAL_TWD, S_UNIT_CNT, S_PRICE_TWD, S_REV_TWD, S_DIVIDEND_TWD As Decimal
        Dim T_FINAL_TWD, T_UNIT_CNT, T_PRICE_TWD, T_REV_TWD, T_DIVIDEND_TWD As Decimal
        Dim iCnt As Decimal

        iCnt = 1
        Do While objDataReader.Read()

            '小計
            If Not IsDBNull(objDataReader(7)) Then
                TEMP_GROUP_SUB = objDataReader(7)
            Else
                TEMP_GROUP_SUB = ""
            End If

            If Trim(OLD_GROUP_SUB) = "" Then
                If Not IsDBNull(objDataReader(7)) Then
                    OLD_GROUP_SUB = objDataReader(7)
                Else
                    OLD_GROUP_SUB = ""
                End If
            End If



            If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "小計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_FINAL_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_DIVIDEND_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_UNIT_CNT

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_PRICE_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                '小計歸零
                'S_START_TWD = 0
                'S_ADD_TWD = 0
                'S_INVT_TWD = 0
                'S_LESS_TWD = 0
                'S_DIFF_TWD = 0
                S_FINAL_TWD = 0
                S_UNIT_CNT = 0
                S_PRICE_TWD = 0
                S_DIVIDEND_TWD = 0

                OLD_GROUP_SUB = TEMP_GROUP_SUB
            End If

            '合計
            If Not IsDBNull(objDataReader(8)) Then
                TEMP_GROUP = objDataReader(8)
            Else
                TEMP_GROUP = ""
            End If

            If Trim(OLD_GROUP) = "" Then
                If Not IsDBNull(objDataReader(8)) Then
                    OLD_GROUP = objDataReader(8)
                Else
                    OLD_GROUP = ""
                End If
            End If



            If OLD_GROUP <> TEMP_GROUP Then

                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "合計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_FINAL_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_DIVIDEND_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_UNIT_CNT

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_PRICE_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                '合計歸零
                'T_START_TWD = 0
                'T_ADD_TWD = 0
                'T_INVT_TWD = 0
                'T_LESS_TWD = 0
                'T_DIFF_TWD = 0
                T_FINAL_TWD = 0
                T_UNIT_CNT = 0
                T_PRICE_TWD = 0
                T_DIVIDEND_TWD = 0

                OLD_GROUP = TEMP_GROUP
            End If


            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            For i = 0 To objDataReader.FieldCount - 1
                Select Case i
                    Case 0
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            TEMP_ITEM_NM = ""
                        Else
                            TEMP_ITEM_NM = objDataReader(i)
                        End If

                        If OLD_ITEM_NM = "" And iCnt = 1 Then
                            If Not IsDBNull(objDataReader(i)) Then
                                OLD_ITEM_NM = objDataReader(i)
                                td_node.text = objDataReader(i)
                            Else
                                OLD_ITEM_NM = ""
                                td_node.text = ""
                            End If
                        End If

                        If OLD_ITEM_NM <> TEMP_ITEM_NM Then
                            OLD_ITEM_NM = TEMP_ITEM_NM
                            td_node.text = OLD_ITEM_NM
                        End If

                    Case 2
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_FINAL_TWD = S_FINAL_TWD + 0
                            T_FINAL_TWD = T_FINAL_TWD + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_FINAL_TWD = S_FINAL_TWD + objDataReader(i)
                            End If
                            T_FINAL_TWD = T_FINAL_TWD + objDataReader(i)
                        End If

                    Case 3
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_DIVIDEND_TWD = S_DIVIDEND_TWD + 0
                            T_DIVIDEND_TWD = T_DIVIDEND_TWD + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_DIVIDEND_TWD = S_DIVIDEND_TWD + objDataReader(i)
                            End If
                            T_DIVIDEND_TWD = T_DIVIDEND_TWD + objDataReader(i)
                        End If

                    Case 4
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_UNIT_CNT = S_UNIT_CNT + 0
                            T_UNIT_CNT = T_UNIT_CNT + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_UNIT_CNT = S_UNIT_CNT + objDataReader(i)
                            End If
                            T_UNIT_CNT = T_UNIT_CNT + objDataReader(i)
                        End If

                    Case 5
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_PRICE_TWD = S_PRICE_TWD + 0
                            T_PRICE_TWD = T_PRICE_TWD + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_PRICE_TWD = S_PRICE_TWD + objDataReader(i)
                            End If
                            T_PRICE_TWD = T_PRICE_TWD + objDataReader(i)
                        End If

                    Case 7, 8

                    Case Else
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                        Else
                            td_node.text = objDataReader(i)
                        End If

                End Select
            Next
            iCnt = iCnt + 1
        Loop

        objDataReader.Close()

        If OLD_GROUP_SUB <> "" Then
            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = "小計"

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_FINAL_TWD

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_DIVIDEND_TWD

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_UNIT_CNT

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_PRICE_TWD

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

        End If

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = "合計"

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_FINAL_TWD

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_DIVIDEND_TWD

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_UNIT_CNT

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_PRICE_TWD

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        '==================================================
        '****************** 長期投資-權益法 ***************
        '==================================================

        sqlStr = "select GQ22_ITEM_NM, GQ25_ITEM_NM, sum(nvl(GQ25_DIFFINV_TWD,0)), sum(nvl(GQ25_CASH_TWD,0)),sum(nvl(GQ25_FSTOCK_CNT,0))"
        sqlStr = sqlStr & ",sum(nvl(GQ25_FINAL_TWD,0)),sum(nvl(GQ25_REV_TWD,0)),sum(nvl(GQ25_FINVEST_PCT,0)),sum(nvl(GQ25_PRICE_TWD,0)),GQ01_NM_CA"
        sqlStr = sqlStr & ",GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " from GQ25, GQ22, GQ01"
        sqlStr = sqlStr & " where GQ25_ITEM_CD = GQ22_ITEM_CD"
        sqlStr = sqlStr & "   and GQ25_COMP_CD = GQ01_COMP_CD"
        sqlStr = sqlStr & "   and GQ22_DATA_CD ='E'"

        'If GQXX_COMP_CD <> "BLANK" Then
        '    sqlStr = sqlStr & " AND GQ25_COMP_CD = '" & GQXX_COMP_CD & "'"
        'End If

        If GQXX_DATA_YY <> "" Then
            sqlStr = sqlStr & " AND GQ25_DATA_YY = '" & GQXX_DATA_YY & "'"
        End If

        If GQXX_SEASON_CD <> "BLANK" Then
            sqlStr = sqlStr & " AND GQ25_SEASON_CD = '" & GQXX_SEASON_CD & "'"
        End If
        sqlStr = sqlStr & " group by GQ22_ITEM_NM,GQ25_ITEM_NM,GQ01_NM_CA,GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " order by GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP,GQ01_NM_CA"

        ' 建立 TABLE 
        table_node = XMLDoc.createElement("TABLE")
        oRootNode.appendChild(table_node)
        table_node.setAttribute("name", "長期投資-權益法")
        'table_node.SetAttribute("border", "1")

        '***************** 建立 表頭
        thead1_node = XMLDoc.createElement("THEAD")
        table_node.appendChild(thead1_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = "長期投資-權益法"
        tr_node.appendChild(td_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)


        headstr = " 項目類別@1_項目名稱@1_已認列投資(損)益@1_現金股利@1_投資股數(股)@1_期末投資金額@1_董監收入@1_投資比例(%)@1_市價@1_投資公司@1_"
        Titletext = (Replace(headstr, "@1", ""))
        title = Split(Titletext, "_")

        For i = LBound(title) To UBound(title)
            td_node = XMLDoc.createElement("TD")
            td_node.text = title(i)
            tr_node.appendChild(td_node)

            '' 斷行
            'If i = 18 Then
            '    tr_node = XMLDoc.createElement("TR")
            '    thead1_node.appendChild(tr_node)

            'End If

        Next
        '***************** 建立 表身體
        'tbody_node = XMLDoc.createElement("TBODY")
        'table_node.appendChild(tbody_node)

        objDataReader = objDataservice.ExecuteReader(sqlStr)

        OLD_GROUP_SUB = ""
        TEMP_GROUP_SUB = ""
        OLD_GROUP = ""
        TEMP_GROUP = ""
        S_FINAL_TWD = 0
        T_FINAL_TWD = 0
        S_PRICE_TWD = 0
        T_PRICE_TWD = 0
        S_REV_TWD = 0
        T_REV_TWD = 0

        Dim S_DIFFINV_TWD, S_CASH_TWD, S_FSTOCK_CNT As Decimal
        Dim T_DIFFINV_TWD, T_CASH_TWD, T_FSTOCK_CNT As Decimal

        iCnt = 1
        Do While objDataReader.Read()
            '************ 小計 ******************

            If Not IsDBNull(objDataReader(10)) Then
                TEMP_GROUP_SUB = objDataReader(10)
            Else
                TEMP_GROUP_SUB = ""
            End If

            If Trim(OLD_GROUP_SUB) = "" Then
                If Not IsDBNull(objDataReader(10)) Then
                    OLD_GROUP_SUB = objDataReader(10)
                Else
                    OLD_GROUP_SUB = ""
                End If
            End If


            If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "小計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_DIFFINV_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_CASH_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_FSTOCK_CNT

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_FINAL_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_REV_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_PRICE_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                '小計歸零
                S_DIFFINV_TWD = 0
                S_CASH_TWD = 0
                S_FSTOCK_CNT = 0
                S_FINAL_TWD = 0
                S_PRICE_TWD = 0
                S_REV_TWD = 0


                OLD_GROUP_SUB = TEMP_GROUP_SUB
            End If

            '************ 合計 ******************
            If Not IsDBNull(objDataReader(11)) Then
                TEMP_GROUP = objDataReader(11)
            Else
                TEMP_GROUP = ""
            End If

            If Trim(OLD_GROUP) = "" Then
                If Not IsDBNull(objDataReader(11)) Then
                    OLD_GROUP = objDataReader(11)
                Else
                    OLD_GROUP = ""
                End If
            End If

            If TEMP_GROUP <> OLD_GROUP Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "合計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_DIFFINV_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_CASH_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_FSTOCK_CNT

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_FINAL_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_REV_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_PRICE_TWD

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                '合計歸零
                T_DIFFINV_TWD = 0
                T_CASH_TWD = 0
                T_FSTOCK_CNT = 0
                T_FINAL_TWD = 0
                T_PRICE_TWD = 0
                T_REV_TWD = 0

                OLD_GROUP = TEMP_GROUP
            End If

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            For i = 0 To objDataReader.FieldCount - 1

                Select Case i

                    Case 0
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            TEMP_ITEM_NM = ""
                        Else
                            TEMP_ITEM_NM = objDataReader(i)
                        End If

                        If OLD_ITEM_NM = "" And iCnt = 1 Then
                            If Not IsDBNull(objDataReader(i)) Then
                                OLD_ITEM_NM = objDataReader(i)
                                td_node.text = objDataReader(i)
                            Else
                                OLD_ITEM_NM = ""
                                td_node.text = ""
                            End If
                        End If

                        If OLD_ITEM_NM <> TEMP_ITEM_NM Then
                            OLD_ITEM_NM = TEMP_ITEM_NM
                            td_node.text = OLD_ITEM_NM
                        End If

                    Case 2
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_DIFFINV_TWD = S_DIFFINV_TWD + 0
                            T_DIFFINV_TWD = T_DIFFINV_TWD + 0
                        Else
                            td_node.text = objDataReader(i)
                            If OLD_GROUP_SUB <> "" Then
                                S_DIFFINV_TWD = S_DIFFINV_TWD + objDataReader(i)
                            End If
                            T_DIFFINV_TWD = T_DIFFINV_TWD + objDataReader(i)
                        End If

                    Case 3
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            S_CASH_TWD = S_CASH_TWD + 0
                            T_CASH_TWD = T_CASH_TWD + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_CASH_TWD = S_CASH_TWD + objDataReader(i)
                            End If
                            T_CASH_TWD = T_CASH_TWD + objDataReader(i)
                        End If

                    Case 4
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            S_FSTOCK_CNT = S_FSTOCK_CNT + 0
                            T_FSTOCK_CNT = T_FSTOCK_CNT + 0
                        Else
                            td_node.text = objDataReader(i)
                            If OLD_GROUP_SUB <> "" Then
                                S_FSTOCK_CNT = S_FSTOCK_CNT + objDataReader(i)
                            End If
                            T_FSTOCK_CNT = T_FSTOCK_CNT + objDataReader(i)
                        End If

                    Case 5
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            S_FINAL_TWD = S_FINAL_TWD + 0
                            T_FINAL_TWD = T_FINAL_TWD + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_FINAL_TWD = S_FINAL_TWD + objDataReader(i)
                            End If
                            T_FINAL_TWD = T_FINAL_TWD + objDataReader(i)
                        End If

                        'Case 6
                        '    td_node = XMLDoc.createElement("TD")
                        '    tr_node.appendChild(td_node)
                        '    If IsDBNull(objDataReader(i)) Then

                        '    Else
                        '        td_node.text = objDataReader(i)
                        '    End If


                    Case 6
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            S_REV_TWD = S_REV_TWD + 0
                            T_REV_TWD = T_REV_TWD + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_REV_TWD = S_REV_TWD + objDataReader(i)
                            End If
                            T_REV_TWD = T_REV_TWD + objDataReader(i)
                        End If

                    Case 8
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            S_PRICE_TWD = S_PRICE_TWD + 0
                            T_PRICE_TWD = T_PRICE_TWD + 0
                        Else
                            td_node.text = objDataReader(i)
                            If OLD_GROUP_SUB <> "" Then
                                S_PRICE_TWD = S_PRICE_TWD + objDataReader(i)
                            End If
                            T_PRICE_TWD = T_PRICE_TWD + objDataReader(i)
                        End If

                    Case 10, 11

                    Case Else
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                        Else
                            td_node.text = objDataReader(i)
                        End If

                End Select
            Next
            iCnt = iCnt + 1
        Loop

        objDataReader.Close()

        If OLD_GROUP_SUB <> "" Then
            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = "小計"

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_DIFFINV_TWD

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_CASH_TWD

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_FSTOCK_CNT

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_FINAL_TWD

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_REV_TWD

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_PRICE_TWD

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

        End If

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = "合計"

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_DIFFINV_TWD

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_CASH_TWD

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_FSTOCK_CNT

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_FINAL_TWD

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_REV_TWD

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_PRICE_TWD

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        '==================================================
        '********************* 短期借款 *******************
        '==================================================

        sqlStr = "select GQ22_ITEM_NM, GQ26_ITEM_NM, GQ26_LEND, sum(nvl(GQ26_TWD,0))"
        sqlStr = sqlStr & ",GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " from GQ26, GQ22"
        sqlStr = sqlStr & " where GQ26_ITEM_CD = GQ22_ITEM_CD"
        sqlStr = sqlStr & "   and GQ22_DATA_CD ='F'"

        'If GQXX_COMP_CD <> "BLANK" Then
        '    sqlStr = sqlStr & " AND GQ26_COMP_CD = '" & GQXX_COMP_CD & "'"
        'End If


        If GQXX_DATA_YY <> "" Then
            sqlStr = sqlStr & " AND GQ26_DATA_YY = '" & GQXX_DATA_YY & "'"
        End If

        If GQXX_SEASON_CD <> "BLANK" Then
            sqlStr = sqlStr & " AND GQ26_SEASON_CD = '" & GQXX_SEASON_CD & "'"
        End If
        sqlStr = sqlStr & " group by GQ22_ITEM_NM,GQ26_ITEM_NM,GQ26_LEND,GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " order by GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"

        ' 建立 TABLE 
        table_node = XMLDoc.createElement("TABLE")
        oRootNode.appendChild(table_node)
        table_node.setAttribute("name", "短期借款")
        'table_node.SetAttribute("border", "1")

        '***************** 建立 表頭
        thead1_node = XMLDoc.createElement("THEAD")
        table_node.appendChild(thead1_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = "短期借款"
        tr_node.appendChild(td_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)


        headstr = " 項目類別@1_項目名稱@1_借款性質@1_金額@1_"
        Titletext = (Replace(headstr, "@1", ""))
        title = Split(Titletext, "_")

        For i = LBound(title) To UBound(title)
            td_node = XMLDoc.createElement("TD")
            td_node.text = title(i)
            tr_node.appendChild(td_node)

        Next
        '***************** 建立 表身體
        'tbody_node = XMLDoc.createElement("TBODY")
        'table_node.appendChild(tbody_node)

        objDataReader = objDataservice.ExecuteReader(sqlStr)

        OLD_GROUP_SUB = ""
        TEMP_GROUP_SUB = ""
        OLD_GROUP = ""
        TEMP_GROUP = ""

        Dim S_AMT, S_RATE As Decimal
        Dim T_AMT, T_RATE As Decimal

        iCnt = 1

        Do While objDataReader.Read()

            '************ 小計 ******************
            If Not IsDBNull(objDataReader(4)) Then
                TEMP_GROUP_SUB = objDataReader(4)
            Else
                TEMP_GROUP_SUB = ""
            End If

            If Trim(OLD_GROUP_SUB) = "" Then
                If Not IsDBNull(objDataReader(4)) Then
                    OLD_GROUP_SUB = objDataReader(4)
                Else
                    OLD_GROUP_SUB = ""
                End If
            End If

            If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "小計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_AMT


                '小計歸零
                S_AMT = 0

                OLD_GROUP_SUB = TEMP_GROUP_SUB
            End If

            '************ 合計 ******************
            If Not IsDBNull(objDataReader(5)) Then
                TEMP_GROUP = objDataReader(5)
            Else
                TEMP_GROUP = ""
            End If

            If Trim(OLD_GROUP) = "" Then
                If Not IsDBNull(objDataReader(5)) Then
                    OLD_GROUP = objDataReader(5)
                Else
                    OLD_GROUP = ""
                End If
            End If

            If TEMP_GROUP <> OLD_GROUP Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "合計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_AMT


                '合計歸零
                T_AMT = 0

                OLD_GROUP = TEMP_GROUP
            End If


            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            For i = 0 To objDataReader.FieldCount - 1

                Select Case i
                    Case 0
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            TEMP_ITEM_NM = ""
                        Else
                            TEMP_ITEM_NM = objDataReader(i)
                        End If

                        If OLD_ITEM_NM = "" And iCnt = 1 Then
                            If Not IsDBNull(objDataReader(i)) Then
                                OLD_ITEM_NM = objDataReader(i)
                                td_node.text = objDataReader(i)
                            Else
                                OLD_ITEM_NM = ""
                                td_node.text = ""
                            End If
                        End If

                        If OLD_ITEM_NM <> TEMP_ITEM_NM Then
                            OLD_ITEM_NM = TEMP_ITEM_NM
                            td_node.text = OLD_ITEM_NM
                        End If

                    Case 3
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_AMT = S_AMT + 0
                            T_AMT = T_AMT + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_AMT = S_AMT + objDataReader(i)
                            End If
                            T_AMT = T_AMT + objDataReader(i)
                        End If

                        'Case 4
                        '    td_node = XMLDoc.createElement("TD")
                        '    tr_node.appendChild(td_node)
                        '    If IsDBNull(objDataReader(i)) Then
                        '        td_node.text = ""
                        '        S_RATE = S_RATE + 0
                        '        T_RATE = T_RATE + 0
                        '    Else
                        '        td_node.text = objDataReader(i)

                        '        If OLD_GROUP_SUB <> "" Then
                        '            S_RATE = S_RATE + objDataReader(i)
                        '        End If
                        '        T_RATE = T_RATE + objDataReader(i)
                        '    End If

                    Case 4, 5
                    Case Else
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                        Else
                            td_node.text = objDataReader(i)
                        End If

                End Select
            Next
            iCnt = iCnt + 1
        Loop

        objDataReader.Close()

        If OLD_GROUP_SUB <> "" Then
            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = "小計"

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_AMT

        End If



        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = "合計"

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_AMT




        '==================================================
        '******************* 應付短期票券 *****************
        '==================================================

        sqlStr = "select GQ22_ITEM_NM, GQ27_ITEM_NM, GQ27_INSTITUT, sum(nvl(GQ27_TWD,0)),sum(nvl(GQ27_DAY,0))"
        sqlStr = sqlStr & ",GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " from GQ27, GQ22"
        sqlStr = sqlStr & " where GQ27_ITEM_CD = GQ22_ITEM_CD"
        sqlStr = sqlStr & "   and GQ22_DATA_CD ='G'"

        'If GQXX_COMP_CD <> "BLANK" Then
        '    sqlStr = sqlStr & " AND GQ27_COMP_CD = '" & GQXX_COMP_CD & "'"
        'End If


        If GQXX_DATA_YY <> "" Then
            sqlStr = sqlStr & " AND GQ27_DATA_YY = '" & GQXX_DATA_YY & "'"
        End If

        If GQXX_SEASON_CD <> "BLANK" Then
            sqlStr = sqlStr & " AND GQ27_SEASON_CD = '" & GQXX_SEASON_CD & "'"
        End If

        sqlStr = sqlStr & " group by GQ22_ITEM_NM,GQ27_ITEM_NM,GQ27_INSTITUT,GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " order by GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"

        ' 建立 TABLE 
        table_node = XMLDoc.createElement("TABLE")
        oRootNode.appendChild(table_node)
        table_node.setAttribute("name", "應付短期票券")
        'table_node.SetAttribute("border", "1")

        '***************** 建立 表頭
        thead1_node = XMLDoc.createElement("THEAD")
        table_node.appendChild(thead1_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = "應付短期票券"
        tr_node.appendChild(td_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)


        headstr = " 項目類別@1_項目名稱@1_保證機構@1_金額@1_期間(天)@1_"
        Titletext = (Replace(headstr, "@1", ""))
        title = Split(Titletext, "_")

        For i = LBound(title) To UBound(title)
            td_node = XMLDoc.createElement("TD")
            td_node.text = title(i)
            tr_node.appendChild(td_node)

        Next
        '***************** 建立 表身體
        'tbody_node = XMLDoc.createElement("TBODY")
        'table_node.appendChild(tbody_node)

        objDataReader = objDataservice.ExecuteReader(sqlStr)
        OLD_GROUP_SUB = ""
        TEMP_GROUP_SUB = ""
        OLD_GROUP = ""
        TEMP_GROUP = ""

        Dim S_GQ27_AMT, S_GQ27_RATE, S_GQ27_DAY As Decimal
        Dim T_GQ27_AMT, T_GQ27_RATE, T_GQ27_DAY As Decimal

        iCnt = 1
        Do While objDataReader.Read()

            '************ 小計 ******************

            If Not IsDBNull(objDataReader(5)) Then
                TEMP_GROUP_SUB = objDataReader(5)
            Else
                TEMP_GROUP_SUB = ""
            End If

            If Trim(OLD_GROUP_SUB) = "" Then
                If Not IsDBNull(objDataReader(5)) Then
                    OLD_GROUP_SUB = objDataReader(5)
                Else
                    OLD_GROUP_SUB = ""
                End If
            End If


            If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "小計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_GQ27_AMT


                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                'td_node.text = S_GQ27_DAY

                '小計歸零
                S_GQ27_AMT = 0
                S_GQ27_RATE = 0
                S_GQ27_DAY = 0


                OLD_GROUP_SUB = TEMP_GROUP_SUB
            End If

            '************ 合計 ******************

            If Not IsDBNull(objDataReader(6)) Then
                TEMP_GROUP = objDataReader(6)
            Else
                TEMP_GROUP = ""
            End If

            If Trim(OLD_GROUP) = "" Then
                If Not IsDBNull(objDataReader(6)) Then
                    OLD_GROUP = objDataReader(6)
                Else
                    OLD_GROUP = ""
                End If
            End If


            If TEMP_GROUP <> OLD_GROUP Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "合計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_GQ27_AMT


                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                'td_node.text = T_GQ27_DAY

                '合計歸零
                T_GQ27_AMT = 0
                T_GQ27_RATE = 0
                T_GQ27_DAY = 0

                OLD_GROUP = TEMP_GROUP
            End If


            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            For i = 0 To objDataReader.FieldCount - 1

                Select Case i
                    Case 0
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            TEMP_ITEM_NM = ""
                        Else
                            TEMP_ITEM_NM = objDataReader(i)
                        End If

                        If OLD_ITEM_NM = "" And iCnt = 1 Then
                            If Not IsDBNull(objDataReader(i)) Then
                                OLD_ITEM_NM = objDataReader(i)
                                td_node.text = objDataReader(i)
                            Else
                                OLD_ITEM_NM = ""
                                td_node.text = ""
                            End If
                        End If

                        If OLD_ITEM_NM <> TEMP_ITEM_NM Then
                            OLD_ITEM_NM = TEMP_ITEM_NM
                            td_node.text = OLD_ITEM_NM
                        End If
                    Case 3
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_GQ27_AMT = S_GQ27_AMT + 0
                            T_GQ27_AMT = T_GQ27_AMT + 0
                        Else
                            td_node.text = objDataReader(i)
                            If OLD_GROUP_SUB <> "" Then
                                S_GQ27_AMT = S_GQ27_AMT + objDataReader(i)
                            End If
                            T_GQ27_AMT = T_GQ27_AMT + objDataReader(i)
                        End If

                    Case 4
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_GQ27_DAY = S_GQ27_DAY + 0
                            T_GQ27_DAY = T_GQ27_DAY + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_GQ27_DAY = S_GQ27_DAY + objDataReader(i)
                            End If
                            T_GQ27_DAY = T_GQ27_DAY + objDataReader(i)
                        End If

                    Case 5, 6
                    Case Else
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                        Else
                            td_node.text = objDataReader(i)
                        End If

                End Select
            Next
            iCnt = iCnt + 1
        Loop

        objDataReader.Close()

        If OLD_GROUP_SUB <> "" Then
            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = "小計"

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_GQ27_AMT

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            'td_node.text = S_GQ27_DAY

        End If


        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = "合計"

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_GQ27_AMT


        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        'td_node.text = T_GQ27_DAY


        '==================================================
        '********************* 長期借款 *******************
        '==================================================

        sqlStr = "select GQ22_ITEM_NM, GQ28_ITEM_NM, GQ28_LEND, sum(nvl(GQ28_ORGLEND_TWD,0)),sum(nvl(GQ28_LEND_TWD,0))"
        sqlStr = sqlStr & ",GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " from GQ28, GQ22"
        sqlStr = sqlStr & " where GQ28_ITEM_CD = GQ22_ITEM_CD"
        sqlStr = sqlStr & "   and GQ22_DATA_CD ='H'"

        'If GQXX_COMP_CD <> "BLANK" Then
        '    sqlStr = sqlStr & " AND GQ28_COMP_CD = '" & GQXX_COMP_CD & "'"
        'End If


        If GQXX_DATA_YY <> "" Then
            sqlStr = sqlStr & " AND GQ28_DATA_YY = '" & GQXX_DATA_YY & "'"
        End If

        If GQXX_SEASON_CD <> "BLANK" Then
            sqlStr = sqlStr & " AND GQ28_SEASON_CD = '" & GQXX_SEASON_CD & "'"
        End If

        sqlStr = sqlStr & " group by GQ22_ITEM_NM,GQ28_ITEM_NM,GQ28_LEND,GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " order by GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"

        ' 建立 TABLE 
        table_node = XMLDoc.createElement("TABLE")
        oRootNode.appendChild(table_node)
        table_node.setAttribute("name", "長期借款")
        'table_node.SetAttribute("border", "1")

        '***************** 建立 表頭
        thead1_node = XMLDoc.createElement("THEAD")
        table_node.appendChild(thead1_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = "長期借款"
        tr_node.appendChild(td_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)


        headstr = " 項目類別@1_項目名稱@1_借款性質@1_原借款金額@1_金額@1_"
        Titletext = (Replace(headstr, "@1", ""))
        title = Split(Titletext, "_")

        For i = LBound(title) To UBound(title)
            td_node = XMLDoc.createElement("TD")
            td_node.text = title(i)
            tr_node.appendChild(td_node)

        Next
        '***************** 建立 表身體
        'tbody_node = XMLDoc.createElement("TBODY")
        'table_node.appendChild(tbody_node)

        objDataReader = objDataservice.ExecuteReader(sqlStr)

        OLD_GROUP_SUB = ""
        TEMP_GROUP_SUB = ""
        OLD_GROUP = ""
        TEMP_GROUP = ""

        Dim S_GQ28_ORGLEND_AMT, S_GQ28_LEND_AMT, S_GQ28_RATE As Decimal
        Dim T_GQ28_ORGLEND_AMT, T_GQ28_LEND_AMT, T_GQ28_RATE As Decimal

        iCnt = 1
        Do While objDataReader.Read()
            '************ 小計 ******************

            If Not IsDBNull(objDataReader(5)) Then
                TEMP_GROUP_SUB = objDataReader(5)
            Else
                TEMP_GROUP_SUB = ""
            End If

            If Trim(OLD_GROUP_SUB) = "" Then
                If Not IsDBNull(objDataReader(5)) Then
                    OLD_GROUP_SUB = objDataReader(5)
                Else
                    OLD_GROUP_SUB = ""
                End If
            End If

            If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "小計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_GQ28_ORGLEND_AMT

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_GQ28_LEND_AMT


                '小計歸零
                S_GQ28_ORGLEND_AMT = 0
                S_GQ28_LEND_AMT = 0

                OLD_GROUP_SUB = TEMP_GROUP_SUB
            End If

            '************ 合計 ******************

            If Not IsDBNull(objDataReader(6)) Then
                TEMP_GROUP = objDataReader(6)
            Else
                TEMP_GROUP = ""
            End If

            If Trim(OLD_GROUP) = "" Then
                If Not IsDBNull(objDataReader(6)) Then
                    OLD_GROUP = objDataReader(6)
                Else
                    OLD_GROUP = ""
                End If
            End If


            If TEMP_GROUP <> OLD_GROUP Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "合計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_GQ28_ORGLEND_AMT

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_GQ28_LEND_AMT

                '合計歸零
                T_GQ28_ORGLEND_AMT = 0
                T_GQ28_LEND_AMT = 0


                OLD_GROUP = TEMP_GROUP
            End If


            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            For i = 0 To objDataReader.FieldCount - 1

                Select Case i
                    Case 0
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            TEMP_ITEM_NM = ""
                        Else
                            TEMP_ITEM_NM = objDataReader(i)
                        End If

                        If OLD_ITEM_NM = "" And iCnt = 1 Then
                            If Not IsDBNull(objDataReader(i)) Then
                                OLD_ITEM_NM = objDataReader(i)
                                td_node.text = objDataReader(i)
                            Else
                                OLD_ITEM_NM = ""
                                td_node.text = ""
                            End If
                        End If

                        If OLD_ITEM_NM <> TEMP_ITEM_NM Then
                            OLD_ITEM_NM = TEMP_ITEM_NM
                            td_node.text = OLD_ITEM_NM
                        End If
                    Case 3
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_GQ28_ORGLEND_AMT = S_GQ28_ORGLEND_AMT + 0
                            T_GQ28_ORGLEND_AMT = T_GQ28_ORGLEND_AMT + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_GQ28_ORGLEND_AMT = S_GQ28_ORGLEND_AMT + objDataReader(i)
                            End If
                            T_GQ28_ORGLEND_AMT = T_GQ28_ORGLEND_AMT + objDataReader(i)

                        End If

                    Case 4
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_GQ28_LEND_AMT = S_GQ28_LEND_AMT + 0
                            T_GQ28_LEND_AMT = T_GQ28_LEND_AMT + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_GQ28_LEND_AMT = S_GQ28_LEND_AMT + objDataReader(i)
                            End If
                            T_GQ28_LEND_AMT = T_GQ28_LEND_AMT + objDataReader(i)
                        End If

                    Case 5, 6

                    Case Else
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                        Else
                            td_node.text = objDataReader(i)
                        End If

                End Select
            Next

            iCnt = iCnt + 1
        Loop

        objDataReader.Close()

        If OLD_GROUP_SUB <> "" Then
            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = "小計"

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_GQ28_ORGLEND_AMT

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_GQ28_LEND_AMT

        End If

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = "合計"

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_GQ28_ORGLEND_AMT

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_GQ28_LEND_AMT



        '==================================================
        '********************* 質押之資產 *****************
        '==================================================

        sqlStr = "select  GQ29_ITEM_NM, GQ07_ACCT_NM_CF, GQ01_NM_CA, sum(nvl(GQ29_TWD,0))"
        sqlStr = sqlStr & ",GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " from GQ29, GQ07, GQ22, GQ01"
        sqlStr = sqlStr & " where GQ29_ACCT_GAH = GQ07_ACCT_GAH(+)"
        sqlStr = sqlStr & " and GQ29_ACCT_SUBAH = GQ07_ACCT_SUBAH(+)"
        sqlStr = sqlStr & " and GQ29_ACCT_SUBBH = GQ07_ACCT_SUBBH(+)"
        sqlStr = sqlStr & " and GQ29_ITEM_CD = GQ22_ITEM_CD(+)"
        sqlStr = sqlStr & " and GQ29_COMP_CD = GQ01_COMP_CD(+)"
        'sqlStr = sqlStr & "   and GQ22_DATA_CD ='I'"

        'If GQXX_COMP_CD <> "BLANK" Then
        '    sqlStr = sqlStr & " AND GQ29_COMP_CD = '" & GQXX_COMP_CD & "'"
        'End If


        If GQXX_DATA_YY <> "" Then
            sqlStr = sqlStr & " AND GQ29_DATA_YY = '" & GQXX_DATA_YY & "'"
        End If

        If GQXX_SEASON_CD <> "BLANK" Then
            sqlStr = sqlStr & " AND GQ29_SEASON_CD = '" & GQXX_SEASON_CD & "'"
        End If

        sqlStr = sqlStr & " group by GQ29_ITEM_NM,GQ07_ACCT_NM_CF,GQ01_NM_CA,GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " order by GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP,GQ01_NM_CA"

        ' 建立 TABLE 
        table_node = XMLDoc.createElement("TABLE")
        oRootNode.appendChild(table_node)
        table_node.setAttribute("name", "質押之資產")
        'table_node.SetAttribute("border", "1")

        '***************** 建立 表頭
        thead1_node = XMLDoc.createElement("THEAD")
        table_node.appendChild(thead1_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = "質押之資產"
        tr_node.appendChild(td_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)


        headstr = " 項目類別@1_帳列科目@1_投資公司@1_金額@1_"
        Titletext = (Replace(headstr, "@1", ""))
        title = Split(Titletext, "_")

        For i = LBound(title) To UBound(title)
            td_node = XMLDoc.createElement("TD")
            td_node.text = title(i)
            tr_node.appendChild(td_node)

        Next
        '***************** 建立 表身體
        'tbody_node = XMLDoc.createElement("TBODY")
        'table_node.appendChild(tbody_node)

        objDataReader = objDataservice.ExecuteReader(sqlStr)

        OLD_GROUP_SUB = ""
        TEMP_GROUP_SUB = ""
        OLD_GROUP = ""
        TEMP_GROUP = ""

        Dim S_GQ29_AMT As Decimal
        Dim T_GQ29_AMT As Decimal


        iCnt = 1
        Do While objDataReader.Read()

            '************ 小計 ******************
            If Not IsDBNull(objDataReader(4)) Then
                TEMP_GROUP_SUB = objDataReader(4)
            Else
                TEMP_GROUP_SUB = ""
            End If

            If Trim(OLD_GROUP_SUB) = "" Then
                If Not IsDBNull(objDataReader(4)) Then
                    OLD_GROUP_SUB = objDataReader(4)
                Else
                    OLD_GROUP_SUB = ""
                End If
            End If

            If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "小計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_GQ29_AMT

                '小計歸零
                S_GQ29_AMT = 0

                OLD_GROUP_SUB = TEMP_GROUP_SUB
            End If

            '************ 合計 ******************
            If Not IsDBNull(objDataReader(5)) Then
                TEMP_GROUP = objDataReader(5)
            Else
                TEMP_GROUP = ""
            End If

            If Trim(OLD_GROUP) = "" Then
                If Not IsDBNull(objDataReader(5)) Then
                    OLD_GROUP = objDataReader(5)
                Else
                    OLD_GROUP = ""
                End If
            End If

            If TEMP_GROUP <> OLD_GROUP Then
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = "合計"

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = T_GQ29_AMT

                '合計歸零
                T_GQ29_AMT = 0


                OLD_GROUP = TEMP_GROUP
            End If


            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            For i = 0 To objDataReader.FieldCount - 1

                Select Case i
                    Case 0
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            TEMP_ITEM_NM = ""
                        Else
                            TEMP_ITEM_NM = objDataReader(i)
                        End If

                        If OLD_ITEM_NM = "" And iCnt = 1 Then
                            If Not IsDBNull(objDataReader(i)) Then
                                OLD_ITEM_NM = objDataReader(i)
                                td_node.text = objDataReader(i)
                            Else
                                OLD_ITEM_NM = ""
                                td_node.text = ""
                            End If
                        End If

                        If OLD_ITEM_NM <> TEMP_ITEM_NM Then
                            OLD_ITEM_NM = TEMP_ITEM_NM
                            td_node.text = OLD_ITEM_NM
                        End If
                    Case 3

                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                            S_GQ29_AMT = S_GQ29_AMT + 0
                            T_GQ29_AMT = T_GQ29_AMT + 0
                        Else
                            td_node.text = objDataReader(i)

                            If OLD_GROUP_SUB <> "" Then
                                S_GQ29_AMT = S_GQ29_AMT + objDataReader(i)
                            End If
                            T_GQ29_AMT = T_GQ29_AMT + objDataReader(i)
                        End If

                    Case 4, 5
                    Case Else
                        td_node = XMLDoc.createElement("TD")
                        tr_node.appendChild(td_node)
                        If IsDBNull(objDataReader(i)) Then
                            td_node.text = ""
                        Else
                            td_node.text = objDataReader(i)
                        End If

                End Select
            Next
            iCnt = iCnt + 1
        Loop

        objDataReader.Close()

        If OLD_GROUP_SUB <> "" Then
            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = "小計"

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = S_GQ29_AMT

        End If


        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = "合計"

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = ""

        td_node = XMLDoc.createElement("TD")
        tr_node.appendChild(td_node)
        td_node.text = T_GQ29_AMT




        '==================================================
        '******** 用人、折舊及攤銷費用之功能別彙總 ********
        '==================================================

        'sqlStr = "select decode(nvl(GQ30_ITEM_NM,'0'),'0',GQ22_ITEM_NM,decode(rtrim(ltrim(GQ30_ITEM_NM)),'',GQ22_ITEM_NM,GQ30_ITEM_NM)) as GQ30_ITEM_NM, GQ30_COST_TWD, GQ30_EXP_TWD"
        sqlStr = ""
        'sqlStr = "select GQ30_ITEM_NM, GQ01_NM_CA, sum(nvl(GQ30_COST_TWD,0)), sum(nvl(GQ30_EXP_TWD,0)) ,GQ22_GROUP_SUB,GQ22_GROUP"
        sqlStr = sqlStr & " from GQ30, GQ22, GQ01"
        sqlStr = sqlStr & " where GQ30_ITEM_CD = GQ22_ITEM_CD"
        sqlStr = sqlStr & "   and GQ30_COMP_CD = GQ01_COMP_CD"
        sqlStr = sqlStr & "   and GQ22_DATA_CD ='J'"

        If GQXX_DATA_YY <> "" Then
            sqlStr = sqlStr & " AND GQ30_DATA_YY = '" & GQXX_DATA_YY & "'"
        End If
        If GQXX_SEASON_CD <> "BLANK" Then
            sqlStr = sqlStr & " AND GQ30_SEASON_CD = '" & GQXX_SEASON_CD & "'"
        End If

        Dim sCompList
        sCompList = "select GQ01_COMP_CD , GQ01_NM_CA " & sqlStr
        sCompList = sCompList & " group by GQ01_COMP_CD,GQ01_NM_CA"
        sCompList = sCompList & " order by GQ01_COMP_CD"
        objDataReader = objDataservice.ExecuteReader(sCompList)

        Dim oCompLists As New ArrayList
        Do While objDataReader.Read()
            Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0}
            oCompLists.Add(oColumn)
        Loop
      
        ' 建立 TABLE 
        table_node = XMLDoc.createElement("TABLE")
        oRootNode.appendChild(table_node)
        table_node.setAttribute("name", "用人、折舊及攤銷費用之功能別彙總")
        'table_node.SetAttribute("border", "1")

        '***************** 建立 表頭
        thead1_node = XMLDoc.createElement("THEAD")
        table_node.appendChild(thead1_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = "用人、折舊及攤銷費用之功能別彙總"
        tr_node.appendChild(td_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)

        td_node = XMLDoc.createElement("TD")
        td_node.text = "項目類別"
        'td_node.setAttribute("rowspan", "2")
        tr_node.appendChild(td_node)

        For i = 0 To oCompLists.Count - 1
            td_node = XMLDoc.createElement("TD")
            td_node.text = oCompLists(i)(1)
            td_node.setAttribute("colspan", "2")
            tr_node.appendChild(td_node)
        Next

        td_node = XMLDoc.createElement("TD")
        td_node.text = "合計"
        td_node.setAttribute("colspan", "2")
        tr_node.appendChild(td_node)

        tr_node = XMLDoc.createElement("TR")
        thead1_node.appendChild(tr_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = ""
        tr_node.appendChild(td_node)

        For i = 0 To oCompLists.Count
            td_node = XMLDoc.createElement("TD")
            td_node.text = "關係企業營業成本"
            tr_node.appendChild(td_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "關係企業營業費用"
            tr_node.appendChild(td_node)
        Next

        sqlStr = "select GQ30_ITEM_NM, GQ01_NM_CA, sum(nvl(GQ30_COST_TWD,0)), sum(nvl(GQ30_EXP_TWD,0)) ,GQ22_GROUP_SUB,GQ22_GROUP " & sqlStr
        sqlStr = sqlStr & " group by GQ30_ITEM_NM,GQ01_NM_CA,GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ22_GROUP,GQ01_COMP_CD"
        sqlStr = sqlStr & " order by GQ22_GROUP_SEQ,GQ22_GROUP_SUB,GQ01_COMP_CD,GQ22_GROUP"
        objDataReader = objDataservice.ExecuteReader(sqlStr)

        OLD_GROUP_SUB = ""
        TEMP_GROUP_SUB = ""
        OLD_GROUP = ""
        TEMP_GROUP = ""

        Dim S_GQ30_COST_AMT, S_GQ30_EXP_AMT As Decimal
        Dim T_GQ30_COST_AMT, T_GQ30_EXP_AMT As Decimal
        Dim iCompCnt
        iCompCnt = 0
        iCnt = 1
        Do While objDataReader.Read()
            TEMP_GROUP = objDataReader(0)
            If OLD_GROUP <> TEMP_GROUP Then

                If OLD_GROUP <> "" Then
                    For i = 1 To oCompLists.Count - iCompCnt
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = "0"
                        tr_node.appendChild(td_node)
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = "0"
                        tr_node.appendChild(td_node)
                    Next

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = T_GQ30_COST_AMT
                    tr_node.appendChild(td_node)
                    td_node = XMLDoc.createElement("TD")
                    td_node.text = T_GQ30_EXP_AMT
                    tr_node.appendChild(td_node)
                    T_GQ30_COST_AMT = 0
                    T_GQ30_EXP_AMT = 0
                End If


                OLD_GROUP = TEMP_GROUP
                tr_node = XMLDoc.createElement("TR")
                thead1_node.appendChild(tr_node)

                td_node = XMLDoc.createElement("TD")
                td_node.text = objDataReader(0)
                tr_node.appendChild(td_node)
                iCompCnt = 0
            End If

            For i = iCompCnt To oCompLists.Count - 1
                iCompCnt = iCompCnt + 1
                If oCompLists(i)(1) = objDataReader(1) Then
                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(2)
                    tr_node.appendChild(td_node)
                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(3)
                    tr_node.appendChild(td_node)
                    oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(2)
                    oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3)
                    T_GQ30_COST_AMT = T_GQ30_COST_AMT + objDataReader(2)
                    T_GQ30_EXP_AMT = T_GQ30_EXP_AMT + objDataReader(3)
                    Exit For
                Else
                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "0"
                    tr_node.appendChild(td_node)
                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "0"
                    tr_node.appendChild(td_node)
                End If
            Next
        Loop

        For i = 1 To oCompLists.Count - iCompCnt
            td_node = XMLDoc.createElement("TD")
            td_node.text = "0"
            tr_node.appendChild(td_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "0"
            tr_node.appendChild(td_node)
        Next
        td_node = XMLDoc.createElement("TD")
        td_node.text = T_GQ30_COST_AMT
        tr_node.appendChild(td_node)
        td_node = XMLDoc.createElement("TD")
        td_node.text = T_GQ30_EXP_AMT
        tr_node.appendChild(td_node)
       
        'tr_node = XMLDoc.createElement("TR")
        'thead1_node.appendChild(tr_node)
        'td_node = XMLDoc.createElement("TD")
        'td_node.text = "合計"
        'tr_node.appendChild(td_node)

        'For i = 0 To oCompLists.Count - 1
        '    td_node = XMLDoc.createElement("TD")
        '    td_node.text = oCompLists(i)(2)
        '    tr_node.appendChild(td_node)
        '    td_node = XMLDoc.createElement("TD")
        '    td_node.text = oCompLists(i)(3)
        '    tr_node.appendChild(td_node)
        'Next

        'tr_node = XMLDoc.createElement("TR")
        'thead1_node.appendChild(tr_node)


        'headstr = " 項目類別@1_投資公司@1_各關係企業營業成本@1_各關係企業營業費用@1_"
        'Titletext = (Replace(headstr, "@1", ""))
        'title = Split(Titletext, "_")

        'For i = LBound(title) To UBound(title)
        '    td_node = XMLDoc.createElement("TD")
        '    td_node.text = title(i)
        '    tr_node.appendChild(td_node)

        'Next
        ''***************** 建立 表身體
        ''tbody_node = XMLDoc.createElement("TBODY")
        ''table_node.appendChild(tbody_node)

        'objDataReader = objDataservice.ExecuteReader(sqlStr)

        'OLD_GROUP_SUB = ""
        'TEMP_GROUP_SUB = ""
        'OLD_GROUP = ""
        'TEMP_GROUP = ""

        'Dim S_GQ30_COST_AMT, S_GQ30_EXP_AMT As Decimal
        'Dim T_GQ30_COST_AMT, T_GQ30_EXP_AMT As Decimal

        'iCnt = 1
        'Do While objDataReader.Read()

        '    '************ 小計 ******************
        '    If Not IsDBNull(objDataReader(4)) Then
        '        TEMP_GROUP_SUB = objDataReader(4)
        '    Else
        '        TEMP_GROUP_SUB = ""
        '    End If

        '    If Trim(OLD_GROUP_SUB) = "" Then
        '        If Not IsDBNull(objDataReader(4)) Then
        '            OLD_GROUP_SUB = objDataReader(4)
        '        Else
        '            OLD_GROUP_SUB = ""
        '        End If
        '    End If

        '    If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
        '        tr_node = XMLDoc.createElement("TR")
        '        thead1_node.appendChild(tr_node)

        '        td_node = XMLDoc.createElement("TD")
        '        tr_node.appendChild(td_node)
        '        td_node.text = "小計"

        '        td_node = XMLDoc.createElement("TD")
        '        tr_node.appendChild(td_node)
        '        td_node.text = ""

        '        td_node = XMLDoc.createElement("TD")
        '        tr_node.appendChild(td_node)
        '        td_node.text = S_GQ30_COST_AMT

        '        td_node = XMLDoc.createElement("TD")
        '        tr_node.appendChild(td_node)
        '        td_node.text = S_GQ30_EXP_AMT

        '        '小計歸零
        '        S_GQ30_COST_AMT = 0
        '        S_GQ30_EXP_AMT = 0


        '        OLD_GROUP_SUB = TEMP_GROUP_SUB
        '    End If

        '    '************ 合計 ******************

        '    If Not IsDBNull(objDataReader(5)) Then
        '        TEMP_GROUP = objDataReader(5)
        '    Else
        '        TEMP_GROUP = ""
        '    End If

        '    If Trim(OLD_GROUP) = "" Then
        '        If Not IsDBNull(objDataReader(5)) Then
        '            OLD_GROUP = objDataReader(5)
        '        Else
        '            OLD_GROUP = ""
        '        End If
        '    End If


        '    If TEMP_GROUP <> OLD_GROUP Then
        '        tr_node = XMLDoc.createElement("TR")
        '        thead1_node.appendChild(tr_node)

        '        td_node = XMLDoc.createElement("TD")
        '        tr_node.appendChild(td_node)
        '        td_node.text = "合計"

        '        td_node = XMLDoc.createElement("TD")
        '        tr_node.appendChild(td_node)
        '        td_node.text = ""

        '        td_node = XMLDoc.createElement("TD")
        '        tr_node.appendChild(td_node)
        '        td_node.text = T_GQ30_COST_AMT

        '        td_node = XMLDoc.createElement("TD")
        '        tr_node.appendChild(td_node)
        '        td_node.text = T_GQ30_EXP_AMT

        '        '小計歸零
        '        T_GQ30_COST_AMT = 0
        '        T_GQ30_EXP_AMT = 0

        '        OLD_GROUP = TEMP_GROUP
        '    End If


        '    tr_node = XMLDoc.createElement("TR")
        '    thead1_node.appendChild(tr_node)

        '    For i = 0 To objDataReader.FieldCount - 1

        '        Select Case i
        '            Case 0
        '                td_node = XMLDoc.createElement("TD")
        '                tr_node.appendChild(td_node)
        '                If IsDBNull(objDataReader(i)) Then
        '                    TEMP_ITEM_NM = ""
        '                Else
        '                    TEMP_ITEM_NM = objDataReader(i)
        '                End If

        '                If OLD_ITEM_NM = "" And iCnt = 1 Then
        '                    If Not IsDBNull(objDataReader(i)) Then
        '                        OLD_ITEM_NM = objDataReader(i)
        '                        td_node.text = objDataReader(i)
        '                    Else
        '                        OLD_ITEM_NM = ""
        '                        td_node.text = ""
        '                    End If
        '                End If

        '                If OLD_ITEM_NM <> TEMP_ITEM_NM Then
        '                    OLD_ITEM_NM = TEMP_ITEM_NM
        '                    td_node.text = OLD_ITEM_NM
        '                End If
        '            Case 2
        '                td_node = XMLDoc.createElement("TD")
        '                tr_node.appendChild(td_node)
        '                If IsDBNull(objDataReader(i)) Then
        '                    td_node.text = ""
        '                    S_GQ30_COST_AMT = S_GQ30_COST_AMT + 0
        '                    T_GQ30_COST_AMT = T_GQ30_COST_AMT + 0
        '                Else
        '                    td_node.text = objDataReader(i)
        '                    If OLD_GROUP_SUB <> "" Then
        '                        S_GQ30_COST_AMT = S_GQ30_COST_AMT + objDataReader(i)
        '                    End If
        '                    T_GQ30_COST_AMT = T_GQ30_COST_AMT + objDataReader(i)

        '                End If
        '            Case 3
        '                td_node = XMLDoc.createElement("TD")
        '                tr_node.appendChild(td_node)
        '                If IsDBNull(objDataReader(i)) Then
        '                    td_node.text = ""
        '                    S_GQ30_EXP_AMT = S_GQ30_EXP_AMT + 0
        '                    T_GQ30_EXP_AMT = T_GQ30_EXP_AMT + 0
        '                Else
        '                    td_node.text = objDataReader(i)

        '                    If OLD_GROUP_SUB <> "" Then
        '                        S_GQ30_EXP_AMT = S_GQ30_EXP_AMT + objDataReader(i)
        '                    End If
        '                    T_GQ30_EXP_AMT = T_GQ30_EXP_AMT + objDataReader(i)

        '                End If

        '            Case 4, 5

        '            Case Else
        '                td_node = XMLDoc.createElement("TD")
        '                tr_node.appendChild(td_node)
        '                If IsDBNull(objDataReader(i)) Then
        '                    td_node.text = ""
        '                Else
        '                    td_node.text = objDataReader(i)
        '                End If

        '        End Select
        '    Next
        '    iCnt = iCnt + 1
        'Loop

        'objDataReader.Close()

        'If OLD_GROUP_SUB <> "" Then
        '    tr_node = XMLDoc.createElement("TR")
        '    thead1_node.appendChild(tr_node)

        '    td_node = XMLDoc.createElement("TD")
        '    tr_node.appendChild(td_node)
        '    td_node.text = "小計"

        '    td_node = XMLDoc.createElement("TD")
        '    tr_node.appendChild(td_node)
        '    td_node.text = ""

        '    td_node = XMLDoc.createElement("TD")
        '    tr_node.appendChild(td_node)
        '    td_node.text = S_GQ30_COST_AMT

        '    td_node = XMLDoc.createElement("TD")
        '    tr_node.appendChild(td_node)
        '    td_node.text = S_GQ30_EXP_AMT
        'End If


        'tr_node = XMLDoc.createElement("TR")
        'thead1_node.appendChild(tr_node)

        'td_node = XMLDoc.createElement("TD")
        'tr_node.appendChild(td_node)
        'td_node.text = "合計"

        'td_node = XMLDoc.createElement("TD")
        'tr_node.appendChild(td_node)
        'td_node.text = ""

        'td_node = XMLDoc.createElement("TD")
        'tr_node.appendChild(td_node)
        'td_node.text = T_GQ30_COST_AMT

        'td_node = XMLDoc.createElement("TD")
        'tr_node.appendChild(td_node)
        'td_node.text = T_GQ30_EXP_AMT



        xmlData = Trim(XMLDoc.childNodes(0).xml)
    End Sub
    Protected Sub oSimpleGrid_OnSimpleGridShowData(ByVal XMLDoc As System.Xml.XmlDocument, ByVal DataNode As System.Xml.XmlElement, ByVal Column As System.Data.DataColumn, ByVal GridRow As System.Data.DataRow, ByRef Changed As Boolean) Handles oSimpleGrid.OnSimpleGridShowData

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' 檢查是否有傳 XMLDATA
        If CheckXMLDATA() Then Call Main()
    End Sub

End Class
