#Region "程式描述及修改歷程"

'+---------+---------------------------------------------------------------------------------------------------------------------------------
'|                        MODIFY  DESCRIPTION                         
'+---------+-----------+---------------------------------------------------------------------------------------------------------------------
'|  DATE   |  AUTHOR   |                    REASON                    
'+---------+-----------+---------------------------------------------------------------------------------------------------------------------
'|05APR2017| Robin Lin | 初版:IFRS合併報表系統-變動表彙總                                             
'+---------+-----------+---------------------------------------------------------------------------------------------------------------------
'|05APR2017| Robin Lin | 需求單號:060217
'|         |           | 調整說明:G 變動表-應付短期票券：增加顯示利率及關係企業資料,並調整為相同細項不合併                                          
'+---------+-----------+---------------------------------------------------------------------------------------------------------------------
'|19MAY2017| Kathy Hsu | 需求單號:060268
'|         |           | 調整說明:Z 變動表-保證函：增加顯示變動表-保證函                                      
'+---------+-----------+---------------------------------------------------------------------------------------------------------------------
'|04MAR2019| Bruce Wang| 需求單號:080156
'|         |           | 調整說明:Q 變動表-使用權資產：增加顯示變動表-使用權資產
'+---------+-----------+---------------------------------------------------------------------------------------------------------------------
'|30DEC2019| Bruce Wang| 需求單號:081075
'|         |           | 調整說明:報表增加欄位 項目代碼
'+---------+-----------+---------------------------------------------------------------------------------------------------------------------

#End Region

Imports TODO
Imports System.Xml

Partial Class CH_GR_GR0590
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
        'Dim GRXX_COMP_CD As String = Trim(getField("GRXX_COMP_CD"))
        Dim GRXX_SEASON_CD As String = Trim(getField("GRXX_SEASON_CD"))


        errmsg = ""
        Dim objDataservice As TODO.clsDataService = New TODO.clsDataService(clsDataService.DsDBName.GR02)
        Dim objDataReader As OleDb.OleDbDataReader
        Dim sqlStr As String
        'Dim GR01_NM_CA As String
        Dim GR05_SEASON_NM As String


        ''*******讀取GR01,關係企業代碼檔
        'sqlStr = " select GR01_NM_CA from GR01 where GR01_COMP_CD = '" & GRXX_COMP_CD & "'"

        'objDataReader = objDataservice.ExecuteReader(sqlStr)

        'If objDataReader.Read() Then
        '    If Not IsDBNull(objDataReader(0)) Then
        '        GR01_NM_CA = objDataReader(0)
        '    Else
        '        GR01_NM_CA = ""
        '    End If
        'Else
        '    GR01_NM_CA = ""
        'End If
        'objDataReader.Close()


        '*******讀取GR05,季節月份代碼檔
        sqlStr = " select GR05_SEASON_NM from GR05 where GR05_SEASON_CD = '" & GRXX_SEASON_CD & "'"

        objDataReader = objDataservice.ExecuteReader(sqlStr)

        If objDataReader.Read() Then
            If Not IsDBNull(objDataReader(0)) Then
                GR05_SEASON_NM = objDataReader(0)
            Else
                GR05_SEASON_NM = ""
            End If
        Else
            GR05_SEASON_NM = ""
        End If
        objDataReader.Close()


        xmlData = "<Collection>"
        xmlData = xmlData & "<MSG></MSG>"
        'xmlData = xmlData & "<GR01_NM_CA>" & GR01_NM_CA & "</GR01_NM_CA>"
        xmlData = xmlData & "<GR05_SEASON_NM>" & GR05_SEASON_NM & "</GR05_SEASON_NM>"
        xmlData = xmlData & "</Collection>"
        objDataReader.Close()

    End Sub

    Sub QueryExcel()

        Dim GRXX_COMP_CD As String = Trim(getField("GRXX_COMP_CD"))
        Dim GRXX_DATA_YY As String = Trim(getField("GRXX_DATA_YY"))
        Dim GRXX_SEASON_CD As String = Trim(getField("GRXX_SEASON_CD"))
        Dim GR18_DATA_CD As String = Trim(getField("GR18_DATA_CD"))

        Dim i As Integer
        Dim title As Object
        Dim Titletext As String
        Dim sqlStr, sqlStr1 As String
        Dim oRootNode, table_node, tbody_node, tr_node, td_node, thead1_node
        Dim lb_group, lb_group_sub As Boolean

        Dim objDataservice As TODO.clsDataService = New TODO.clsDataService(clsDataService.DsDBName.GR02)
        Dim objDataReader As OleDb.OleDbDataReader
        '
        Dim OLD_GROUP_SUB, TEMP_GROUP_SUB As String
        Dim OLD_GROUP, TEMP_GROUP As String
        Dim OLD_ITEM_NM, TEMP_ITEM_NM As String
        'Dim S_START_TWD, S_ADD_TWD, S_INVT_TWD, S_LESS_TWD, S_DIFF_TWD, S_FINAL_TWD, S_UNIT_CNT, S_PRICE_TWD As Decimal
        'Dim T_START_TWD, T_ADD_TWD, T_INVT_TWD, T_LESS_TWD, T_DIFF_TWD, T_FINAL_TWD, T_UNIT_CNT, T_PRICE_TWD As Decimal
        Dim S_FINAL_TWD, S_UNIT_CNT, S_PRICE_TWD, S_REV_TWD, S_DIVIDEND_TWD As Decimal
        Dim T_FINAL_TWD, T_UNIT_CNT, T_PRICE_TWD, T_REV_TWD, T_DIVIDEND_TWD As Decimal
        Dim iCnt As Decimal

        oRootNode = XMLDoc.createElement("Collection")
        XMLDoc.documentElement = oRootNode



        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "Z" Then
            '==================================================
            '******************* Z 保証函 ***************
            '==================================================

            sqlStr = "select "
            sqlStr = sqlStr & " GR56_ITEM_NM"                                       '0
            sqlStr = sqlStr & ",GR56_BIDS_NM"                                       '1
            sqlStr = sqlStr & " , GR56_BANK_NM"                                     '2
            sqlStr = sqlStr & " , GR56_DURA"                                     '3    
            sqlStr = sqlStr & ",nvl(GR56_TWD,0) GR56_TWD"                           '4
            sqlStr = sqlStr & ",nvl(GR56_FEE,'') GR56_FEE"                          '5
            sqlStr = sqlStr & ",nvl(GR56_RATE,'') GR56_RATE"                        '6
            sqlStr = sqlStr & ", GR56_comp_cd ||' '||GR01_nm_ca AS GR01_nm_ca"      '7
            sqlStr = sqlStr & ", GR56_comp_cd ||' '||GR01_nm_ca AS GR01_nm_ca"      '8
            sqlStr = sqlStr & ", GR56_comp_cd ||' '||GR01_nm_ca AS GR01_nm_ca"      '9
            sqlStr = sqlStr & " from GR56, GR01"
            sqlStr = sqlStr & " where GR01_comp_cd   = GR56_comp_cd"


            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR56_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If

            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR56_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            sqlStr = sqlStr & " group by GR01_COMP_CD,GR56_ITEM_NM,GR56_BIDS_NM,GR56_BANK_NM , GR56_DURA,  GR56_COMP_CD,  GR56_TWD, GR56_FEE , GR56_RATE,GR01_nm_ca "
            sqlStr = sqlStr & " order by GR01_COMP_CD,GR56_ITEM_NM,GR56_BIDS_NM,GR56_BANK_NM , GR56_DURA"

            ' 建立 TABLE 
            table_node = XMLDoc.createElement("TABLE")
            oRootNode.appendChild(table_node)
            table_node.setAttribute("name", "保証函")
            'table_node.SetAttribute("border", "1")

            '***************** 建立 表頭
            thead1_node = XMLDoc.createElement("THEAD")
            table_node.appendChild(thead1_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "保證函"
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)


            headstr = " 單位名稱@1_標案性質@1_保証銀行@1_保証起迄@1_金額@1_手續費@1_手續費利率(%)@1_關係企業@1_"
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

            Dim S_GR56_TWD, S_GR56_RATE, S_GR56_FEE As Decimal
            Dim T_GR56_TWD, T_GR56_RATE, T_GR56_FEE As Decimal

            iCnt = 1
            Do While objDataReader.Read()


                '************ 合計 ******************

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


                'If TEMP_GROUP <> OLD_GROUP Then
                '    tr_node = XMLDoc.createElement("TR")
                '    thead1_node.appendChild(tr_node)

                '    td_node = XMLDoc.createElement("TD")
                '    tr_node.appendChild(td_node)
                '    td_node.text = ""

                '    td_node = XMLDoc.createElement("TD")
                '    tr_node.appendChild(td_node)
                '    td_node.text = ""

                '    td_node = XMLDoc.createElement("TD")
                '    tr_node.appendChild(td_node)
                '    td_node.text = ""

                '    td_node = XMLDoc.createElement("TD")
                '    tr_node.appendChild(td_node)
                '    td_node.text = ""


                '    td_node = XMLDoc.createElement("TD")
                '    tr_node.appendChild(td_node)
                '    td_node.text = ""

                '    td_node = XMLDoc.createElement("TD")
                '    tr_node.appendChild(td_node)
                '    td_node.text = ""

                '    td_node = XMLDoc.createElement("TD")
                '    tr_node.appendChild(td_node)
                '    td_node.text = ""

                '    td_node = XMLDoc.createElement("TD")
                '    tr_node.appendChild(td_node)
                '    td_node.text = ""


                '    OLD_GROUP = TEMP_GROUP
                'End If


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
                        Case 4
                            td_node = XMLDoc.createElement("TD")
                            tr_node.appendChild(td_node)
                            If IsDBNull(objDataReader(i)) Then
                                td_node.text = ""
                                S_GR56_TWD = S_GR56_TWD + 0
                                T_GR56_TWD = T_GR56_TWD + 0
                            Else
                                td_node.text = objDataReader(i)
                                If OLD_GROUP_SUB <> "" Then
                                    S_GR56_TWD = S_GR56_TWD + objDataReader(i)
                                End If
                                T_GR56_TWD = T_GR56_TWD + objDataReader(i)
                            End If

                        Case 5
                            td_node = XMLDoc.createElement("TD")
                            tr_node.appendChild(td_node)
                            If IsDBNull(objDataReader(i)) Then
                                td_node.text = ""
                                S_GR56_FEE = S_GR56_FEE + 0
                                T_GR56_FEE = T_GR56_FEE + 0
                            Else
                                td_node.text = objDataReader(i)

                                If OLD_GROUP_SUB <> "" Then
                                    S_GR56_FEE = S_GR56_FEE + objDataReader(i)
                                End If
                                T_GR56_FEE = T_GR56_FEE + objDataReader(i)
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


            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)


            td_node = XMLDoc.createElement("TD")
            td_node.text = "合計"
            td_node.setAttribute("colspan", "4")
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = T_GR56_TWD


            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = T_GR56_FEE

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""
        End If






        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "D" Then
            '==================================================
            '******************* D 金融資產 *******************
            '==================================================

            sqlStr = "select GR22_ITEM_CD,GR22_ITEM_NM, GR24_ITEM_NM "
            'sqlStr = sqlStr & ", GR24_START_TWD, GR24_ADD_TWD "
            'sqlStr = sqlStr & ",GR24_INVT_TWD,GR24_LESS_TWD ,GR24_DIFF_TWD"    
            sqlStr = sqlStr & ",sum(nvl(GR24_FINAL_TWD,0)) ,sum(nvl(GR24_DIVIDEND_TWD,0)) ,sum(nvl(GR24_UNIT_CNT,0)) ,sum(nvl(GR24_PRICE_TWD,0)) ,GR01_NM_CA "
            sqlStr = sqlStr & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR24, GR22, GR01"
            sqlStr = sqlStr & " where GR24_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR24_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='D'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR24_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR24_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr = sqlStr & " group by GR22_ITEM_CD,GR22_ITEM_NM,GR24_ITEM_NM,GR01_NM_CA,GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP,GR01_NM_CA"

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

            headstr = " 項目代碼@1_項目類別@1_被投資標的@1_期末投資金額@1_股利收入@1_單位數@1_市價@1_投資公司@1_"

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

            'Dim OLD_GROUP_SUB, TEMP_GROUP_SUB As String
            'Dim OLD_GROUP, TEMP_GROUP As String
            'Dim OLD_ITEM_NM, TEMP_ITEM_NM As String
            ''Dim S_START_TWD, S_ADD_TWD, S_INVT_TWD, S_LESS_TWD, S_DIFF_TWD, S_FINAL_TWD, S_UNIT_CNT, S_PRICE_TWD As Decimal
            ''Dim T_START_TWD, T_ADD_TWD, T_INVT_TWD, T_LESS_TWD, T_DIFF_TWD, T_FINAL_TWD, T_UNIT_CNT, T_PRICE_TWD As Decimal
            'Dim S_FINAL_TWD, S_UNIT_CNT, S_PRICE_TWD, S_REV_TWD, S_DIVIDEND_TWD As Decimal
            'Dim T_FINAL_TWD, T_UNIT_CNT, T_PRICE_TWD, T_REV_TWD, T_DIVIDEND_TWD As Decimal
            'Dim iCnt As Decimal

            iCnt = 1
            Do While objDataReader.Read()

                '小計
                If Not IsDBNull(objDataReader(8)) Then
                    TEMP_GROUP_SUB = objDataReader(8)
                Else
                    TEMP_GROUP_SUB = ""
                End If

                If Trim(OLD_GROUP_SUB) = "" Then
                    If Not IsDBNull(objDataReader(8)) Then
                        OLD_GROUP_SUB = objDataReader(8)
                    Else
                        OLD_GROUP_SUB = ""
                    End If
                End If



                If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    td_node.text = ""

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
                If Not IsDBNull(objDataReader(9)) Then
                    TEMP_GROUP = objDataReader(9)
                Else
                    TEMP_GROUP = ""
                End If

                If Trim(OLD_GROUP) = "" Then
                    If Not IsDBNull(objDataReader(9)) Then
                        OLD_GROUP = objDataReader(9)
                    Else
                        OLD_GROUP = ""
                    End If
                End If



                If OLD_GROUP <> TEMP_GROUP Then

                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    td_node.text = ""

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
                                td_node.text = ""
                            Else
                                td_node.text = "D" & objDataReader(i)
                            End If
                        Case 1
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
                                S_FINAL_TWD = S_FINAL_TWD + 0
                                T_FINAL_TWD = T_FINAL_TWD + 0
                            Else
                                td_node.text = objDataReader(i)

                                If OLD_GROUP_SUB <> "" Then
                                    S_FINAL_TWD = S_FINAL_TWD + objDataReader(i)
                                End If
                                T_FINAL_TWD = T_FINAL_TWD + objDataReader(i)
                            End If

                        Case 4
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

                        Case 5
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

                        Case 6
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

                        Case 8, 9

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
                td_node.text = ""

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
            td_node.text = ""

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


        End If

        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "E" Then
            '==================================================
            '****************** E 長期投資-權益法 *************
            '==================================================

            sqlStr = "select GR22_ITEM_NM, GR25_ITEM_NM, sum(nvl(GR25_DIFFINV_TWD,0)), sum(nvl(GR25_CASH_TWD,0)),sum(nvl(GR25_FSTOCK_CNT,0))"
            sqlStr = sqlStr & ",sum(nvl(GR25_FINAL_TWD,0)),sum(nvl(GR25_REV_TWD,0)),sum(nvl(GR25_FINVEST_PCT,0)),sum(nvl(GR25_PRICE_TWD,0)),GR01_NM_CA"
            sqlStr = sqlStr & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR25, GR22, GR01"
            sqlStr = sqlStr & " where GR25_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR25_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='E'"

            'If GRXX_COMP_CD <> "BLANK" Then
            '    sqlStr = sqlStr & " AND GR25_COMP_CD = '" & GRXX_COMP_CD & "'"
            'End If

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR25_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If

            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR25_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr = sqlStr & " group by GR22_ITEM_NM,GR25_ITEM_NM,GR01_NM_CA,GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP,GR01_NM_CA"

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
        End If

        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "F" Then
            '==================================================
            '********************* F 短期借款 *****************
            '==================================================

            sqlStr = "select GR22_ITEM_NM, GR26_ITEM_NM, GR26_LEND, sum(nvl(GR26_TWD,0))"
            sqlStr = sqlStr & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR26, GR22"
            sqlStr = sqlStr & " where GR26_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='F'"

            'If GRXX_COMP_CD <> "BLANK" Then
            '    sqlStr = sqlStr & " AND GR26_COMP_CD = '" & GRXX_COMP_CD & "'"
            'End If


            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR26_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If

            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR26_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr = sqlStr & " group by GR22_ITEM_NM,GR26_ITEM_NM,GR26_LEND,GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP"

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

        End If


        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "G" Then
            '==================================================
            '******************* G 應付短期票券 ***************
            '==================================================

            sqlStr = "select GR22_ITEM_CD,GR22_ITEM_NM, GR27_ITEM_NM, GR27_INSTITUT" ', sum(nvl(GR27_TWD,0)),sum(nvl(GR27_DAY,0))"
            sqlStr = sqlStr & ",NVL(GR27_TWD,0) ,  NVL(GR27_DAY,0),  GR27_RATE,  GR27_comp_cd ||' '||GR01_nm_ca"
            sqlStr = sqlStr & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR27, GR22, GR01"
            sqlStr = sqlStr & " where GR27_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='G'"
            sqlStr = sqlStr & "   AND gr01_comp_cd   = GR27_comp_cd"

            'If GRXX_COMP_CD <> "BLANK" Then
            '    sqlStr = sqlStr & " AND GR27_COMP_CD = '" & GRXX_COMP_CD & "'"
            'End If


            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR27_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If

            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR27_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            sqlStr = sqlStr & " group by GR22_ITEM_CD,  GR22_ITEM_NM,  GR27_ITEM_NM,  GR27_comp_cd,  GR27_TWD,  GR27_ITEM_CD,  GR27_RATE,  GR27_DAY,  GR27_INSTITUT,  GR22_GROUP_SEQ,  GR22_GROUP_SUB,  GR22_GROUP,GR01_nm_ca"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,  GR22_GROUP_SUB, GR22_GROUP,  GR27_comp_cd"

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


            headstr = " 項目代碼@1_項目類別@1_項目名稱@1_保證機構@1_金額@1_期間(天)@1_年利率@1_關係企業@1_"
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

            Dim S_GR27_AMT, S_GR27_RATE, S_GR27_DAY As Decimal
            Dim T_GR27_AMT, T_GR27_RATE, T_GR27_DAY As Decimal

            iCnt = 1
            Do While objDataReader.Read()

                '************ 小計 ******************

                If Not IsDBNull(objDataReader(8)) Then
                    TEMP_GROUP_SUB = objDataReader(8)
                Else
                    TEMP_GROUP_SUB = ""
                End If

                If Trim(OLD_GROUP_SUB) = "" Then
                    If Not IsDBNull(objDataReader(8)) Then
                        OLD_GROUP_SUB = objDataReader(8)
                    Else
                        OLD_GROUP_SUB = ""
                    End If
                End If


                If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    td_node.text = ""

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
                    td_node.text = S_GR27_AMT


                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    'td_node.text = S_GR27_DAY

                    '小計歸零
                    S_GR27_AMT = 0
                    S_GR27_RATE = 0
                    S_GR27_DAY = 0


                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If

                '************ 合計 ******************

                If Not IsDBNull(objDataReader(9)) Then
                    TEMP_GROUP = objDataReader(9)
                Else
                    TEMP_GROUP = ""
                End If

                If Trim(OLD_GROUP) = "" Then
                    If Not IsDBNull(objDataReader(9)) Then
                        OLD_GROUP = objDataReader(9)
                    Else
                        OLD_GROUP = ""
                    End If
                End If


                If TEMP_GROUP <> OLD_GROUP Then
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    td_node.text = ""

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
                    td_node.text = T_GR27_AMT


                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    'td_node.text = T_GR27_DAY

                    '合計歸零
                    T_GR27_AMT = 0
                    T_GR27_RATE = 0
                    T_GR27_DAY = 0

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
                                td_node.text = ""
                            Else
                                td_node.text = "G" & objDataReader(i)
                            End If
                        Case 1
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
                        Case 4
                            td_node = XMLDoc.createElement("TD")
                            tr_node.appendChild(td_node)
                            If IsDBNull(objDataReader(i)) Then
                                td_node.text = ""
                                S_GR27_AMT = S_GR27_AMT + 0
                                T_GR27_AMT = T_GR27_AMT + 0
                            Else
                                td_node.text = objDataReader(i)
                                If OLD_GROUP_SUB <> "" Then
                                    S_GR27_AMT = S_GR27_AMT + objDataReader(i)
                                End If
                                T_GR27_AMT = T_GR27_AMT + objDataReader(i)
                            End If

                        Case 5
                            td_node = XMLDoc.createElement("TD")
                            tr_node.appendChild(td_node)
                            If IsDBNull(objDataReader(i)) Then
                                td_node.text = ""
                                S_GR27_DAY = S_GR27_DAY + 0
                                T_GR27_DAY = T_GR27_DAY + 0
                            Else
                                td_node.text = objDataReader(i)

                                If OLD_GROUP_SUB <> "" Then
                                    S_GR27_DAY = S_GR27_DAY + objDataReader(i)
                                End If
                                T_GR27_DAY = T_GR27_DAY + objDataReader(i)
                            End If

                        Case 8, 9
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
                td_node.text = ""

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
                td_node.text = S_GR27_AMT

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                'td_node.text = S_GR27_DAY

            End If


            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

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
            td_node.text = T_GR27_AMT


            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            'td_node.text = T_GR27_DAY
        End If

        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "H" Then
            '==================================================
            '********************* H 長期借款 *****************
            '==================================================

            'sqlStr = "select GR22_ITEM_NM, GR28_ITEM_NM, GR28_LEND, sum(nvl(GR28_ORGLEND_TWD,0)),sum(nvl(GR28_LEND_TWD,0))"
            sqlStr = "select GR22_ITEM_CD, GR22_ITEM_NM, GR28_ITEM_NM,  TO_CHAR(GR28_LEND_FM_DT,'YYYY/MM/DD') GR28_LEND_FM_DT,TO_CHAR(GR28_LEND_TO_DT,'YYYY/MM/DD') GR28_LEND_TO_DT, sum(nvl(GR28_ORGLEND_TWD,0)),sum(nvl(GR28_LEND_TWD,0))"
            sqlStr = sqlStr & ",GR28_RATE,GR28_RATE_2,GR28_CAPI_MH,GR28_INTR_MH,GR28_REF,GR01_NM_CA" '20120302 CCYang add
            sqlStr = sqlStr & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR28, GR22,GR01"
            sqlStr = sqlStr & " where GR28_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='H'"
            sqlStr = sqlStr & "   and GR28_COMP_CD = GR01_COMP_CD (+) "

            'If GRXX_COMP_CD <> "BLANK" Then
            '    sqlStr = sqlStr & " AND GR28_COMP_CD = '" & GRXX_COMP_CD & "'"
            'End If


            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR28_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If

            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR28_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            sqlStr = sqlStr & " group by GR22_ITEM_CD, GR22_ITEM_NM,GR28_ITEM_NM, GR28_LEND_FM_DT,GR28_LEND_TO_DT,GR28_RATE,GR28_RATE_2,GR28_CAPI_MH,GR28_INTR_MH,GR28_REF,GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP,GR01_NM_CA"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP"

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


            headstr = "項目代碼@1_項目類別@1_項目名稱@1_借款期間@1_@1_原借款金額@1_金額@1_年利率(%)@1_@1_幾個月償還本金一次@1_幾個月償還利息一次@1_備註@1_借款公司@1_"
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

            Dim S_GR28_ORGLEND_AMT, S_GR28_LEND_AMT, S_GR28_RATE As Decimal
            Dim T_GR28_ORGLEND_AMT, T_GR28_LEND_AMT, T_GR28_RATE As Decimal

            iCnt = 1
            Do While objDataReader.Read()
                '************ 小計 ******************

                If Not IsDBNull(objDataReader(12)) Then
                    TEMP_GROUP_SUB = objDataReader(12)
                Else
                    TEMP_GROUP_SUB = ""
                End If

                If Trim(OLD_GROUP_SUB) = "" Then
                    If Not IsDBNull(objDataReader(12)) Then
                        OLD_GROUP_SUB = objDataReader(12)
                    Else
                        OLD_GROUP_SUB = ""
                    End If
                End If

                If OLD_GROUP_SUB <> TEMP_GROUP_SUB And OLD_GROUP_SUB <> "" Then
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    td_node.text = ""

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
                    td_node.text = ""

                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    td_node.text = S_GR28_ORGLEND_AMT

                    td_node = XMLDoc.createElement("TD")
                    tr_node.appendChild(td_node)
                    td_node.text = S_GR28_LEND_AMT


                    '小計歸零
                    S_GR28_ORGLEND_AMT = 0
                    S_GR28_LEND_AMT = 0

                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If

                '************ 合計 ******************

                If Not IsDBNull(objDataReader(13)) Then
                    TEMP_GROUP = objDataReader(13)
                Else
                    TEMP_GROUP = ""
                End If

                If Trim(OLD_GROUP) = "" Then
                    If Not IsDBNull(objDataReader(13)) Then
                        OLD_GROUP = objDataReader(13)
                    Else
                        OLD_GROUP = ""
                    End If
                End If


                If TEMP_GROUP <> OLD_GROUP Then
                    '2012/12/21取消不印 CCYang
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
                    'td_node.text = ""

                    'td_node = XMLDoc.createElement("TD")
                    'tr_node.appendChild(td_node)
                    'td_node.text = ""

                    'td_node = XMLDoc.createElement("TD")
                    'tr_node.appendChild(td_node)
                    'td_node.text = T_GR28_ORGLEND_AMT

                    'td_node = XMLDoc.createElement("TD")
                    'tr_node.appendChild(td_node)
                    'td_node.text = T_GR28_LEND_AMT

                    '合計歸零
                    T_GR28_ORGLEND_AMT = 0
                    T_GR28_LEND_AMT = 0


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
                                td_node.text = ""
                            Else
                                td_node.text = "H" & objDataReader(i)
                            End If
                        Case 1
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
                        Case 5
                            td_node = XMLDoc.createElement("TD")
                            tr_node.appendChild(td_node)
                            If IsDBNull(objDataReader(i)) Then
                                td_node.text = ""
                                S_GR28_ORGLEND_AMT = S_GR28_ORGLEND_AMT + 0
                                T_GR28_ORGLEND_AMT = T_GR28_ORGLEND_AMT + 0
                            Else
                                td_node.text = objDataReader(i)

                                If OLD_GROUP_SUB <> "" Then
                                    S_GR28_ORGLEND_AMT = S_GR28_ORGLEND_AMT + objDataReader(i)
                                End If
                                T_GR28_ORGLEND_AMT = T_GR28_ORGLEND_AMT + objDataReader(i)

                            End If

                        Case 6
                            td_node = XMLDoc.createElement("TD")
                            tr_node.appendChild(td_node)
                            If IsDBNull(objDataReader(i)) Then
                                td_node.text = ""
                                S_GR28_LEND_AMT = S_GR28_LEND_AMT + 0
                                T_GR28_LEND_AMT = T_GR28_LEND_AMT + 0
                            Else
                                td_node.text = objDataReader(i)

                                If OLD_GROUP_SUB <> "" Then
                                    S_GR28_LEND_AMT = S_GR28_LEND_AMT + objDataReader(i)
                                End If
                                T_GR28_LEND_AMT = T_GR28_LEND_AMT + objDataReader(i)
                            End If

                        Case 13, 14

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
                td_node.text = ""

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
                td_node.text = ""

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_GR28_ORGLEND_AMT

                td_node = XMLDoc.createElement("TD")
                tr_node.appendChild(td_node)
                td_node.text = S_GR28_LEND_AMT

            End If

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

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
            td_node.text = ""

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = T_GR28_ORGLEND_AMT

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = T_GR28_LEND_AMT
        End If

        ' 30DEC2019 BruceWang (081075)增加欄位 代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "I" Then
            '==================================================
            '********************* I 質押之資產 ***************
            '==================================================

            sqlStr = "select  GR07_ACCT_GAH, GR29_ITEM_NM, GR07_ACCT_NM_CF, GR01_NM_CA, sum(nvl(GR29_TWD,0))"
            sqlStr = sqlStr & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR29, GR07, GR22, GR01"
            sqlStr = sqlStr & " where GR29_ACCT_GAH = GR07_ACCT_GAH(+)"
            sqlStr = sqlStr & " and GR29_ACCT_SUBAH = GR07_ACCT_SUBAH(+)"
            sqlStr = sqlStr & " and GR29_ACCT_SUBBH = GR07_ACCT_SUBBH(+)"
            sqlStr = sqlStr & " and GR29_ITEM_CD = GR22_ITEM_CD(+)"
            sqlStr = sqlStr & " and GR29_COMP_CD = GR01_COMP_CD(+)"
            'sqlStr = sqlStr & "   and GR22_DATA_CD ='I'"

            'If GRXX_COMP_CD <> "BLANK" Then
            '    sqlStr = sqlStr & " AND GR29_COMP_CD = '" & GRXX_COMP_CD & "'"
            'End If


            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR29_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If

            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR29_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            sqlStr = sqlStr & " group by GR29_ITEM_NM,GR07_ACCT_NM_CF,GR01_NM_CA,GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP,GR07_ACCT_GAH"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP,GR01_NM_CA"

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


            headstr = " 代碼@1_項目類別@1_帳列科目@1_投資公司@1_金額@1_"
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

            Dim S_GR29_AMT As Decimal
            Dim T_GR29_AMT As Decimal


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
                    td_node.text = ""

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
                    td_node.text = S_GR29_AMT

                    '小計歸零
                    S_GR29_AMT = 0

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
                    td_node.text = ""

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
                    td_node.text = T_GR29_AMT

                    '合計歸零
                    T_GR29_AMT = 0


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
                                td_node.text = ""
                            Else
                                td_node.text = "I" & objDataReader(i)
                            End If

                        Case 1
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
                        Case 4

                            td_node = XMLDoc.createElement("TD")
                            tr_node.appendChild(td_node)
                            If IsDBNull(objDataReader(i)) Then
                                td_node.text = ""
                                S_GR29_AMT = S_GR29_AMT + 0
                                T_GR29_AMT = T_GR29_AMT + 0
                            Else
                                td_node.text = objDataReader(i)

                                If OLD_GROUP_SUB <> "" Then
                                    S_GR29_AMT = S_GR29_AMT + objDataReader(i)
                                End If
                                T_GR29_AMT = T_GR29_AMT + objDataReader(i)
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
                td_node.text = ""

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
                td_node.text = S_GR29_AMT

            End If


            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            tr_node.appendChild(td_node)
            td_node.text = ""

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
            td_node.text = T_GR29_AMT
        End If


        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "J" Then
            '==================================================
            '******** J 用人、折舊及攤銷費用之功能別彙總 ******
            '==================================================

            'sqlStr = "select decode(nvl(GR30_ITEM_NM,'0'),'0',GR22_ITEM_NM,decode(rtrim(ltrim(GR30_ITEM_NM)),'',GR22_ITEM_NM,GR30_ITEM_NM)) as GR30_ITEM_NM, GR30_COST_TWD, GR30_EXP_TWD"
            sqlStr = ""
            'sqlStr = "select GR30_ITEM_NM, GR01_NM_CA, sum(nvl(GR30_COST_TWD,0)), sum(nvl(GR30_EXP_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR30, GR22, GR01"
            sqlStr = sqlStr & " where GR30_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR30_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='J'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR30_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR30_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            Dim sCompList
            sCompList = "select GR01_COMP_CD , GR01_NM_CA " & sqlStr
            sCompList = sCompList & " group by GR01_COMP_CD,GR01_NM_CA"
            sCompList = sCompList & " order by GR01_COMP_CD"
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
            td_node.text = "項目代碼"
            'td_node.setAttribute("rowspan", "2")
            tr_node.appendChild(td_node)

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
            td_node.setAttribute("colspan", "3")
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = ""
            tr_node.appendChild(td_node)
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
            td_node = XMLDoc.createElement("TD")
            td_node.text = "合計"
            tr_node.appendChild(td_node)


            sqlStr = "select GR22_ITEM_CD, GR30_ITEM_NM, GR01_NM_CA, sum(nvl(GR30_COST_TWD,0)), sum(nvl(GR30_EXP_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP " & sqlStr
            sqlStr = sqlStr & " group by GR30_ITEM_NM,GR01_NM_CA,GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP,GR01_COMP_CD,GR22_ITEM_CD"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR01_COMP_CD,GR22_GROUP"
            objDataReader = objDataservice.ExecuteReader(sqlStr)

            OLD_GROUP_SUB = ""
            TEMP_GROUP_SUB = ""
            OLD_GROUP = ""
            TEMP_GROUP = ""

            Dim S_GR30_COST_AMT, S_GR30_EXP_AMT As Decimal
            Dim T_GR30_COST_AMT, T_GR30_EXP_AMT As Decimal
            Dim iCompCnt
            iCompCnt = 0
            iCnt = 1
            Do While objDataReader.Read()
                TEMP_GROUP = objDataReader(1)
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
                        '合計金額
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = T_GR30_COST_AMT
                        tr_node.appendChild(td_node)
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = T_GR30_EXP_AMT
                        tr_node.appendChild(td_node)

                        td_node = XMLDoc.createElement("TD")
                        td_node.text = T_GR30_COST_AMT + T_GR30_EXP_AMT
                        tr_node.appendchild(td_node)

                        T_GR30_COST_AMT = 0
                        T_GR30_EXP_AMT = 0
                    End If


                    OLD_GROUP = TEMP_GROUP
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "J" & objDataReader(0)
                    tr_node.appendChild(td_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(1)
                    tr_node.appendChild(td_node)
                    iCompCnt = 0
                End If

                For i = iCompCnt To oCompLists.Count - 1
                    iCompCnt = iCompCnt + 1
                    If oCompLists(i)(1) = objDataReader(2) Then
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = objDataReader(3)
                        tr_node.appendChild(td_node)
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = objDataReader(4)
                        tr_node.appendChild(td_node)
                        oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(3)
                        oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(4)
                        T_GR30_COST_AMT = T_GR30_COST_AMT + objDataReader(3)
                        T_GR30_EXP_AMT = T_GR30_EXP_AMT + objDataReader(4)
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
            '合計金額
            td_node = XMLDoc.createElement("TD")
            td_node.text = T_GR30_COST_AMT
            tr_node.appendChild(td_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = T_GR30_EXP_AMT
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = T_GR30_COST_AMT + T_GR30_EXP_AMT
            tr_node.appendchild(td_node)

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

            'Dim S_GR30_COST_AMT, S_GR30_EXP_AMT As Decimal
            'Dim T_GR30_COST_AMT, T_GR30_EXP_AMT As Decimal

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
            '        td_node.text = S_GR30_COST_AMT

            '        td_node = XMLDoc.createElement("TD")
            '        tr_node.appendChild(td_node)
            '        td_node.text = S_GR30_EXP_AMT

            '        '小計歸零
            '        S_GR30_COST_AMT = 0
            '        S_GR30_EXP_AMT = 0


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
            '        td_node.text = T_GR30_COST_AMT

            '        td_node = XMLDoc.createElement("TD")
            '        tr_node.appendChild(td_node)
            '        td_node.text = T_GR30_EXP_AMT

            '        '小計歸零
            '        T_GR30_COST_AMT = 0
            '        T_GR30_EXP_AMT = 0

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
            '                    S_GR30_COST_AMT = S_GR30_COST_AMT + 0
            '                    T_GR30_COST_AMT = T_GR30_COST_AMT + 0
            '                Else
            '                    td_node.text = objDataReader(i)
            '                    If OLD_GROUP_SUB <> "" Then
            '                        S_GR30_COST_AMT = S_GR30_COST_AMT + objDataReader(i)
            '                    End If
            '                    T_GR30_COST_AMT = T_GR30_COST_AMT + objDataReader(i)

            '                End If
            '            Case 3
            '                td_node = XMLDoc.createElement("TD")
            '                tr_node.appendChild(td_node)
            '                If IsDBNull(objDataReader(i)) Then
            '                    td_node.text = ""
            '                    S_GR30_EXP_AMT = S_GR30_EXP_AMT + 0
            '                    T_GR30_EXP_AMT = T_GR30_EXP_AMT + 0
            '                Else
            '                    td_node.text = objDataReader(i)

            '                    If OLD_GROUP_SUB <> "" Then
            '                        S_GR30_EXP_AMT = S_GR30_EXP_AMT + objDataReader(i)
            '                    End If
            '                    T_GR30_EXP_AMT = T_GR30_EXP_AMT + objDataReader(i)

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
            '    td_node.text = S_GR30_COST_AMT

            '    td_node = XMLDoc.createElement("TD")
            '    tr_node.appendChild(td_node)
            '    td_node.text = S_GR30_EXP_AMT
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
            'td_node.text = T_GR30_COST_AMT

            'td_node = XMLDoc.createElement("TD")
            'tr_node.appendChild(td_node)
            'td_node.text = T_GR30_EXP_AMT
        End If

        ' 20190304 BruceWang (080156) add Q 變動表-使用權資產
        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "Q" Then
            '==================================================
            '******** Q 使用權資產 GR57  **************
            '==================================================
            sqlStr = ""
            sqlStr = sqlStr & " from GR57, GR22, GR01"
            sqlStr = sqlStr & " where GR57_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR57_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='Q'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR57_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR57_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            Dim sCompList '關係企業 List
            sCompList = "select GR01_COMP_CD , GR01_NM_CA " & sqlStr
            sCompList = sCompList & " group by GR01_COMP_CD,GR01_NM_CA"
            sCompList = sCompList & " order by GR01_COMP_CD"
            objDataReader = objDataservice.ExecuteReader(sCompList)

            Dim oCompLists As New ArrayList
            oCompLists.Clear()
            Do While objDataReader.Read()
                Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                oCompLists.Add(oColumn)
            Loop

            ' 建立 TABLE 
            table_node = XMLDoc.createElement("TABLE")
            oRootNode.appendChild(table_node)
            table_node.setAttribute("name", "使用權資產")
            'table_node.SetAttribute("border", "1")

            '***************** 建立 表頭
            thead1_node = XMLDoc.createElement("THEAD")
            table_node.appendChild(thead1_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "使用權資產"
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目代碼"
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目類別"
            'td_node.setAttribute("rowspan", "2")
            tr_node.appendChild(td_node)

            For i = 0 To oCompLists.Count - 1
                td_node = XMLDoc.createElement("TD")
                td_node.text = oCompLists(i)(1)
                td_node.setAttribute("colspan", "1")
                tr_node.appendChild(td_node)
            Next

            td_node = XMLDoc.createElement("TD")
            td_node.text = "合計"
            td_node.setAttribute("colspan", "1")
            tr_node.appendChild(td_node)

            'sqlStr = "select GR57_ITEM_NM, GR01_NM_CA, sum(nvl(GR57_COST_TWD,0)), sum(nvl(GR57_EXP_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP " & sqlStr
            Dim sqlStrT As String
            sqlStrT = "select GR22_ITEM_CD, GR22_ITEM_NM, GR57_ITEM_NM"
            sqlStrT = sqlStrT & ",GR57_TWD"
            sqlStrT = sqlStrT & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStrT = sqlStrT & ",GR22_GROUP_SEQ,NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP,GR01_COMP_CD"
            sqlStr = sqlStrT & sqlStr


            sqlStr = sqlStr & " UNION "
            sqlStr = sqlStr & "select GR01v_ITEM_CD,GR01v_ITEM_NM, GR01v_ITEM_NM"
            sqlStr = sqlStr & ",GR01v_TWD"
            sqlStr = sqlStr & ",GR01v_GROUP_SUB,GR01v_GROUP"
            sqlStr = sqlStr & ",GR01v_GROUP_SEQ,GR01v_HEAD_YN ,GR01v_GROUP_OP,GR01v_COMP_CD"
            sqlStr = sqlStr & " from GR01v, "
            sqlStr = sqlStr & " ("
            sqlStr = sqlStr & "select GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & "  from GR57, GR22, GR01"
            sqlStr = sqlStr & " where GR57_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR57_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD = 'Q'"
            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR57_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR57_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr = sqlStr & " group by GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & " order by GR01_COMP_CD"
            sqlStr = sqlStr & " ) GR01"

            sqlStr = sqlStr & " where GR01v_DATA_CD = 'Q'"
            sqlStr = sqlStr & " AND GR01v_COMP_CD = GR01.GR01_COMP_CD"

            sqlStr = sqlStr & " group by gr01v_ITEM_NM,gr01v_ITEM_NM,gr01v_TWD,gr01v_GROUP_SUB,gr01v_GROUP, gr01v_GROUP_SEQ, gr01v_HEAD_YN, gr01v_GROUP_OP, gr01v_COMP_CD, GR01v_ITEM_CD"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR01_COMP_CD,GR22_GROUP"
            objDataReader = objDataservice.ExecuteReader(sqlStr)

            Dim TEMP_GROUP_SEQ, OLD_GROUP_SEQ As String

            OLD_GROUP_SEQ = ""
            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SUB = ""
            TEMP_GROUP_SUB = ""
            OLD_GROUP = ""
            TEMP_GROUP = ""
            OLD_ITEM_NM = ""

            Dim S_GR57_AMT As Decimal
            Dim T_GR57_AMT As Decimal
            Dim T_GR57_AMT_GROUP As Decimal
            Dim T_GR57_AMT_GROUP_SUB As Decimal
            Dim iCompCnt '關係企業數

            iCompCnt = 0
            iCnt = 1
            Dim j = 0
            T_GR57_AMT_GROUP = 0 'GROUP 合計
            T_GR57_AMT_GROUP_SUB = 0 'GROUP_SUB 合計
            S_GR57_AMT = 0 ' 右方小計
            Dim ld_data As Decimal '暫存資料
            lb_group = False
            lb_group_sub = False

            Do While objDataReader.Read()
                'Group_seq,Group_sub,Group 
                'Group_seq
                If IsDBNull(objDataReader(6)) Then
                    TEMP_GROUP_SEQ = ""
                Else
                    TEMP_GROUP_SEQ = objDataReader(6)
                End If
                If OLD_GROUP_SEQ = "" Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                End If

                If OLD_GROUP_SEQ <> TEMP_GROUP_SEQ Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                    '關係企業數歸零
                    'iCompCnt = 0
                End If
                'Group_sub
                If IsDBNull(objDataReader(4)) Then
                    TEMP_GROUP_SUB = ""
                Else
                    TEMP_GROUP_SUB = objDataReader(4)
                End If
                If OLD_GROUP_SUB = "" Then
                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If
                'Group
                If IsDBNull(objDataReader(5)) Then
                    TEMP_GROUP = ""
                Else
                    TEMP_GROUP = objDataReader(5)
                End If
                If OLD_GROUP = "" Then
                    OLD_GROUP = TEMP_GROUP
                End If


                '*** new, GR22_ITEM_NM
                TEMP_ITEM_NM = objDataReader(1)
                If OLD_ITEM_NM <> TEMP_ITEM_NM Then

                    If OLD_ITEM_NM <> "" Then
                        For i = 1 To oCompLists.Count - iCompCnt
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)
                        Next

                        lb_group_sub = False
                        lb_group = False
                        If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                            lb_group_sub = True
                            OLD_GROUP_SUB = TEMP_GROUP_SUB

                            'td_node.text = T_GR57_AMT_GROUP_SUB
                            'T_GR57_AMT_GROUP_SUB = 0 '3
                        End If
                        If OLD_GROUP <> TEMP_GROUP Then 'Group
                            lb_group = True ' 2 擇 1
                            lb_group_sub = False
                            OLD_GROUP = TEMP_GROUP
                            'td_node.text = T_GR57_AMT_GROUP
                            'T_GR57_AMT_GROUP = 0 '2
                            'OLD_GROUP = ""
                        End If
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = S_GR57_AMT
                        tr_node.appendChild(td_node)
                        S_GR57_AMT = 0


                    End If  'OLD_ITEM_NM <> ""


                    OLD_ITEM_NM = TEMP_ITEM_NM
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "Q" & objDataReader(0)
                    tr_node.appendChild(td_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(1)
                    tr_node.appendChild(td_node)
                    iCompCnt = 0
                End If  'OLD_ITEM_NM <> TEMP_ITEM_NM


                '關係企業數 Loop
                For i = iCompCnt To oCompLists.Count - 1
                    iCompCnt = iCompCnt + 1
                    '關係企業 COMP_CD
                    If oCompLists(i)(0) = objDataReader(9) Then
                        'If iCompCnt = 0 Then
                        '    'Item 項目
                        '    tr_node = XMLDoc.createElement("TR")
                        '    thead1_node.appendChild(tr_node)

                        '    td_node = XMLDoc.createElement("TD")
                        '    td_node.text = objDataReader(0)
                        '    tr_node.appendChild(td_node)
                        'End If

                        '處理說明欄位與非說明欄位
                        If Not IsDBNull(objDataReader(7)) And objDataReader(7) = "N" Then  '為 非說明欄 head_yn = 'N' 
                            td_node = XMLDoc.createElement("TD")
                            'Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                            td_node.text = objDataReader(3)
                            tr_node.appendChild(td_node)
                            If Not IsDBNull(objDataReader(8)) And objDataReader(8) = "-" Then
                                oCompLists(i)(2) = oCompLists(i)(2) - objDataReader(3) ' Group
                                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) ' Group_sub
                                T_GR57_AMT = T_GR57_AMT - objDataReader(3) '合計
                                S_GR57_AMT += objDataReader(3) '右方合計
                            Else
                                oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(3) ' Group
                                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) ' Group_sub
                                T_GR57_AMT = T_GR57_AMT + objDataReader(3) '合計
                                S_GR57_AMT += objDataReader(3) '右方合計
                            End If
                            Exit For

                        Else
                            td_node = XMLDoc.createelement("TD")
                            If lb_group_sub Then 'Group_sub
                                td_node.text = oCompLists(i)(3)
                                T_GR57_AMT_GROUP_SUB += oCompLists(i)(3)
                                S_GR57_AMT += oCompLists(i)(3) '右方合計
                                oCompLists(i)(3) = 0
                            End If
                            If lb_group Then 'Group
                                td_node.text = oCompLists(i)(2)
                                T_GR57_AMT_GROUP += oCompLists(i)(2)
                                S_GR57_AMT += oCompLists(i)(2) '右方合計
                                oCompLists(i)(2) = 0
                                oCompLists(i)(3) = 0
                            End If

                            tr_node.appendChild(td_node)
                            Exit For

                        End If

                    Else  '自進入LOOP都未找到該公司,表無資料,故直接放0
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = 0
                        tr_node.appendChild(td_node)

                    End If 'END oCompLists(i)(0) = objDataReader(8)

                Next

            Loop

            ''項目合計-右方
            For i = 1 To oCompLists.Count - iCompCnt
                td_node = XMLDoc.createElement("TD")
                td_node.text = "0"
                tr_node.appendChild(td_node)
            Next
            td_node = XMLDoc.createElement("TD")
            td_node.text = S_GR57_AMT
            tr_node.appendChild(td_node)

            ''項目合計-右方
            'For i = 1 To oCompLists.Count - iCompCnt
            '    td_node = XMLDoc.createElement("TD")
            '    td_node.text = "0"
            '    tr_node.appendChild(td_node)
            'Next
            ''If iCompCnt = oCompLists.Count Then
            'OLD_GROUP_SEQ = TEMP_GROUP_SEQ

            'td_node = XMLDoc.createElement("TD")
            'td_node.text = T_GR57_AMT
            'If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
            '    td_node.text = T_GR57_AMT_GROUP_SUB
            '    T_GR57_AMT_GROUP_SUB = 0 '3
            '    OLD_GROUP_SUB = ""
            'End If
            'If OLD_GROUP <> TEMP_GROUP Then 'Group
            '    td_node.text = T_GR57_AMT_GROUP
            '    T_GR57_AMT_GROUP = 0 '2
            '    'OLD_GROUP_SUB = TEMP_GROUP_SUB
            '    OLD_GROUP = ""
            'End If
            'tr_node.appendChild(td_node)

            '關係企業數歸零
            'iCompCnt = 0
            'T_GR57_AMT = 0
            'End If
        End If

        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "R" Then
            '==================================================
            '******** R 不動產、廠房及設備 GR43  **************
            '==================================================
            sqlStr = ""
            sqlStr = sqlStr & " from GR43, GR22, GR01"
            sqlStr = sqlStr & " where GR43_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR43_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='R'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR43_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR43_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            Dim sCompList '關係企業 List
            sCompList = "select GR01_COMP_CD , GR01_NM_CA " & sqlStr
            sCompList = sCompList & " group by GR01_COMP_CD,GR01_NM_CA"
            sCompList = sCompList & " order by GR01_COMP_CD"
            objDataReader = objDataservice.ExecuteReader(sCompList)

            Dim oCompLists As New ArrayList
            oCompLists.Clear()
            Do While objDataReader.Read()
                Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                oCompLists.Add(oColumn)
            Loop

            ' 建立 TABLE 
            table_node = XMLDoc.createElement("TABLE")
            oRootNode.appendChild(table_node)
            table_node.setAttribute("name", "不動產、廠房及設備彙總")
            'table_node.SetAttribute("border", "1")

            '***************** 建立 表頭
            thead1_node = XMLDoc.createElement("THEAD")
            table_node.appendChild(thead1_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "不動產、廠房及設備彙總"
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目代碼"
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目類別"
            'td_node.setAttribute("rowspan", "2")
            tr_node.appendChild(td_node)

            For i = 0 To oCompLists.Count - 1
                td_node = XMLDoc.createElement("TD")
                td_node.text = oCompLists(i)(1)
                td_node.setAttribute("colspan", "1")
                tr_node.appendChild(td_node)
            Next

            td_node = XMLDoc.createElement("TD")
            td_node.text = "合計"
            td_node.setAttribute("colspan", "1")
            tr_node.appendChild(td_node)

            'sqlStr = "select GR43_ITEM_NM, GR01_NM_CA, sum(nvl(GR43_COST_TWD,0)), sum(nvl(GR43_EXP_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP " & sqlStr
            Dim sqlStrT As String
            sqlStrT = "select GR22_ITEM_CD, GR22_ITEM_NM, GR43_ITEM_NM"
            sqlStrT = sqlStrT & ",GR43_TWD"
            sqlStrT = sqlStrT & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStrT = sqlStrT & ",GR22_GROUP_SEQ,NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP,GR01_COMP_CD"
            sqlStr = sqlStrT & sqlStr


            sqlStr = sqlStr & " UNION "
            sqlStr = sqlStr & "select GR01v_ITEM_CD, GR01v_ITEM_NM, GR01v_ITEM_NM"
            sqlStr = sqlStr & ",GR01v_TWD"
            sqlStr = sqlStr & ",GR01v_GROUP_SUB,GR01v_GROUP"
            sqlStr = sqlStr & ",GR01v_GROUP_SEQ,GR01v_HEAD_YN ,GR01v_GROUP_OP,GR01v_COMP_CD"
            sqlStr = sqlStr & " from GR01v, "
            sqlStr = sqlStr & " ("
            sqlStr = sqlStr & "select GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & "  from GR43, GR22, GR01"
            sqlStr = sqlStr & " where GR43_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR43_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD = 'R'"
            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR43_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR43_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr = sqlStr & " group by GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & " order by GR01_COMP_CD"
            sqlStr = sqlStr & " ) GR01"

            sqlStr = sqlStr & " where GR01v_DATA_CD = 'R'"
            sqlStr = sqlStr & " AND GR01v_COMP_CD = GR01.GR01_COMP_CD"

            sqlStr = sqlStr & " group by gr01v_ITEM_NM,gr01v_ITEM_NM,gr01v_TWD,gr01v_GROUP_SUB,gr01v_GROUP, gr01v_GROUP_SEQ, gr01v_HEAD_YN, gr01v_GROUP_OP, gr01v_COMP_CD,GR01v_ITEM_CD"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR01_COMP_CD,GR22_GROUP"
            objDataReader = objDataservice.ExecuteReader(sqlStr)

            Dim TEMP_GROUP_SEQ, OLD_GROUP_SEQ As String

            OLD_GROUP_SEQ = ""
            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SUB = ""
            TEMP_GROUP_SUB = ""
            OLD_GROUP = ""
            TEMP_GROUP = ""
            OLD_ITEM_NM = ""

            Dim S_GR43_AMT As Decimal
            Dim T_GR43_AMT As Decimal
            Dim T_GR43_AMT_GROUP As Decimal
            Dim T_GR43_AMT_GROUP_SUB As Decimal
            Dim iCompCnt '關係企業數

            iCompCnt = 0
            iCnt = 1
            Dim j = 0
            T_GR43_AMT_GROUP = 0 'GROUP 合計
            T_GR43_AMT_GROUP_SUB = 0 'GROUP_SUB 合計
            S_GR43_AMT = 0 ' 右方小計
            Dim ld_data As Decimal '暫存資料
            lb_group = False
            lb_group_sub = False

            Do While objDataReader.Read()
                'Group_seq,Group_sub,Group 
                'Group_seq
                If IsDBNull(objDataReader(6)) Then
                    TEMP_GROUP_SEQ = ""
                Else
                    TEMP_GROUP_SEQ = objDataReader(6)
                End If
                If OLD_GROUP_SEQ = "" Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                End If

                If OLD_GROUP_SEQ <> TEMP_GROUP_SEQ Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                    '關係企業數歸零
                    'iCompCnt = 0
                End If
                'Group_sub
                If IsDBNull(objDataReader(4)) Then
                    TEMP_GROUP_SUB = ""
                Else
                    TEMP_GROUP_SUB = objDataReader(4)
                End If
                If OLD_GROUP_SUB = "" Then
                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If
                'Group
                If IsDBNull(objDataReader(5)) Then
                    TEMP_GROUP = ""
                Else
                    TEMP_GROUP = objDataReader(5)
                End If
                If OLD_GROUP = "" Then
                    OLD_GROUP = TEMP_GROUP
                End If


                '*** new, GR22_ITEM_NM
                TEMP_ITEM_NM = objDataReader(1)
                If OLD_ITEM_NM <> TEMP_ITEM_NM Then

                    If OLD_ITEM_NM <> "" Then
                        For i = 1 To oCompLists.Count - iCompCnt
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)
                        Next

                        lb_group_sub = False
                        lb_group = False
                        If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                            lb_group_sub = True
                            OLD_GROUP_SUB = TEMP_GROUP_SUB

                            'td_node.text = T_GR43_AMT_GROUP_SUB
                            'T_GR43_AMT_GROUP_SUB = 0 '3
                        End If
                        If OLD_GROUP <> TEMP_GROUP Then 'Group
                            lb_group = True ' 2 擇 1
                            lb_group_sub = False
                            OLD_GROUP = TEMP_GROUP
                            'td_node.text = T_GR43_AMT_GROUP
                            'T_GR43_AMT_GROUP = 0 '2
                            'OLD_GROUP = ""
                        End If
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = S_GR43_AMT
                        tr_node.appendChild(td_node)
                        S_GR43_AMT = 0


                    End If  'OLD_ITEM_NM <> ""


                    OLD_ITEM_NM = TEMP_ITEM_NM
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "R" & objDataReader(0)
                    tr_node.appendChild(td_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(1)
                    tr_node.appendChild(td_node)
                    iCompCnt = 0
                End If  'OLD_ITEM_NM <> TEMP_ITEM_NM


                '關係企業數 Loop
                For i = iCompCnt To oCompLists.Count - 1
                    iCompCnt = iCompCnt + 1
                    '關係企業 COMP_CD
                    If oCompLists(i)(0) = objDataReader(9) Then
                        'If iCompCnt = 0 Then
                        '    'Item 項目
                        '    tr_node = XMLDoc.createElement("TR")
                        '    thead1_node.appendChild(tr_node)

                        '    td_node = XMLDoc.createElement("TD")
                        '    td_node.text = objDataReader(0)
                        '    tr_node.appendChild(td_node)
                        'End If

                        '處理說明欄位與非說明欄位
                        If Not IsDBNull(objDataReader(7)) And objDataReader(7) = "N" Then  '為 非說明欄 head_yn = 'N' 
                            td_node = XMLDoc.createElement("TD")
                            'Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                            td_node.text = objDataReader(3)
                            tr_node.appendChild(td_node)
                            If Not IsDBNull(objDataReader(8)) And objDataReader(8) = "-" Then
                                oCompLists(i)(2) = oCompLists(i)(2) - objDataReader(3) ' Group
                                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) ' Group_sub
                                T_GR43_AMT = T_GR43_AMT - objDataReader(3) '合計
                                S_GR43_AMT += objDataReader(3) '右方合計
                            Else
                                oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(3) ' Group
                                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) ' Group_sub
                                T_GR43_AMT = T_GR43_AMT + objDataReader(3) '合計
                                S_GR43_AMT += objDataReader(3) '右方合計
                            End If
                            Exit For

                        Else
                            td_node = XMLDoc.createelement("TD")
                            If lb_group_sub Then 'Group_sub
                                td_node.text = oCompLists(i)(3)
                                T_GR43_AMT_GROUP_SUB += oCompLists(i)(3)
                                S_GR43_AMT += oCompLists(i)(3) '右方合計
                                oCompLists(i)(3) = 0
                            End If
                            If lb_group Then 'Group
                                td_node.text = oCompLists(i)(2)
                                T_GR43_AMT_GROUP += oCompLists(i)(2)
                                S_GR43_AMT += oCompLists(i)(2) '右方合計
                                oCompLists(i)(2) = 0
                                oCompLists(i)(3) = 0
                            End If

                            tr_node.appendChild(td_node)
                            Exit For

                        End If

                    Else  '自進入LOOP都未找到該公司,表無資料,故直接放0
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = 0
                        tr_node.appendChild(td_node)

                    End If 'END oCompLists(i)(0) = objDataReader(8)

                Next

            Loop

            ''項目合計-右方
            For i = 1 To oCompLists.Count - iCompCnt
                td_node = XMLDoc.createElement("TD")
                td_node.text = "0"
                tr_node.appendChild(td_node)
            Next
            td_node = XMLDoc.createElement("TD")
            td_node.text = S_GR43_AMT
            tr_node.appendChild(td_node)

            ''項目合計-右方
            'For i = 1 To oCompLists.Count - iCompCnt
            '    td_node = XMLDoc.createElement("TD")
            '    td_node.text = "0"
            '    tr_node.appendChild(td_node)
            'Next
            ''If iCompCnt = oCompLists.Count Then
            'OLD_GROUP_SEQ = TEMP_GROUP_SEQ

            'td_node = XMLDoc.createElement("TD")
            'td_node.text = T_GR43_AMT
            'If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
            '    td_node.text = T_GR43_AMT_GROUP_SUB
            '    T_GR43_AMT_GROUP_SUB = 0 '3
            '    OLD_GROUP_SUB = ""
            'End If
            'If OLD_GROUP <> TEMP_GROUP Then 'Group
            '    td_node.text = T_GR43_AMT_GROUP
            '    T_GR43_AMT_GROUP = 0 '2
            '    'OLD_GROUP_SUB = TEMP_GROUP_SUB
            '    OLD_GROUP = ""
            'End If
            'tr_node.appendChild(td_node)

            '關係企業數歸零
            'iCompCnt = 0
            'T_GR43_AMT = 0
            'End If
        End If

        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "S" Then
            '==================================================
            '************ S 負債準備彙總 GR44  ****************
            '==================================================
            sqlStr = ""
            sqlStr = sqlStr & " from GR44, GR22, GR01"
            sqlStr = sqlStr & " where GR44_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR44_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='S'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR44_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR44_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            Dim sCompList ' 關係企業 List
            sCompList = "select GR01_COMP_CD , GR01_NM_CA " & sqlStr
            sCompList = sCompList & " group by GR01_COMP_CD,GR01_NM_CA"
            sCompList = sCompList & " order by GR01_COMP_CD"
            objDataReader = objDataservice.ExecuteReader(sCompList)


            Dim oCompLists As New ArrayList
            oCompLists.Clear()  'initial 

            Do While objDataReader.Read()
                Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                oCompLists.Add(oColumn)
            Loop

            ' 建立 TABLE 
            table_node = XMLDoc.createElement("TABLE")
            oRootNode.appendChild(table_node)
            table_node.setAttribute("name", "負債準備彙總")
            'table_node.SetAttribute("border", "1")

            '***************** 建立 表頭
            thead1_node = XMLDoc.createElement("THEAD")
            table_node.appendChild(thead1_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "負債準備彙總"
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目代碼"
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目類別"
            'td_node.setAttribute("rowspan", "2")
            tr_node.appendChild(td_node)

            For i = 0 To oCompLists.Count - 1
                td_node = XMLDoc.createElement("TD")
                td_node.text = oCompLists(i)(1)
                td_node.setAttribute("colspan", "1")
                tr_node.appendChild(td_node)
            Next

            td_node = XMLDoc.createElement("TD")
            td_node.text = "合計"
            td_node.setAttribute("colspan", "1")
            tr_node.appendChild(td_node)

            'sqlStr = "select GR44_ITEM_NM, GR01_NM_CA, sum(nvl(GR44_COST_TWD,0)), sum(nvl(GR44_EXP_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP " & sqlStr
            Dim sqlStrT As String
            sqlStrT = "select GR22_ITEM_CD, GR22_ITEM_NM, GR44_ITEM_NM"
            sqlStrT = sqlStrT & ",GR44_TWD"
            sqlStrT = sqlStrT & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStrT = sqlStrT & ",GR22_GROUP_SEQ,NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP,GR01_COMP_CD"
            sqlStr = sqlStrT & sqlStr


            sqlStr = sqlStr & " UNION "
            sqlStr = sqlStr & "select GR01v_ITEM_CD, GR01v_ITEM_NM, GR01v_ITEM_NM"
            sqlStr = sqlStr & ",GR01v_TWD"
            sqlStr = sqlStr & ",GR01v_GROUP_SUB,GR01v_GROUP"
            sqlStr = sqlStr & ",GR01v_GROUP_SEQ,GR01v_HEAD_YN ,GR01v_GROUP_OP,GR01v_COMP_CD"
            sqlStr = sqlStr & " from GR01v, "
            sqlStr = sqlStr & " ("
            sqlStr = sqlStr & "select GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & "  from GR44, GR22, GR01"
            sqlStr = sqlStr & " where GR44_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR44_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD = 'S'"
            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR44_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR44_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr = sqlStr & " group by GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & " order by GR01_COMP_CD"
            sqlStr = sqlStr & " ) GR01"

            sqlStr = sqlStr & " where GR01v_DATA_CD = 'S'"
            sqlStr = sqlStr & " AND GR01v_COMP_CD = GR01_COMP_CD"

            sqlStr = sqlStr & " group by gr01v_ITEM_NM,gr01v_ITEM_NM,gr01v_TWD,gr01v_GROUP_SUB,gr01v_GROUP, gr01v_GROUP_SEQ, gr01v_HEAD_YN, gr01v_GROUP_OP, gr01v_COMP_CD, GR01v_ITEM_CD"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR01_COMP_CD,GR22_GROUP"
            objDataReader = objDataservice.ExecuteReader(sqlStr)

            Dim TEMP_GROUP_SEQ, OLD_GROUP_SEQ As String

            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SEQ = ""

            OLD_GROUP_SEQ = ""
            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SUB = ""
            TEMP_GROUP_SUB = ""
            OLD_GROUP = ""
            TEMP_GROUP = ""
            OLD_ITEM_NM = ""

            Dim S_GR44_AMT As Decimal
            Dim T_GR44_AMT As Decimal
            Dim T_GR44_AMT_GROUP As Decimal
            Dim T_GR44_AMT_GROUP_SUB As Decimal

            Dim iCompCnt ' 關係企業數
            Dim j
            iCompCnt = 0
            iCnt = 1
            j = 0
            T_GR44_AMT_GROUP = 0 'GROUP 合計
            T_GR44_AMT_GROUP_SUB = 0 'GROUP_SUB 合計

            'Dim TEMP_GROUP_SEQ, OLD_GROUP_SEQ As String
            'TEMP_GROUP_SEQ = ""
            'OLD_GROUP_SEQ = ""

            'OLD_GROUP_SEQ = ""
            'TEMP_GROUP_SEQ = ""
            'OLD_GROUP_SUB = ""
            'TEMP_GROUP_SUB = ""
            'OLD_GROUP = ""
            'TEMP_GROUP = ""
            'Dim S_GR44_TWD As Decimal
            'Dim T_GR44_AMT As Decimal
            'Dim T_GR44_AMT_GROUP As Decimal
            'Dim T_GR44_AMT_GROUP_SUB As Decimal

            'Dim iCompCnt '關係企業數
            'Dim j
            'iCompCnt = 0
            'iCnt = 1
            'j = 0
            'T_GR44_AMT_GROUP = 0 'GROUP 合計
            'T_GR44_AMT_GROUP_SUB = 0 'GROUP_SUB 合計


            Do While objDataReader.Read()
                TEMP_ITEM_NM = objDataReader(1)
                'Group_seq,Group_sub,Group 
                'Group_seq
                If IsDBNull(objDataReader(6)) Then
                    TEMP_GROUP_SEQ = ""
                Else
                    TEMP_GROUP_SEQ = objDataReader(6)
                End If
                If OLD_GROUP_SEQ = "" Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                End If

                If OLD_GROUP_SEQ <> TEMP_GROUP_SEQ Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                    '關係企業數歸零
                    'iCompCnt = 0
                End If
                'Group_sub
                If IsDBNull(objDataReader(4)) Then
                    TEMP_GROUP_SUB = ""
                Else
                    TEMP_GROUP_SUB = objDataReader(4)
                End If
                If OLD_GROUP_SUB = "" Then
                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If
                'Group
                If IsDBNull(objDataReader(5)) Then
                    TEMP_GROUP = ""
                Else
                    TEMP_GROUP = objDataReader(5)
                End If
                If OLD_GROUP = "" Then
                    OLD_GROUP = TEMP_GROUP
                End If

                '*** new, GR22_ITEM_NM
                TEMP_ITEM_NM = objDataReader(1)
                If OLD_ITEM_NM <> TEMP_ITEM_NM Then

                    If OLD_ITEM_NM <> "" Then
                        For i = 1 To oCompLists.Count - iCompCnt
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)
                        Next

                        lb_group_sub = False
                        lb_group = False
                        If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                            lb_group_sub = True
                            OLD_GROUP_SUB = TEMP_GROUP_SUB
                        End If
                        If OLD_GROUP <> TEMP_GROUP Then 'Group
                            lb_group = True ' 2 擇 1
                            lb_group_sub = False
                            OLD_GROUP = TEMP_GROUP
                        End If
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = S_GR44_AMT
                        tr_node.appendChild(td_node)
                        S_GR44_AMT = 0
                    End If  'OLD_ITEM_NM <> ""


                    OLD_ITEM_NM = TEMP_ITEM_NM
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "S" & objDataReader(0)
                    tr_node.appendChild(td_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(1)
                    tr_node.appendChild(td_node)
                    iCompCnt = 0
                End If  'OLD_ITEM_NM <> TEMP_ITEM_NM

                '關係企業數 Loop
                For i = iCompCnt To oCompLists.Count - 1
                    iCompCnt = iCompCnt + 1
                    '關係企業 COMP_CD
                    If oCompLists(i)(0) = objDataReader(9) Then
                        'If iCompCnt = 0 Then
                        '    'Item 項目
                        '    tr_node = XMLDoc.createElement("TR")
                        '    thead1_node.appendChild(tr_node)

                        '    td_node = XMLDoc.createElement("TD")
                        '    td_node.text = objDataReader(0)
                        '    tr_node.appendChild(td_node)
                        'End If
                        '金額
                        If Not IsDBNull(objDataReader(7)) And objDataReader(7) = "N" Then  '為 非說明欄 head_yn = 'N'
                            td_node = XMLDoc.createElement("TD")
                            'Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                            td_node.text = objDataReader(3)
                            tr_node.appendChild(td_node)
                            If Not IsDBNull(objDataReader(8)) And objDataReader(8) = "-" Then
                                oCompLists(i)(2) = oCompLists(i)(2) - objDataReader(3) ' Group
                                T_GR44_AMT = T_GR44_AMT - objDataReader(3) '合計
                                S_GR44_AMT += objDataReader(3) '右方合計
                            Else
                                oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(3) ' Group
                                T_GR44_AMT = T_GR44_AMT + objDataReader(3) '合計
                                S_GR44_AMT += objDataReader(3) '右方合計
                            End If
                            oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) ' Group_sub
                            Exit For

                        Else '說明欄
                            td_node = XMLDoc.createelement("TD")
                            If lb_group_sub Then 'Group_sub
                                td_node.text = oCompLists(i)(3)
                                T_GR44_AMT_GROUP_SUB += oCompLists(i)(3)
                                S_GR44_AMT += oCompLists(i)(3) '右方合計
                                oCompLists(i)(3) = 0
                            End If
                            If lb_group Then 'Group
                                td_node.text = oCompLists(i)(2)
                                T_GR44_AMT_GROUP += oCompLists(i)(2)
                                S_GR44_AMT += oCompLists(i)(2) '右方合計
                                oCompLists(i)(3) = 0
                                oCompLists(i)(2) = 0
                            End If
                            tr_node.appendChild(td_node)
                            Exit For


                        End If '為 非說明欄 head_yn = 'N'

                    Else  '自進入LOOP都未找到該公司,表無資料,故直接放0
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = 0
                        tr_node.appendChild(td_node)

                    End If 'END oCompLists(i)(0) = objDataReader(8)

                Next


                ''關係企業數 Loop
                'For i = 0 To oCompLists.Count - 1
                '    '2013/5/3 Modifided by Mikli ;  Ref No : 10234060 
                '    'For i = iCompCnt To oCompLists.Count - 1

                '    '關係企業 COMP_CD
                '    If oCompLists(i)(0) = objDataReader(8) Then
                '        If iCompCnt = 0 Then
                '            'Item 項目
                '            tr_node = XMLDoc.createElement("TR")
                '            thead1_node.appendChild(tr_node)

                '            td_node = XMLDoc.createElement("TD")
                '            td_node.text = objDataReader(0)
                '            tr_node.appendChild(td_node)
                '        End If
                '        iCompCnt = iCompCnt + 1
                '        '金額

                '        td_node = XMLDoc.createElement("TD")
                '        'Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                '        td_node.text = objDataReader(2) '輸入之資料
                '        If Not IsDBNull(objDataReader(7)) And objDataReader(7) = "-" Then
                '            oCompLists(i)(2) = oCompLists(i)(2) - objDataReader(2) ' Group
                '        Else
                '            oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(2) ' Group
                '        End If
                '        oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(2) ' Group_sub
                '        T_GR44_AMT = T_GR44_AMT + objDataReader(2) '合計
                '        If Not IsDBNull(objDataReader(6)) And objDataReader(6) = "Y" Then

                '            If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                '                td_node.text = oCompLists(i)(3)
                '                T_GR44_AMT_GROUP_SUB += oCompLists(i)(3)
                '                oCompLists(i)(3) = 0
                '            End If
                '            If OLD_GROUP <> TEMP_GROUP Then 'Group
                '                td_node.text = oCompLists(i)(2)
                '                T_GR44_AMT_GROUP += oCompLists(i)(2)
                '                oCompLists(i)(2) = 0
                '            End If

                '        End If
                '        tr_node.appendChild(td_node)

                '        'Else '自進入LOOP都未找到該公司,表無資料,故直接放0                    
                '        '    td_node = XMLDoc.createElement("TD")
                '        '    td_node.text = 0 'objDataReader(8) '0
                '        '    tr_node.appendChild(td_node)

                '    End If 'END oCompLists(i)(0) = objDataReader(8)

                'Next


                '項目合計-右方
                'If iCompCnt = oCompLists.Count Then
                '    OLD_GROUP_SEQ = TEMP_GROUP_SEQ

                '    td_node = XMLDoc.createElement("TD")
                '    td_node.text = T_GR44_AMT ' 合計-右方

                '    If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                '        td_node.text = T_GR44_AMT_GROUP_SUB
                '        T_GR44_AMT_GROUP_SUB = 0 '3
                '        OLD_GROUP_SUB = ""
                '    End If
                '    If OLD_GROUP <> TEMP_GROUP Then 'Group
                '        td_node.text = T_GR44_AMT_GROUP
                '        T_GR44_AMT_GROUP = 0 '2
                '        'OLD_GROUP_SUB = TEMP_GROUP_SUB
                '        OLD_GROUP = ""
                '    End If
                '    tr_node.appendChild(td_node)

                '    '關係企業數歸零
                '    iCompCnt = 0
                '    T_GR44_AMT = 0
                'End If
            Loop

            ''項目合計-右方
            For i = 1 To oCompLists.Count - iCompCnt
                td_node = XMLDoc.createElement("TD")
                td_node.text = "0"
                tr_node.appendChild(td_node)
            Next
            td_node = XMLDoc.createElement("TD")
            td_node.text = S_GR44_AMT 'T_GR44_AMT 's_gr44_amt
            tr_node.appendChild(td_node)

        End If


        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "T" Then
            '==================================================
            '******** T 金融負債到期日狀況 GR45  **************
            '==================================================
            sqlStr = ""
            sqlStr = sqlStr & " from GR45, GR22, GR01"
            sqlStr = sqlStr & " where GR45_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR45_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='T'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR45_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR45_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            Dim sCompList '關係企業 List
            sCompList = "select DECODE(GR01_COMP_CD,'00','ZZ',GR01_COMP_CD) as GR01_COMP_CD, GR01_NM_CA " & sqlStr
            sCompList = sCompList & " group by GR01_COMP_CD,GR01_NM_CA"
            sCompList = sCompList & " order by GR01_COMP_CD"
            objDataReader = objDataservice.ExecuteReader(sCompList)


            Dim oCompLists As New ArrayList
            oCompLists.Clear()
            Do While objDataReader.Read()
                Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                oCompLists.Add(oColumn)
            Loop

            ' 建立 TABLE 
            table_node = XMLDoc.createElement("TABLE")
            oRootNode.appendChild(table_node)
            table_node.setAttribute("name", "金融負債到期日狀況彙總")
            'table_node.SetAttribute("border", "1")

            '***************** 建立 表頭
            thead1_node = XMLDoc.createElement("THEAD")
            table_node.appendChild(thead1_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "金融負債到期日狀況彙總"
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目代碼"
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目類別"
            'td_node.setAttribute("rowspan", "2")
            tr_node.appendChild(td_node)

            For i = 0 To oCompLists.Count - 1
                td_node = XMLDoc.createElement("TD")
                td_node.text = oCompLists(i)(1)
                td_node.setAttribute("colspan", "1")
                tr_node.appendChild(td_node)
            Next

            td_node = XMLDoc.createElement("TD")
            td_node.text = "合計"
            td_node.setAttribute("colspan", "1")
            tr_node.appendChild(td_node)

            'sqlStr = "select GR45_ITEM_NM, GR01_NM_CA, sum(nvl(GR45_COST_TWD,0)), sum(nvl(GR45_EXP_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP " & sqlStr
            Dim sqlStrT As String
            sqlStrT = "select GR22_ITEM_CD, GR22_ITEM_NM, GR45_ITEM_NM"
            sqlStrT = sqlStrT & ",GR45_TWD"
            sqlStrT = sqlStrT & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStrT = sqlStrT & ",GR22_GROUP_SEQ,NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP,DECODE(GR01_COMP_CD,'00','ZZ',GR01_COMP_CD) as GR01_COMP_CD "
            sqlStr = sqlStrT & sqlStr


            sqlStr = sqlStr & " UNION "
            sqlStr = sqlStr & "select gr01v_ITEM_CD, gr01v_ITEM_NM, gr01v_ITEM_NM"
            sqlStr = sqlStr & ",gr01v_TWD"
            sqlStr = sqlStr & ",gr01v_GROUP_SUB,gr01v_GROUP"
            sqlStr = sqlStr & ",gr01v_GROUP_SEQ,gr01v_HEAD_YN ,gr01v_GROUP_OP,DECODE(GR01V_COMP_CD,'00','ZZ',GR01V_COMP_CD) as GR01_COMP_CD"
            sqlStr = sqlStr & " from GR01v, "
            sqlStr = sqlStr & " ("
            sqlStr = sqlStr & "select GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & "  from GR45, GR22, GR01"
            sqlStr = sqlStr & " where GR45_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR45_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD = 'T'"
            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR45_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR45_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr = sqlStr & " group by GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & " order by GR01_COMP_CD"
            sqlStr = sqlStr & " ) GR01"

            sqlStr = sqlStr & " where gr01v_DATA_CD = 'T'"
            sqlStr = sqlStr & " AND GR01V_COMP_CD = GR01.GR01_COMP_CD"

            sqlStr = sqlStr & " group by gr01v_ITEM_NM,gr01v_ITEM_NM,gr01v_TWD,gr01v_GROUP_SUB,gr01v_GROUP, gr01v_GROUP_SEQ, gr01v_HEAD_YN, gr01v_GROUP_OP, GR01V_COMP_CD, gr01v_ITEM_CD"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR01_COMP_CD,GR22_GROUP"
            objDataReader = objDataservice.ExecuteReader(sqlStr)

            Dim TEMP_GROUP_SEQ, OLD_GROUP_SEQ As String

            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SEQ = ""

            OLD_GROUP_SEQ = ""
            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SUB = ""
            TEMP_GROUP_SUB = ""
            OLD_GROUP = ""
            TEMP_GROUP = ""
            OLD_ITEM_NM = ""

            Dim S_GR45_AMT As Decimal
            Dim T_GR45_AMT As Decimal
            Dim T_GR45_AMT_GROUP As Decimal
            Dim T_GR45_AMT_GROUP_SUB As Decimal

            Dim iCompCnt ' 關係企業數
            Dim j
            iCompCnt = 0
            iCnt = 1
            j = 0
            T_GR45_AMT_GROUP = 0 'GROUP 合計
            T_GR45_AMT_GROUP_SUB = 0 'GROUP_SUB 合計


            Do While objDataReader.Read()
                TEMP_ITEM_NM = objDataReader(1)
                'Group_seq,Group_sub,Group 
                'Group_seq
                If IsDBNull(objDataReader(6)) Then
                    TEMP_GROUP_SEQ = ""
                Else
                    TEMP_GROUP_SEQ = objDataReader(6)
                End If
                If OLD_GROUP_SEQ = "" Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                End If

                If OLD_GROUP_SEQ <> TEMP_GROUP_SEQ Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                    '關係企業數歸零
                    'iCompCnt = 0
                End If
                'Group_sub
                If IsDBNull(objDataReader(4)) Then
                    TEMP_GROUP_SUB = ""
                Else
                    TEMP_GROUP_SUB = objDataReader(4)
                End If
                If OLD_GROUP_SUB = "" Then
                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If
                'Group
                If IsDBNull(objDataReader(5)) Then
                    TEMP_GROUP = ""
                Else
                    TEMP_GROUP = objDataReader(5)
                End If
                If OLD_GROUP = "" Then
                    OLD_GROUP = TEMP_GROUP
                End If


                '*** new, GR22_ITEM_NM
                TEMP_ITEM_NM = objDataReader(1)
                If OLD_ITEM_NM <> TEMP_ITEM_NM Then

                    If OLD_ITEM_NM <> "" Then
                        For i = 1 To oCompLists.Count - iCompCnt
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)
                        Next

                        lb_group_sub = False
                        lb_group = False
                        If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                            lb_group_sub = True
                            OLD_GROUP_SUB = TEMP_GROUP_SUB
                        End If
                        If OLD_GROUP <> TEMP_GROUP Then 'Group
                            lb_group = True ' 2 擇 1
                            lb_group_sub = False
                            OLD_GROUP = TEMP_GROUP
                        End If
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = S_GR45_AMT
                        tr_node.appendChild(td_node)
                        S_GR45_AMT = 0
                    End If  'OLD_ITEM_NM <> ""


                    OLD_ITEM_NM = TEMP_ITEM_NM
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "T" & objDataReader(0)
                    tr_node.appendChild(td_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(1)
                    tr_node.appendChild(td_node)
                    iCompCnt = 0
                End If  'OLD_ITEM_NM <> TEMP_ITEM_NM

                '關係企業數 Loop
                For i = iCompCnt To oCompLists.Count - 1
                    iCompCnt = iCompCnt + 1
                    '關係企業 COMP_CD
                    If oCompLists(i)(0) = objDataReader(9) Then
                        'If iCompCnt = 0 Then
                        '    'Item 項目
                        '    tr_node = XMLDoc.createElement("TR")
                        '    thead1_node.appendChild(tr_node)

                        '    td_node = XMLDoc.createElement("TD")
                        '    td_node.text = objDataReader(0)
                        '    tr_node.appendChild(td_node)
                        'End If
                        '金額
                        If Not IsDBNull(objDataReader(7)) And objDataReader(7) = "N" Then  '為 非說明欄 head_yn = 'N'
                            td_node = XMLDoc.createElement("TD")
                            'Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                            td_node.text = objDataReader(3)
                            tr_node.appendChild(td_node)
                            If Not IsDBNull(objDataReader(8)) And objDataReader(8) = "-" Then
                                oCompLists(i)(2) = oCompLists(i)(2) - objDataReader(3) ' Group
                                T_GR45_AMT = T_GR45_AMT - objDataReader(3) '合計
                                S_GR45_AMT += objDataReader(3) '右方合計
                            Else
                                oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(3) ' Group
                                T_GR45_AMT = T_GR45_AMT + objDataReader(3) '合計
                                S_GR45_AMT += objDataReader(3) '右方合計
                            End If
                            oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) ' Group_sub
                            Exit For

                        Else '說明欄
                            td_node = XMLDoc.createelement("TD")
                            If lb_group_sub Then 'Group_sub
                                td_node.text = oCompLists(i)(3)
                                T_GR45_AMT_GROUP_SUB += oCompLists(i)(3)
                                S_GR45_AMT += oCompLists(i)(3) '右方合計
                                oCompLists(i)(3) = 0
                            End If
                            If lb_group Then 'Group
                                td_node.text = oCompLists(i)(2)
                                T_GR45_AMT_GROUP += oCompLists(i)(2)
                                S_GR45_AMT += oCompLists(i)(2) '右方合計
                                oCompLists(i)(3) = 0
                                oCompLists(i)(2) = 0
                            End If
                            tr_node.appendChild(td_node)
                            Exit For


                        End If '為 非說明欄 head_yn = 'N'

                    Else  '自進入LOOP都未找到該公司,表無資料,故直接放0
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = 0
                        tr_node.appendChild(td_node)

                    End If 'END oCompLists(i)(0) = objDataReader(8)

                Next

            Loop

            ''項目合計-右方
            For i = 1 To oCompLists.Count - iCompCnt
                td_node = XMLDoc.createElement("TD")
                td_node.text = "0"
                tr_node.appendChild(td_node)
            Next
            td_node = XMLDoc.createElement("TD")
            td_node.text = S_GR45_AMT
            tr_node.appendChild(td_node)


            'If iCompCnt = oCompLists.Count Then
            '    OLD_GROUP_SEQ = TEMP_GROUP_SEQ

            '    td_node = XMLDoc.createElement("TD")
            '    td_node.text = T_GR45_AMT
            '    If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
            '        td_node.text = T_GR45_AMT_GROUP_SUB
            '        T_GR45_AMT_GROUP_SUB = 0 '3
            '        OLD_GROUP_SUB = ""
            '    End If
            '    If OLD_GROUP <> TEMP_GROUP Then 'Group
            '        td_node.text = T_GR45_AMT_GROUP
            '        T_GR45_AMT_GROUP = 0 '2
            '        'OLD_GROUP_SUB = TEMP_GROUP_SUB
            '        OLD_GROUP = ""
            '    End If
            '    tr_node.appendChild(td_node)

            '    '關係企業數歸零
            '    iCompCnt = 0
            '    T_GR45_AMT = 0
            'End If
        End If

        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "U" Then
            '==================================================
            '**************** U 融資租賃 GR46  ****************
            '==================================================
            sqlStr = ""
            sqlStr = sqlStr & " from GR46, GR22, GR01"
            sqlStr = sqlStr & " where GR46_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR46_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='U'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR46_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR46_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            Dim sCompList ' 關係企業 List
            sCompList = "select GR01_COMP_CD , GR01_NM_CA " & sqlStr
            sCompList = sCompList & " group by GR01_COMP_CD,GR01_NM_CA"
            sCompList = sCompList & " order by GR01_COMP_CD"
            objDataReader = objDataservice.ExecuteReader(sCompList)


            Dim oCompLists As New ArrayList
            oCompLists.Clear()
            Do While objDataReader.Read()
                Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                oCompLists.Add(oColumn)
            Loop

            ' 建立 TABLE 
            table_node = XMLDoc.createElement("TABLE")
            oRootNode.appendChild(table_node)
            table_node.setAttribute("name", "融資租賃彙總")
            'table_node.SetAttribute("border", "1")

            '***************** 建立 表頭
            thead1_node = XMLDoc.createElement("THEAD")
            table_node.appendChild(thead1_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "融資租賃彙總"
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
                td_node.setAttribute("colspan", "1")
                tr_node.appendChild(td_node)
            Next

            td_node = XMLDoc.createElement("TD")
            td_node.text = "合計"
            td_node.setAttribute("colspan", "1")
            tr_node.appendChild(td_node)

            'sqlStr = "select GR46_ITEM_NM, GR01_NM_CA, sum(nvl(GR46_COST_TWD,0)), sum(nvl(GR46_EXP_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP " & sqlStr
            Dim sqlStrT As String
            sqlStrT = "select GR22_ITEM_NM, GR46_ITEM_NM"
            sqlStrT = sqlStrT & ",GR46_TWD"
            sqlStrT = sqlStrT & ",GR22_GROUP_SUB,GR22_GROUP"
            sqlStrT = sqlStrT & ",GR22_GROUP_SEQ,NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP,GR01_COMP_CD"
            sqlStr = sqlStrT & sqlStr


            sqlStr = sqlStr & " UNION "
            sqlStr = sqlStr & "select gr01v_ITEM_NM, gr01v_ITEM_NM"
            sqlStr = sqlStr & ",gr01v_TWD"
            sqlStr = sqlStr & ",gr01v_GROUP_SUB,gr01v_GROUP"
            sqlStr = sqlStr & ",gr01v_GROUP_SEQ,gr01v_HEAD_YN ,gr01v_GROUP_OP,gr01v_COMP_CD"
            sqlStr = sqlStr & " from GR01v, "
            sqlStr = sqlStr & " ("
            sqlStr = sqlStr & "select GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & "  from GR46, GR22, GR01"
            sqlStr = sqlStr & " where GR46_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR46_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD = 'U'"
            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR46_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR46_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr = sqlStr & " group by GR01_COMP_CD, GR01_NM_CA"
            sqlStr = sqlStr & " order by GR01_COMP_CD"
            sqlStr = sqlStr & " ) GR01"

            sqlStr = sqlStr & " where gr01v_DATA_CD = 'U'"
            sqlStr = sqlStr & " AND GR01V_COMP_CD = GR01.GR01_COMP_CD"

            sqlStr = sqlStr & " group by gr01v_ITEM_NM,gr01v_ITEM_NM,gr01v_TWD,gr01v_GROUP_SUB,gr01v_GROUP, gr01v_GROUP_SEQ, gr01v_HEAD_YN, gr01v_GROUP_OP, GR01V_COMP_CD"
            sqlStr = sqlStr & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR01_COMP_CD,GR22_GROUP"
            objDataReader = objDataservice.ExecuteReader(sqlStr)

            Dim TEMP_GROUP_SEQ, OLD_GROUP_SEQ As String

            OLD_GROUP_SEQ = ""
            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SUB = ""
            TEMP_GROUP_SUB = ""
            OLD_GROUP = ""
            TEMP_GROUP = ""
            OLD_ITEM_NM = ""

            Dim S_GR46_AMT As Decimal
            Dim T_GR46_AMT As Decimal
            Dim T_GR46_AMT_GROUP As Decimal
            Dim T_GR46_AMT_GROUP_SUB As Decimal

            Dim iCompCnt '關係企業數
            iCompCnt = 0
            iCnt = 1
            T_GR46_AMT_GROUP = 0 'GROUP 合計
            T_GR46_AMT_GROUP_SUB = 0 'GROUP_SUB 合計


            Do While objDataReader.Read()
                TEMP_ITEM_NM = objDataReader(0)
                'Group_seq,Group_sub,Group 
                'Group_seq
                If IsDBNull(objDataReader(5)) Then
                    TEMP_GROUP_SEQ = ""
                Else
                    TEMP_GROUP_SEQ = objDataReader(5)
                End If
                If OLD_GROUP_SEQ = "" Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                End If

                If OLD_GROUP_SEQ <> TEMP_GROUP_SEQ Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                    '關係企業數歸零
                    'iCompCnt = 0
                End If
                'Group_sub
                If IsDBNull(objDataReader(3)) Then
                    TEMP_GROUP_SUB = ""
                Else
                    TEMP_GROUP_SUB = objDataReader(3)
                End If
                If OLD_GROUP_SUB = "" Then
                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If
                'Group
                If IsDBNull(objDataReader(4)) Then
                    TEMP_GROUP = ""
                Else
                    TEMP_GROUP = objDataReader(4)
                End If
                If OLD_GROUP = "" Then
                    OLD_GROUP = TEMP_GROUP
                End If


                '*** new, GR22_ITEM_NM
                TEMP_ITEM_NM = objDataReader(0)
                If OLD_ITEM_NM <> TEMP_ITEM_NM Then

                    If OLD_ITEM_NM <> "" Then
                        For i = 1 To oCompLists.Count - iCompCnt
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)
                        Next

                        lb_group_sub = False
                        lb_group = False
                        If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                            lb_group_sub = True
                            OLD_GROUP_SUB = TEMP_GROUP_SUB
                        End If
                        If OLD_GROUP <> TEMP_GROUP Then 'Group
                            lb_group = True ' 2 擇 1
                            lb_group_sub = False
                            OLD_GROUP = TEMP_GROUP
                        End If
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = S_GR46_AMT
                        tr_node.appendChild(td_node)
                        S_GR46_AMT = 0
                    End If  'OLD_ITEM_NM <> ""


                    OLD_ITEM_NM = TEMP_ITEM_NM
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(0)
                    tr_node.appendChild(td_node)
                    iCompCnt = 0
                End If  'OLD_ITEM_NM <> TEMP_ITEM_NM


                '關係企業數 Loop
                For i = iCompCnt To oCompLists.Count - 1
                    iCompCnt = iCompCnt + 1
                    '關係企業 COMP_CD
                    If oCompLists(i)(0) = objDataReader(8) Then
                        'If iCompCnt = 0 Then
                        '    'Item 項目
                        '    tr_node = XMLDoc.createElement("TR")
                        '    thead1_node.appendChild(tr_node)

                        '    td_node = XMLDoc.createElement("TD")
                        '    td_node.text = objDataReader(0)
                        '    tr_node.appendChild(td_node)
                        'End If                    
                        '金額

                        If Not IsDBNull(objDataReader(6)) And objDataReader(6) = "N" Then  '為 非說明欄 head_yn = 'N'
                            td_node = XMLDoc.createElement("TD")
                            'Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0} 'GR01_COMP_CD,GR01_NM_CA,group,group_sub 對應位置=[0,1,2,3]
                            td_node.text = objDataReader(2)
                            tr_node.appendChild(td_node)
                            If Not IsDBNull(objDataReader(7)) And objDataReader(7) = "-" Then
                                oCompLists(i)(2) = oCompLists(i)(2) - objDataReader(2) ' Group
                                oCompLists(i)(3) = oCompLists(i)(3) - objDataReader(2) ' Group_sub
                                T_GR46_AMT = T_GR46_AMT - objDataReader(2) '合計
                                S_GR46_AMT -= objDataReader(2) '右方合計
                            Else
                                oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(2) ' Group
                                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(2) ' Group_sub
                                T_GR46_AMT = T_GR46_AMT + objDataReader(2) '合計
                                S_GR46_AMT += objDataReader(2) '右方合計
                            End If
                            Exit For

                        Else

                            td_node = XMLDoc.createelement("TD")
                            If lb_group_sub Then 'Group_sub
                                td_node.text = oCompLists(i)(3)
                                T_GR46_AMT_GROUP_SUB += oCompLists(i)(3)
                                S_GR46_AMT += oCompLists(i)(3) '右方合計
                                oCompLists(i)(3) = 0
                            End If
                            If lb_group Then 'Group
                                td_node.text = oCompLists(i)(2)
                                T_GR46_AMT_GROUP += oCompLists(i)(2)
                                S_GR46_AMT += oCompLists(i)(2) '右方合計
                                oCompLists(i)(2) = 0
                                oCompLists(i)(3) = 0
                            End If
                            tr_node.appendChild(td_node)
                            Exit For

                        End If  'Not IsDBNull(objDataReader(6)) And objDataReader(6) = "N" Then  '為 非說明欄 head_yn = 'N'
                    Else  '自進入LOOP都未找到該公司,表無資料,故直接放0
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = 0
                        tr_node.appendChild(td_node)

                    End If 'END oCompLists(i)(0) = objDataReader(8)

                Next

            Loop

            For i = 1 To oCompLists.Count - iCompCnt
                td_node = XMLDoc.createElement("TD")
                td_node.text = "0"
                tr_node.appendChild(td_node)
            Next
            td_node = XMLDoc.createElement("TD")
            td_node.text = S_GR46_AMT
            tr_node.appendChild(td_node)

        End If


        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "V" Then
            '==================================================
            '************** V 投資性不動產彙總 ****************
            '==================================================

            'sqlStr = "select decode(nvl(GR47_ITEM_NM,'0'),'0',GR22_ITEM_NM,decode(rtrim(ltrim(GR47_ITEM_NM)),'',GR22_ITEM_NM,GR47_ITEM_NM)) as GR47_ITEM_NM, GR47_LAND_TWD, GR47_BUILDING_TWD"
            sqlStr = ""
            'sqlStr = "select GR47_ITEM_NM, GR01_NM_CA, sum(nvl(GR47_LAND_TWD,0)), sum(nvl(GR47_BUILDING_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR47, GR22, GR01"
            sqlStr = sqlStr & " where GR47_ITEM_CD = GR22_ITEM_CD"
            sqlStr = sqlStr & "   and GR47_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='V'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR47_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR47_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            Dim sCompList
            sCompList = ""
            sCompList = "select GR01_COMP_CD , GR01_NM_CA " & sqlStr
            sCompList = sCompList & " group by GR01_COMP_CD,GR01_NM_CA"
            sCompList = sCompList & " order by GR01_COMP_CD"
            objDataReader = objDataservice.ExecuteReader(sCompList)

            Dim oCompLists As New ArrayList
            oCompLists = New ArrayList
            Do While objDataReader.Read()
                'GR01_COMP_CD(0),GR01_NM_CA(1),土地group(2),土地group_sub(3),建築物group(4),建築物group_sub(5) 對應位置=[0,1,2,3,4,5]
                Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0, 0, 0}
                oCompLists.Add(oColumn)
            Loop

            ' 建立 TABLE 
            table_node = XMLDoc.createElement("TABLE")
            oRootNode.appendChild(table_node)
            table_node.setAttribute("name", "投資性不動產彙總")
            'table_node.SetAttribute("border", "1")

            '***************** 建立 表頭
            thead1_node = XMLDoc.createElement("THEAD")
            table_node.appendChild(thead1_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "投資性不動產彙總"
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目代碼"
            tr_node.appendChild(td_node)

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
            td_node = XMLDoc.createElement("TD")
            td_node.text = ""
            tr_node.appendChild(td_node)

            For i = 0 To oCompLists.Count
                td_node = XMLDoc.createElement("TD")
                td_node.text = "土地"
                tr_node.appendChild(td_node)
                td_node = XMLDoc.createElement("TD")
                td_node.text = "建築物"
                tr_node.appendChild(td_node)
            Next

            '                  0             1                             2(土地)    3(建築物) 4  5  6              7              8
            sqlStr1 = " select GR22_ITEM_CD, GR22_ITEM_NM as GR47_ITEM_NM, GR01_NM_CA, 0,0,GR22_GROUP,GR22_GROUP_SUB,GR22_GROUP_SEQ,GR01_COMP_CD, "
            '                       9                                        10
            sqlStr1 += " NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP "
            sqlStr1 += " from gr22, "

            sqlStr1 += " (select GR01_COMP_CD , GR01_NM_CA  from GR47, GR22, GR01 "
            sqlStr1 += "where GR47_ITEM_CD = GR22_ITEM_CD (+)    and GR47_COMP_CD = GR01_COMP_CD "
            sqlStr1 += " and GR22_DATA_CD ='V' "
            ' AND GR47_DATA_YY = '2012' AND GR47_SEASON_CD = '1' "
            ' 2013/04/25 Mikli Modify
            If GRXX_DATA_YY <> "" Then
                sqlStr1 += " AND GR47_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr1 += " AND GR47_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If
            sqlStr1 += "group by GR01_COMP_CD,GR01_NM_CA order by GR01_COMP_CD) gr01 "

            sqlStr1 += " where GR22_DATA_CD ='V' and nvl(gr22_head_yn,'N') = 'Y' "
            sqlStr1 += " UNION ALL "
            sqlStr1 += " select GR22_ITEM_CD, GR47_ITEM_NM, GR01_NM_CA, sum(nvl(GR47_LAND_TWD,0)), sum(nvl(GR47_BUILDING_TWD,0)) ,GR22_GROUP,GR22_GROUP_SUB,GR22_GROUP_SEQ,GR01_COMP_CD,NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP "
            sqlStr1 += sqlStr
            sqlStr1 = sqlStr1 & " group by GR47_ITEM_NM,GR01_NM_CA,GR22_GROUP_SEQ,GR22_GROUP,GR22_GROUP_SUB,GR01_COMP_CD,GR22_HEAD_YN,GR22_GROUP_OP,GR22_ITEM_CD"
            sqlStr1 = sqlStr1 & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR01_COMP_CD,GR22_GROUP"

            objDataReader = objDataservice.ExecuteReader(sqlStr1)

            '***
            Dim TEMP_GROUP_SEQ, OLD_GROUP_SEQ As String

            OLD_GROUP_SEQ = ""
            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SUB = ""
            TEMP_GROUP_SUB = ""
            OLD_GROUP = ""
            TEMP_GROUP = ""
            OLD_ITEM_NM = ""
            TEMP_ITEM_NM = ""


            Dim S_GR47_AMT As Decimal
            Dim T_GR47_AMT As Decimal
            Dim S_GR47_LAND_AMT, S_GR47_BUILDING_AMT As Decimal
            Dim T_GR47_LAND_AMT, T_GR47_BUILDING_AMT As Decimal
            Dim iCompCnt
            Dim j
            Dim ld_data

            iCompCnt = 0
            iCnt = 1
            j = 0
            ld_data = 0 '暫存資料
            lb_group = False
            lb_group_sub = False
            ' 讀取sqlStr資料
            Do While objDataReader.Read()
                '*** NEW
                'Group_seq(4),Group_sub(5),Group(6) 
                'Group_seq
                If IsDBNull(objDataReader(7)) Then
                    TEMP_GROUP_SEQ = ""
                Else
                    TEMP_GROUP_SEQ = objDataReader(7)
                End If
                If OLD_GROUP_SEQ = "" Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                End If
                If OLD_GROUP_SEQ <> TEMP_GROUP_SEQ Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                    '關係企業數歸零
                    'iCompCnt = 0
                End If
                'Group_sub
                If IsDBNull(objDataReader(6)) Then
                    TEMP_GROUP_SUB = ""
                Else
                    TEMP_GROUP_SUB = objDataReader(6)
                End If
                If OLD_GROUP_SUB = "" Then
                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If
                'Group
                If IsDBNull(objDataReader(5)) Then
                    TEMP_GROUP = ""
                Else
                    TEMP_GROUP = objDataReader(5)
                End If
                If OLD_GROUP = "" Then
                    OLD_GROUP = TEMP_GROUP
                End If
                'END NEW

                TEMP_ITEM_NM = objDataReader(1)
                If OLD_ITEM_NM <> TEMP_ITEM_NM Then

                    If OLD_ITEM_NM <> "" Then
                        For i = 1 To oCompLists.Count - iCompCnt
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)
                        Next

                        lb_group_sub = False
                        lb_group = False
                        If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                            lb_group_sub = True
                            OLD_GROUP_SUB = TEMP_GROUP_SUB
                        End If
                        If OLD_GROUP <> TEMP_GROUP Then 'Group
                            lb_group = True ' 2 擇 1
                            lb_group_sub = False
                            OLD_GROUP = TEMP_GROUP
                        End If

                        td_node = XMLDoc.createElement("TD")
                        td_node.text = T_GR47_LAND_AMT ' 合計 橫向資料
                        tr_node.appendChild(td_node)
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = T_GR47_BUILDING_AMT
                        tr_node.appendChild(td_node)
                        T_GR47_LAND_AMT = 0
                        T_GR47_BUILDING_AMT = 0
                    End If  'OLD_ITEM_NM <> ""


                    OLD_ITEM_NM = TEMP_ITEM_NM
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "V" & objDataReader(0)
                    tr_node.appendChild(td_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(1)
                    tr_node.appendChild(td_node)
                    iCompCnt = 0
                End If

                For i = iCompCnt To oCompLists.Count - 1
                    iCompCnt = iCompCnt + 1
                    If oCompLists(i)(1) = objDataReader(2) Then
                        '處理說明欄位與非說明欄位
                        If Not IsDBNull(objDataReader(9)) And objDataReader(9) = "N" Then  '為 非說明欄 head_yn = 'N' 
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = objDataReader(3)
                            tr_node.appendChild(td_node)
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = objDataReader(4)
                            tr_node.appendChild(td_node)
                            If Not IsDBNull(objDataReader(10)) And objDataReader(10) = "-" Then
                                oCompLists(i)(2) = oCompLists(i)(2) - objDataReader(3) '土地 group
                                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) '土地 subgroup

                                oCompLists(i)(4) = oCompLists(i)(4) - objDataReader(4) '建築物 group
                                oCompLists(i)(5) = oCompLists(i)(5) + objDataReader(4) '建築物 subgroup
                            Else
                                oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(3) '土地 group
                                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) '土地 subgroup

                                oCompLists(i)(4) = oCompLists(i)(4) + objDataReader(4) '建築物 group
                                oCompLists(i)(5) = oCompLists(i)(5) + objDataReader(4) '建築物 subgroup

                            End If



                            T_GR47_LAND_AMT = T_GR47_LAND_AMT + objDataReader(3)
                            T_GR47_BUILDING_AMT = T_GR47_BUILDING_AMT + objDataReader(4)
                            Exit For
                        Else  '說明欄處理


                            If lb_group_sub Then 'Group_sub

                                T_GR47_LAND_AMT = T_GR47_LAND_AMT + oCompLists(i)(3)
                                T_GR47_BUILDING_AMT = T_GR47_BUILDING_AMT + oCompLists(i)(5)

                                td_node = XMLDoc.createElement("TD")    '土地 Group_sub
                                td_node.text = oCompLists(i)(3) '成本-期末餘額
                                oCompLists(i)(3) = 0
                                tr_node.appendChild(td_node)


                                td_node = XMLDoc.createElement("TD")    '建築物 Group_sub
                                td_node.text = oCompLists(i)(5)
                                oCompLists(i)(5) = 0
                                tr_node.appendChild(td_node)

                            End If
                            If lb_group Then 'Group

                                T_GR47_LAND_AMT = T_GR47_LAND_AMT + oCompLists(i)(2)
                                T_GR47_BUILDING_AMT = T_GR47_BUILDING_AMT + oCompLists(i)(4)


                                td_node = XMLDoc.createElement("TD")   ' 土地 Group
                                td_node.text = oCompLists(i)(2)
                                '2013/04/26 Modified by Mikli ; Remark
                                oCompLists(i)(2) = 0 ' 不歸0
                                tr_node.appendChild(td_node)

                                td_node = XMLDoc.createElement("TD") ' 建築物 Group
                                td_node.text = oCompLists(i)(4)
                                '2013/04/26 Modified by Mikli ; Remark
                                oCompLists(i)(4) = 0 ' 不歸0
                                tr_node.appendChild(td_node)
                            End If
                            Exit For



                        End If
                    Else
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = "13141919"
                        tr_node.appendChild(td_node)
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = "0"
                        tr_node.appendChild(td_node)
                    End If
                Next
            Loop
            '右下方合計
            For i = 1 To oCompLists.Count - iCompCnt
                td_node = XMLDoc.createElement("TD")
                td_node.text = "0"
                tr_node.appendChild(td_node)
                td_node = XMLDoc.createElement("TD")
                td_node.text = "0"
                tr_node.appendChild(td_node)
            Next
            td_node = XMLDoc.createElement("TD")
            td_node.text = T_GR47_LAND_AMT
            tr_node.appendChild(td_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = T_GR47_BUILDING_AMT
            ' for testing
            'td_node.text = sqlStr1  
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

            'Dim S_GR47_LAND_AMT, S_GR47_BUILDING_AMT As Decimal
            'Dim T_GR47_LAND_AMT, T_GR47_BUILDING_AMT As Decimal

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
            '        td_node.text = S_GR47_LAND_AMT

            '        td_node = XMLDoc.createElement("TD")
            '        tr_node.appendChild(td_node)
            '        td_node.text = S_GR47_BUILDING_AMT

            '        '小計歸零
            '        S_GR47_LAND_AMT = 0
            '        S_GR47_BUILDING_AMT = 0


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
            '        td_node.text = T_GR47_LAND_AMT

            '        td_node = XMLDoc.createElement("TD")
            '        tr_node.appendChild(td_node)
            '        td_node.text = T_GR47_BUILDING_AMT

            '        '小計歸零
            '        T_GR47_LAND_AMT = 0
            '        T_GR47_BUILDING_AMT = 0

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
            '                    S_GR47_LAND_AMT = S_GR47_LAND_AMT + 0
            '                    T_GR47_LAND_AMT = T_GR47_LAND_AMT + 0
            '                Else
            '                    td_node.text = objDataReader(i)
            '                    If OLD_GROUP_SUB <> "" Then
            '                        S_GR47_LAND_AMT = S_GR47_LAND_AMT + objDataReader(i)
            '                    End If
            '                    T_GR47_LAND_AMT = T_GR47_LAND_AMT + objDataReader(i)

            '                End If
            '            Case 3
            '                td_node = XMLDoc.createElement("TD")
            '                tr_node.appendChild(td_node)
            '                If IsDBNull(objDataReader(i)) Then
            '                    td_node.text = ""
            '                    S_GR47_BUILDING_AMT = S_GR47_BUILDING_AMT + 0
            '                    T_GR47_BUILDING_AMT = T_GR47_BUILDING_AMT + 0
            '                Else
            '                    td_node.text = objDataReader(i)

            '                    If OLD_GROUP_SUB <> "" Then
            '                        S_GR47_BUILDING_AMT = S_GR47_BUILDING_AMT + objDataReader(i)
            '                    End If
            '                    T_GR47_BUILDING_AMT = T_GR47_BUILDING_AMT + objDataReader(i)

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
            '    td_node.text = S_GR47_LAND_AMT

            '    td_node = XMLDoc.createElement("TD")
            '    tr_node.appendChild(td_node)
            '    td_node.text = S_GR47_BUILDING_AMT
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
            'td_node.text = T_GR47_LAND_AMT

            'td_node = XMLDoc.createElement("TD")
            'tr_node.appendChild(td_node)
            'td_node.text = T_GR47_BUILDING_AMT
        End If


        ' 30DEC2019 BruceWang (081075)增加欄位 項目代碼
        If GR18_DATA_CD = "ALL" Or GR18_DATA_CD = "W" Then
            '==================================================
            '***************** W 無形資產彙總 *****************
            '==================================================

            'sqlStr = "select decode(nvl(GR48_ITEM_NM,'0'),'0',GR22_ITEM_NM,decode(rtrim(ltrim(GR48_ITEM_NM)),'',GR22_ITEM_NM,GR48_ITEM_NM)) as GR48_ITEM_NM, GR48_GOODWILL_TWD, GR48_LICENSE_TWD"
            sqlStr = ""
            'sqlStr = "select GR48_ITEM_NM, GR01_NM_CA, sum(nvl(GR48_GOODWILL_TWD,0)), sum(nvl(GR48_LICENSE_TWD,0)) ,GR22_GROUP_SUB,GR22_GROUP"
            sqlStr = sqlStr & " from GR48, GR22, GR01"
            sqlStr = sqlStr & " where GR48_ITEM_CD = GR22_ITEM_CD (+) "
            sqlStr = sqlStr & "   and GR48_COMP_CD = GR01_COMP_CD"
            sqlStr = sqlStr & "   and GR22_DATA_CD ='W'"

            If GRXX_DATA_YY <> "" Then
                sqlStr = sqlStr & " AND GR48_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr = sqlStr & " AND GR48_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            Dim sCompList
            sCompList = ""
            sCompList = "select GR01_COMP_CD , GR01_NM_CA " & sqlStr
            sCompList = sCompList & " group by GR01_COMP_CD,GR01_NM_CA"
            sCompList = sCompList & " order by GR01_COMP_CD"
            objDataReader = objDataservice.ExecuteReader(sCompList)

            Dim oCompLists As New ArrayList
            oCompLists = New ArrayList
            Do While objDataReader.Read()
                'GR01_COMP_CD(0),GR01_NM_CA(1),商譽group(2),商譽group_sub(3),軟體group(4),軟體group_sub(5),專利group(6),專利group_sub(7) 對應位置=[0,1,2,3,4,5,6,7]
                Dim oColumn() As Object = {objDataReader(0), objDataReader(1), 0, 0, 0, 0, 0, 0}
                oCompLists.Add(oColumn)
            Loop

            ' 建立 TABLE 
            table_node = XMLDoc.createElement("TABLE")
            oRootNode.appendChild(table_node)
            table_node.setAttribute("name", "無形資產彙總")
            'table_node.SetAttribute("border", "1")

            '***************** 建立 表頭
            thead1_node = XMLDoc.createElement("THEAD")
            table_node.appendChild(thead1_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = "無形資產彙總"
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目代碼"
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = "項目類別"
            'td_node.setAttribute("rowspan", "2")
            tr_node.appendChild(td_node)

            For i = 0 To oCompLists.Count - 1
                td_node = XMLDoc.createElement("TD")
                td_node.text = oCompLists(i)(1)
                td_node.setAttribute("colspan", "3")
                tr_node.appendChild(td_node)
            Next

            td_node = XMLDoc.createElement("TD")
            td_node.text = "合計"
            td_node.setAttribute("colspan", "3")
            tr_node.appendChild(td_node)

            tr_node = XMLDoc.createElement("TR")
            thead1_node.appendChild(tr_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = ""
            tr_node.appendChild(td_node)
            td_node = XMLDoc.createElement("TD")
            td_node.text = ""
            tr_node.appendChild(td_node)

            For i = 0 To oCompLists.Count
                td_node = XMLDoc.createElement("TD")
                td_node.text = "商譽"
                tr_node.appendChild(td_node)

                td_node = XMLDoc.createElement("TD")
                td_node.text = "電腦軟體成本"
                tr_node.appendChild(td_node)

                td_node = XMLDoc.createElement("TD")
                td_node.text = "專利及執照"
                tr_node.appendChild(td_node)
            Next

            '                            0                      1           2(商譽) 3(軟體) 4(專利)   5   6       7             8            9
            sqlStr1 = " select GR22_ITEM_CD, GR22_ITEM_NM as GR48_ITEM_NM, GR01_NM_CA, 0,0,0 ,GR22_GROUP,GR22_GROUP_SUB,GR22_GROUP_SEQ,GR01_COMP_CD, "
            '                      10                                        11
            sqlStr1 += " NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP "
            sqlStr1 += " from gr22, "

            sqlStr1 += " (select GR01_COMP_CD , GR01_NM_CA  from GR48, GR22, GR01 "
            sqlStr1 += "where GR48_ITEM_CD = GR22_ITEM_CD (+)    and GR48_COMP_CD = GR01_COMP_CD "
            sqlStr1 += " and GR22_DATA_CD ='W' "
            ' 2013/04/22 Mikli Modify
            If GRXX_DATA_YY <> "" Then
                sqlStr1 += " AND GR48_DATA_YY = '" & GRXX_DATA_YY & "'"
            End If
            If GRXX_SEASON_CD <> "BLANK" Then
                sqlStr1 += " AND GR48_SEASON_CD = '" & GRXX_SEASON_CD & "'"
            End If

            sqlStr1 += "group by GR01_COMP_CD,GR01_NM_CA order by GR01_COMP_CD) gr01 "

            sqlStr1 += " where GR22_DATA_CD ='W' and nvl(gr22_head_yn,'N') = 'Y' "
            sqlStr1 += " UNION ALL "
            sqlStr1 += " select GR22_ITEM_CD, GR48_ITEM_NM, GR01_NM_CA, sum(nvl(GR48_GOODWILL_TWD,0)),sum(nvl(GR48_SOFTWARE_AMT,0)), sum(nvl(GR48_LICENSE_TWD,0)) ,GR22_GROUP,GR22_GROUP_SUB,GR22_GROUP_SEQ,GR01_COMP_CD, "
            sqlStr1 += " NVL(gr22_head_yn,'N') as GR22_HEAD_YN ,NVL(gr22_group_op,'+') as GR22_GROUP_OP "
            sqlStr1 += sqlStr

            sqlStr1 = sqlStr1 & " group by GR48_ITEM_NM,GR01_NM_CA,GR22_GROUP_SEQ,GR22_GROUP_SUB,GR22_GROUP,GR01_COMP_CD,GR22_HEAD_YN,GR22_GROUP_OP,GR22_ITEM_CD"
            sqlStr1 = sqlStr1 & " order by GR22_GROUP_SEQ,GR22_GROUP_SUB,GR01_COMP_CD,GR22_GROUP"
            objDataReader = objDataservice.ExecuteReader(sqlStr1)

            Dim S_GR48_GOODWILL_AMT, S_GR48_SOFTWARE_AMT, S_GR48_LICENSE_AMT As Decimal
            Dim T_GR48_GOODWILL_AMT, T_GR48_SOFTWARE_AMT, T_GR48_LICENSE_AMT As Decimal

            '***
            Dim TEMP_GROUP_SEQ, OLD_GROUP_SEQ As String

            OLD_GROUP_SEQ = ""
            TEMP_GROUP_SEQ = ""
            OLD_GROUP_SUB = ""
            TEMP_GROUP_SUB = ""
            OLD_GROUP = ""
            TEMP_GROUP = ""
            OLD_ITEM_NM = ""
            TEMP_ITEM_NM = ""

            Dim S_GR48_TWD As Decimal
            Dim T_GR48_AMT As Decimal
            Dim T_GR48_AMT_GROUP As Decimal
            Dim T_GR48_AMT_GROUP_SUB As Decimal
            'initial
            T_GR48_AMT_GROUP = 0 'GROUP 合計
            T_GR48_AMT_GROUP_SUB = 0 'GROUP_SUB 合計

            Dim iCompCnt '(關係企業數)
            Dim j
            Dim ld_data

            iCompCnt = 0
            iCnt = 1
            j = 0
            ld_data = 0 '暫存資料
            '*** new
            ' 讀取sqlStr資料
            Do While objDataReader.Read()
                'Group_seq(7),Group_sub(6),Group(5) 
                'Group_seq
                If IsDBNull(objDataReader(8)) Then
                    TEMP_GROUP_SEQ = ""
                Else
                    TEMP_GROUP_SEQ = objDataReader(8)
                End If
                If OLD_GROUP_SEQ = "" Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                End If
                If OLD_GROUP_SEQ <> TEMP_GROUP_SEQ Then
                    OLD_GROUP_SEQ = TEMP_GROUP_SEQ
                    '關係企業數歸零
                    'iCompCnt = 0
                End If
                'Group_sub
                If IsDBNull(objDataReader(7)) Then
                    TEMP_GROUP_SUB = ""
                Else
                    TEMP_GROUP_SUB = objDataReader(7)
                End If
                If OLD_GROUP_SUB = "" Then
                    OLD_GROUP_SUB = TEMP_GROUP_SUB
                End If
                'Group
                If IsDBNull(objDataReader(6)) Then
                    TEMP_GROUP = ""
                Else
                    TEMP_GROUP = objDataReader(6)
                End If
                If OLD_GROUP = "" Then
                    OLD_GROUP = TEMP_GROUP
                End If
                '    '
                '    '*** new
                '    '關係企業數 Loop
                '    For i = 0 To oCompLists.Count - 1

                '        '關係企業 COMP_CD
                '        If oCompLists(i)(0) = objDataReader(8) Then
                '            If iCompCnt = 0 Then
                '                'Item 項目
                '                tr_node = XMLDoc.createElement("TR")
                '                thead1_node.appendChild(tr_node)

                '                td_node = XMLDoc.createElement("TD")
                '                td_node.text = objDataReader(0)
                '                tr_node.appendChild(td_node)
                '            End If
                '            iCompCnt = iCompCnt + 1

                '            '寫下商譽(2)金額{2,3}
                '            If Not IsDBNull(objDataReader(9)) And objDataReader(9) = "N" Then  '為 非說明欄 head_yn = 'N'
                '                td_node = XMLDoc.createElement("TD")
                '                td_node.text = objDataReader(2)
                '                oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(2) ' Group
                '                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(2) ' Group_sub
                '                T_GR48_AMT = T_GR48_AMT + objDataReader(2) '合計

                '            Else
                '                If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                '                    td_node.text = oCompLists(i)(3)
                '                    T_GR48_AMT_GROUP_SUB += oCompLists(i)(3)
                '                    oCompLists(i)(3) = 0
                '                End If
                '                If OLD_GROUP <> TEMP_GROUP Then 'Group
                '                    td_node.text = oCompLists(i)(2)
                '                    T_GR48_AMT_GROUP += oCompLists(i)(2)
                '                    oCompLists(i)(2) = 0
                '                End If

                '            End If
                '            tr_node.appendChild(td_node)

                '            '寫下軟體(3)金額{4,5}
                '            If Not IsDBNull(objDataReader(9)) And objDataReader(9) = "N" Then  '為 非說明欄 head_yn = 'N'
                '                td_node = XMLDoc.createElement("TD")
                '                td_node.text = objDataReader(3)
                '                oCompLists(i)(4) = oCompLists(i)(4) + objDataReader(3) ' Group
                '                oCompLists(i)(5) = oCompLists(i)(5) + objDataReader(3) ' Group_sub
                '                T_GR48_AMT = T_GR48_AMT + objDataReader(3) '合計

                '            Else
                '                If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                '                    td_node.text = oCompLists(i)(5)
                '                    T_GR48_AMT_GROUP_SUB += oCompLists(i)(5)
                '                    oCompLists(i)(5) = 0
                '                End If
                '                If OLD_GROUP <> TEMP_GROUP Then 'Group
                '                    td_node.text = oCompLists(i)(4)
                '                    T_GR48_AMT_GROUP += oCompLists(i)(4)
                '                    oCompLists(i)(4) = 0
                '                End If

                '            End If
                '            tr_node.appendChild(td_node)

                '            '寫下軟體(4)金額{6,7}
                '            If Not IsDBNull(objDataReader(9)) And objDataReader(9) = "N" Then  '為 非說明欄 head_yn = 'N'
                '                td_node = XMLDoc.createElement("TD")
                '                td_node.text = objDataReader(4)
                '                oCompLists(i)(6) = oCompLists(i)(6) + objDataReader(4) ' Group
                '                oCompLists(i)(7) = oCompLists(i)(7) + objDataReader(4) ' Group_sub
                '                T_GR48_AMT = T_GR48_AMT + objDataReader(4) '合計

                '            Else
                '                If OLD_GROUP_SUB <> TEMP_GROUP_SUB Then 'Group_sub
                '                    td_node.text = oCompLists(i)(7)
                '                    T_GR48_AMT_GROUP_SUB += oCompLists(i)(7)
                '                    oCompLists(i)(7) = 0
                '                End If
                '                If OLD_GROUP <> TEMP_GROUP Then 'Group
                '                    td_node.text = oCompLists(i)(6)
                '                    T_GR48_AMT_GROUP += oCompLists(i)(6)
                '                    oCompLists(i)(6) = 0
                '                End If

                '            End If
                '            tr_node.appendChild(td_node)

                '        End If 'END oCompLists(i)(0) = objDataReader(8)

                '    Next
                '    iCompCnt = 0 ' 跑下一項目


                '*** old, GR22_ITEM_NM
                TEMP_ITEM_NM = objDataReader(1)
                If OLD_ITEM_NM <> TEMP_ITEM_NM Then

                    If OLD_ITEM_NM <> "" Then
                        For i = 1 To oCompLists.Count - iCompCnt
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)

                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)

                            td_node = XMLDoc.createElement("TD")
                            td_node.text = "0"
                            tr_node.appendChild(td_node)
                        Next

                        td_node = XMLDoc.createElement("TD")
                        td_node.text = T_GR48_GOODWILL_AMT
                        tr_node.appendChild(td_node)

                        td_node = XMLDoc.createElement("TD")
                        td_node.text = T_GR48_SOFTWARE_AMT ' 合計 電腦軟體成本
                        tr_node.appendChild(td_node)

                        td_node = XMLDoc.createElement("TD")
                        td_node.text = T_GR48_LICENSE_AMT
                        tr_node.appendChild(td_node)

                        T_GR48_GOODWILL_AMT = 0
                        T_GR48_SOFTWARE_AMT = 0
                        T_GR48_LICENSE_AMT = 0
                    End If


                    OLD_ITEM_NM = TEMP_ITEM_NM
                    tr_node = XMLDoc.createElement("TR")
                    thead1_node.appendChild(tr_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = "W" & objDataReader(0)
                    tr_node.appendChild(td_node)

                    td_node = XMLDoc.createElement("TD")
                    td_node.text = objDataReader(1)
                    tr_node.appendChild(td_node)
                    iCompCnt = 0
                End If

                For i = iCompCnt To oCompLists.Count - 1
                    iCompCnt = iCompCnt + 1
                    If oCompLists(i)(1) = objDataReader(2) Then
                        '處理說明欄位與非說明欄位
                        If Not IsDBNull(objDataReader(10)) And objDataReader(10) = "N" Then  '為 非說明欄 head_yn = 'N' 
                            td_node = XMLDoc.createElement("TD")
                            td_node.text = objDataReader(3)
                            tr_node.appendChild(td_node)


                            td_node = XMLDoc.createElement("TD")
                            td_node.text = objDataReader(4)
                            tr_node.appendChild(td_node)

                            td_node = XMLDoc.createElement("TD")
                            td_node.text = objDataReader(5)
                            tr_node.appendChild(td_node)

                            If Not IsDBNull(objDataReader(11)) And objDataReader(11) = "-" Then
                                oCompLists(i)(2) = oCompLists(i)(2) - objDataReader(3) '商譽 Group
                                oCompLists(i)(3) = oCompLists(i)(3) - objDataReader(3) ' Group_sub

                                oCompLists(i)(4) = oCompLists(i)(4) - objDataReader(4) '軟體 Group
                                oCompLists(i)(5) = oCompLists(i)(5) - objDataReader(4) ' Group_sub

                                oCompLists(i)(6) = oCompLists(i)(6) - objDataReader(5) '專利 Group
                                oCompLists(i)(7) = oCompLists(i)(7) - objDataReader(5) ' Group_sub
                                T_GR48_GOODWILL_AMT = T_GR48_GOODWILL_AMT + objDataReader(3)
                                T_GR48_SOFTWARE_AMT = T_GR48_SOFTWARE_AMT + objDataReader(4)
                                T_GR48_LICENSE_AMT = T_GR48_LICENSE_AMT + objDataReader(5)
                            Else
                                oCompLists(i)(2) = oCompLists(i)(2) + objDataReader(3) '商譽 Group
                                oCompLists(i)(3) = oCompLists(i)(3) + objDataReader(3) ' Group_sub

                                oCompLists(i)(4) = oCompLists(i)(4) + objDataReader(4) '軟體 Group
                                oCompLists(i)(5) = oCompLists(i)(5) + objDataReader(4) ' Group_sub

                                oCompLists(i)(6) = oCompLists(i)(6) + objDataReader(5) '專利 Group
                                oCompLists(i)(7) = oCompLists(i)(7) + objDataReader(5) ' Group_sub
                                T_GR48_GOODWILL_AMT = T_GR48_GOODWILL_AMT + objDataReader(3)
                                T_GR48_SOFTWARE_AMT = T_GR48_SOFTWARE_AMT + objDataReader(4)
                                T_GR48_LICENSE_AMT = T_GR48_LICENSE_AMT + objDataReader(5)

                            End If

                            Exit For
                        Else  '說明欄處理

                            If OLD_GROUP_SUB <> TEMP_GROUP_SUB And TEMP_GROUP_SUB <> "" Then 'Group_sub

                                If Not IsDBNull(objDataReader(11)) And objDataReader(11) = "-" Then
                                    T_GR48_GOODWILL_AMT = T_GR48_GOODWILL_AMT - oCompLists(i)(3)
                                    T_GR48_SOFTWARE_AMT = T_GR48_SOFTWARE_AMT - oCompLists(i)(5)
                                    T_GR48_LICENSE_AMT = T_GR48_LICENSE_AMT - oCompLists(i)(7)
                                Else
                                    T_GR48_GOODWILL_AMT = T_GR48_GOODWILL_AMT + oCompLists(i)(3)
                                    T_GR48_SOFTWARE_AMT = T_GR48_SOFTWARE_AMT + oCompLists(i)(5)
                                    T_GR48_LICENSE_AMT = T_GR48_LICENSE_AMT + oCompLists(i)(7)
                                End If

                                td_node = XMLDoc.createElement("TD")    '商譽 Group_sub
                                td_node.text = oCompLists(i)(3)
                                oCompLists(i)(3) = 0
                                tr_node.appendChild(td_node)


                                td_node = XMLDoc.createElement("TD")    '軟體 Group_sub
                                ' 2013/4/23 Modified By Mikli
                                If Not IsDBNull(objDataReader(11)) And objDataReader(11) = "-" Then
                                    td_node.text = (-1) * oCompLists(i)(5)        ' 期末餘額  列
                                Else
                                    td_node.text = oCompLists(i)(5)               ' 期末餘額  列
                                End If
                                oCompLists(i)(5) = 0
                                tr_node.appendChild(td_node)

                                td_node = XMLDoc.createElement("TD")    '專利 Group_sub
                                td_node.text = oCompLists(i)(7)
                                oCompLists(i)(7) = 0
                                tr_node.appendChild(td_node)

                            End If
                            If OLD_GROUP <> TEMP_GROUP And TEMP_GROUP <> "" Then 'Group
                                If Not IsDBNull(objDataReader(11)) And objDataReader(11) = "-" Then
                                    T_GR48_GOODWILL_AMT = T_GR48_GOODWILL_AMT - oCompLists(i)(2)
                                    T_GR48_SOFTWARE_AMT = T_GR48_SOFTWARE_AMT - oCompLists(i)(4)
                                    T_GR48_LICENSE_AMT = T_GR48_LICENSE_AMT - oCompLists(i)(6)
                                Else
                                    T_GR48_GOODWILL_AMT = T_GR48_GOODWILL_AMT + oCompLists(i)(2)
                                    T_GR48_SOFTWARE_AMT = T_GR48_SOFTWARE_AMT + oCompLists(i)(4)
                                    T_GR48_LICENSE_AMT = T_GR48_LICENSE_AMT + oCompLists(i)(6)
                                End If


                                td_node = XMLDoc.createElement("TD")   ' 商譽 Group
                                td_node.text = oCompLists(i)(2)
                                oCompLists(i)(2) = 0
                                tr_node.appendChild(td_node)

                                td_node = XMLDoc.createElement("TD") ' 軟體 Group 列加總
                                td_node.text = oCompLists(i)(4)
                                oCompLists(i)(4) = 0
                                tr_node.appendChild(td_node)

                                td_node = XMLDoc.createElement("TD")  ' 專利 Group
                                td_node.text = oCompLists(i)(6)
                                oCompLists(i)(6) = 0
                                tr_node.appendChild(td_node)

                            End If
                            Exit For


                        End If
                    Else
                        td_node = XMLDoc.createElement("TD")
                        td_node.text = "0"
                        tr_node.appendChild(td_node)
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
                td_node = XMLDoc.createElement("TD")
                td_node.text = "0"
                tr_node.appendChild(td_node)
            Next
            td_node = XMLDoc.createElement("TD")
            td_node.text = T_GR48_GOODWILL_AMT
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = T_GR48_SOFTWARE_AMT '淨帳面金額
            tr_node.appendChild(td_node)

            td_node = XMLDoc.createElement("TD")
            td_node.text = T_GR48_LICENSE_AMT
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

            'Dim S_GR48_GOODWILL_AMT, S_GR48_LICENSE_AMT As Decimal
            'Dim T_GR48_GOODWILL_AMT, T_GR48_LICENSE_AMT As Decimal

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
            '        td_node.text = S_GR48_GOODWILL_AMT

            '        td_node = XMLDoc.createElement("TD")
            '        tr_node.appendChild(td_node)
            '        td_node.text = S_GR48_LICENSE_AMT

            '        '小計歸零
            '        S_GR48_GOODWILL_AMT = 0
            '        S_GR48_LICENSE_AMT = 0


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
            '        td_node.text = T_GR48_GOODWILL_AMT

            '        td_node = XMLDoc.createElement("TD")
            '        tr_node.appendChild(td_node)
            '        td_node.text = T_GR48_LICENSE_AMT

            '        '小計歸零
            '        T_GR48_GOODWILL_AMT = 0
            '        T_GR48_LICENSE_AMT = 0

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
            '                    S_GR48_GOODWILL_AMT = S_GR48_GOODWILL_AMT + 0
            '                    T_GR48_GOODWILL_AMT = T_GR48_GOODWILL_AMT + 0
            '                Else
            '                    td_node.text = objDataReader(i)
            '                    If OLD_GROUP_SUB <> "" Then
            '                        S_GR48_GOODWILL_AMT = S_GR48_GOODWILL_AMT + objDataReader(i)
            '                    End If
            '                    T_GR48_GOODWILL_AMT = T_GR48_GOODWILL_AMT + objDataReader(i)

            '                End If
            '            Case 3
            '                td_node = XMLDoc.createElement("TD")
            '                tr_node.appendChild(td_node)
            '                If IsDBNull(objDataReader(i)) Then
            '                    td_node.text = ""
            '                    S_GR48_LICENSE_AMT = S_GR48_LICENSE_AMT + 0
            '                    T_GR48_LICENSE_AMT = T_GR48_LICENSE_AMT + 0
            '                Else
            '                    td_node.text = objDataReader(i)

            '                    If OLD_GROUP_SUB <> "" Then
            '                        S_GR48_LICENSE_AMT = S_GR48_LICENSE_AMT + objDataReader(i)
            '                    End If
            '                    T_GR48_LICENSE_AMT = T_GR48_LICENSE_AMT + objDataReader(i)

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
            '    td_node.text = S_GR48_GOODWILL_AMT

            '    td_node = XMLDoc.createElement("TD")
            '    tr_node.appendChild(td_node)
            '    td_node.text = S_GR48_LICENSE_AMT
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
            'td_node.text = T_GR48_GOODWILL_AMT

            'td_node = XMLDoc.createElement("TD")
            'tr_node.appendChild(td_node)
            'td_node.text = T_GR48_LICENSE_AMT
        End If

        xmlData = Trim(XMLDoc.childNodes(0).xml)
    End Sub
    Protected Sub oSimpleGrid_OnSimpleGridShowData(ByVal XMLDoc As System.Xml.XmlDocument, ByVal DataNode As System.Xml.XmlElement, ByVal Column As System.Data.DataColumn, ByVal GridRow As System.Data.DataRow, ByRef Changed As Boolean) Handles oSimpleGrid.OnSimpleGridShowData

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' 檢查是否有傳 XMLDATA
        If CheckXMLDATA() Then Call Main()
    End Sub

End Class
