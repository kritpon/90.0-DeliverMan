

Module DBtools

    '=========   ���������ͧ��� �ѹ�֡������
    Public chkSaveOK As Boolean = False

    Public strFactor As String = ""
    Public strCon As String
    'Public SLid As String

    Public pathDB As String
    Public pathDB02 As String


    Public Conn As SqlClient.SqlConnection = New SqlClient.SqlConnection()
    Public subCom As SqlClient.SqlCommand = New SqlClient.SqlCommand
    Public txtSQL As String
    Public txtSQL1 As String
    Public txtSQL2 As String

    Public docType As String = "S"

    Public Str01 As String
    Public gDA As SqlClient.SqlDataAdapter
    Public gDs As DataSet = New DataSet()
    Public selWH As String = ""  '���͡��ѧ�Թ����

    Public selCusName As String
    'Public selCountry As String
    'Public selAmphur As String
    'Public selZip As String
    Public selCusiD As String

    Public selSale As String
    Public selSaleName As String


    Public strConn As String

    Public CusId As String
    Public CusName As String
    Public CusCreTerm As Integer
    Public CusSaleName As String
    Public CusSaleID As String
    Public CusDiscount As String 'Double
    Public CusLimit As Double

    Public chkNew As Boolean = False
    Public chkEditDoc As Boolean = False
    Public EditNo As String
    Public EditCus As String

    Public PId As String = "" '�������١���
    Public PName As String = "" '�纪����١���
    Public Pdate As String = ""
    Public Pdate2 As String = ""
    Public SelectDate As Date
    Public SelectNo As String = "" '���Ţ��� Order

    Public SelectCode As String = "" '�������Թ���
    Public SelectName As String = "" '�纪����Թ���
    Public SelectNum As Integer = 0 '�纨ӹǹ
    Public SelectPrice As Double = 0 '
    Public SelectDisc As String = ""

    Public SelectWeight As Double = 0 '�纹��˹ѡ��
    Public SelectStock As Double = 0 '��Stock
    Public SelectPNo As String
    Public Stock As Integer = 0  'stock

    Public SelectNo2 As String '���Ţ��� Order
    Public SelectName2 As String = "" '�纪����Թ���
    Public SelectNum2 As Integer = 0 '�纨ӹǹ
    Public SelectPrice2 As Double = 0 '���Ҥ�
    Public SelectCode2 As String = "" '�������Թ���
    Public SelectWeight2 As Double = 0  '�纹��˹ѡ��
    Public SelectStock2 As Double = 0 '��Stock
    Public SelectPNo2 As String = ""
    Public SelectRevNo2 As String = ""  ' �红�����㺨ͧ
    Public SelectTypeP_No2 As String = ""  ' ����ԡ
    Public selTypeVAT As String = ""

    Public selectDetail As String = ""


    Public CodeT As String = ""
    Public CodeG As String = ""
    Public CodeColor As String = ""
    Public CodeTh As String = ""
    Public CodeSize As String = ""
    Public CodePaper As String = ""
    Public CodeGrade As String = ""

    Public Ddate As String = ""
    Public Dno As String = ""
    Public orderNum As String = ""
    Public Dvat As String = ""
    Public DPNo As String = ""
    Public Dcus As String = ""
    Public Dpro As String = ""
    Public Dnum As Integer = 0
    Public Dprice As Single = 0
    Public Dweight As Single = 0
    Public Dproduct As String = ""
    Public Dcusname As String = ""
    Public Ditem As String = ""
    Public DOrder As String = ""
    Public DDetail2 As String = ""
    Public DDisc As String = ""
    Public NoDoc As String = ""

    Public LEdit As ListViewItem
    Public selStkID As String   ' �������������Ѻ�纤�� �����Թ���

    Public trh_RunID As Integer = 0

    Public subDs As DataSet = New DataSet
    Public subDa As SqlClient.SqlDataAdapter

    'Public CustomerId As String
    Public ThaiBaht01 As String
    Public Num As Integer
    Public chkID As Boolean
    Public NumberDoc As String
    'Public TypeDoc As String

    Public ChkDClick As Boolean = True
    Public chkItem As Boolean = False
    Public chkLoad As Boolean = False


    Declare Function GetUserName Lib "advapi32.dll" Alias _
                  "GetUserNameA" (ByVal lpBuffer As String, _
                  ByRef nSize As Integer) As Integer

    Public Function GetUserName() As String '��Username Password�ͧ����ͧ�����
        Dim iReturn As Integer
        Dim userName As String
        userName = New String(CChar(" "), 50)
        iReturn = GetUserName(userName, 50)
        GetUserName = userName.Substring(0, userName.IndexOf(Chr(0)))
    End Function

    Sub DBConnection()

        strConn = DBConnString.strConn2
        With Conn
            If .State = ConnectionState.Open Then .Close()
            .ConnectionString = strConn
            .Open()
        End With

    End Sub

    Sub openDB()
        strConn = DBConnString.strConn2
        With Conn
            If .State = ConnectionState.Open Then .Close()
            .ConnectionString = strConn
            .Open()
        End With

    End Sub

    Sub closeDB()
        Conn.Close()
    End Sub

    'Sub DBConnection()
    '    'pathDB = "i:\center\PC50.mdb"
    '    'pathDB = "i:\center\PC50.mdb"
    '    'pathDB = "i:\test49\db\test49.mdb"
    '    'pathDB02 = "i:\center\acct50.mdb"
    '    'strConn = "server=(local);database=TestDB2006;Trusted_Connection=yes"
    '    'DB01 = DAODBEngine_definst.OpenDatabase(pathDB)
    '    ''DB02 = DAODBEngine_definst.OpenDatabase(pathDB02)
    '    strConn = "server=ENJOY-PC\sqlexpress;database=TestDB2006;Trusted_Connection=yes;" 'connect����ͧ����ͧ
    '    'strConn = "server=EDP;database=DB2006;Trusted_Connection=yes;"
    '    'strConn = "Data Source=EDP2\SQLEXPRESS;Initial Catalog=db2006;User ID=sa;Password=sys0500" 'connect server
    '    'strConn = "Data Source=Time\SQLEXPRESS;Initial Catalog=backup_db2006;User ID=sa;Password=sys0500" 'connect server
    '    'strConn = "server=TIME\sqlexpress;database=backup_db2006;Trusted_Connection=yes;" 'connect����ͧ����ͧ
    '    With Conn
    '        If .State = ConnectionState.Open Then .Close()
    '        .ConnectionString = strConn
    '        .Open()
    '    End With
    'End Sub

    Sub dbDelDATA(ByVal txtSQL As String, ByVal txtDisy As String)
        Try
            'If MessageBox.Show("��ͧ���ź������ ' " & txtDisy & " ' ����к��������", "���׹�ѹ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            'DB01.Execute(txtSQL) ' �ѹ�֡������ŧ Business sc50
            'DB02.Execute(txtSQL) ' �ѹ�֡������ŧ Business acct50
            If Conn.State = ConnectionState.Closed Then
                DBtools.DBConnection()
            End If

            With subCom
                .CommandType = CommandType.Text
                .CommandText = txtSQL
                .Connection = Conn
                .ExecuteNonQuery()
            End With
            'End If
        Catch errprocess As Exception
            MessageBox.Show("�������öź�����������ͧ�ҡ " & errprocess.Message, "��ͼԴ��Ҵ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Sub dbDelSQLsrv(ByVal txtSQL As String, ByVal txtDisy As String)

        Try
            If MessageBox.Show("��ͧ���ź������ ' " & txtDisy & " ' ����к��������", "���׹�ѹ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                'DB01.Execute(txtSQL) ' �ѹ�֡������ŧ Business sc50
                'DB02.Execute(txtSQL) ' �ѹ�֡������ŧ Business acct50
                If Conn.State = ConnectionState.Closed Then
                    DBtools.DBConnection()
                End If
                With subCom
                    .CommandType = CommandType.Text
                    .CommandText = txtSQL
                    .Connection = Conn
                    .ExecuteNonQuery()
                End With
            End If
        Catch errprocess As Exception
            MessageBox.Show("�������öź�����������ͧ�ҡ " & errprocess.Message, "��ͼԴ��Ҵ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Sub dbSaveDATA(ByVal txtSQL As String, ByVal txtDisy As String)

        Try
            '        If MessageBox.Show("��ͧ��úѹ�֡������ ' " & txtDisy & " ' ����к��������", "���׹�ѹ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            With subCom
                .CommandType = CommandType.Text
                .CommandText = txtSQL
                .Connection = Conn
                .ExecuteNonQuery()
            End With
            'DB01.Execute(txtSQL)  ' �ѹ�֡������ŧ Business 
            'DB02.Execute(txtSQL)
            'MsgBox("�����Ŷ١�ѹ�֡���º��������", MsgBoxStyle.OkOnly, "�駼š�÷ӧҹ")
            'Else
            'MsgBox("�������ѧ�����١�ѹ�֡", MsgBoxStyle.OkOnly, "�駼š�÷ӧҹ")
            'End If

        Catch errprocess As Exception
            MessageBox.Show("�������ö���������������ͧ�ҡ " & errprocess.Message, "��ͼԴ��Ҵ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Sub dbSaveSQLsrv(ByVal txtSQL As String, ByVal txtDisy As String)

        Try
            'If MessageBox.Show("��ͧ��úѹ�֡������ ' " & txtDisy & " ' ����к��������", "���׹�ѹ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            If Conn.State = ConnectionState.Closed Then
                DBtools.DBConnection()
            End If

            With subCom
                .CommandType = CommandType.Text
                .CommandText = txtSQL
                .Connection = Conn
                .ExecuteNonQuery()
            End With
            'DB01.Execute(txtSQL) ' �ѹ�֡������ŧ Business 
            'DB02.Execute(txtSQL) ' �ѹ�֡������ŧ Business acct50
            'End If
        Catch errprocess As Exception
            MessageBox.Show("�������ö���������������ͧ�ҡ " & errprocess.Message, "��ͼԴ��Ҵ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    'Sub readDB()
    '    Dim ansTB As String
    '    Dim ansF As String
    '    Dim sizeF As String
    '    Dim countTB As Integer
    '    Dim countF As Integer
    '    With DB01
    '        For countTB = 0 To DB01.TableDefs.Count - 1
    '            ansTB = .TableDefs(countTB).Name
    '            For countF = 0 To .TableDefs(countTB).Fields.Count - 1
    '                ansF = .TableDefs(countTB).Fields(countF).Name
    '                sizeF = Convert.ToString(.TableDefs(countTB).Fields(countF).Size)

    '            Next
    '        Next
    '    End With
    '    With Conn

    '    End With
    'End Sub
    Sub dbSaveUser(ByVal txtSQL As String, ByVal txtDisy As String)


        Try

            If Conn.State = ConnectionState.Closed Then
                DBtools.DBConnection()
            End If
            With subCom
                .CommandType = CommandType.Text
                .CommandText = txtSQL
                .Connection = Conn
                .ExecuteNonQuery()
            End With



        Catch errprocess As Exception
            MessageBox.Show("�������ö���������������ͧ�ҡ " & errprocess.Message, "��ͼԴ��Ҵ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Function getSaleName(ByVal saleId As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From salesman "

        txtSQL = txtSQL & "WHERE SL_id='" & saleId & "'"


        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "slList")

        If subDS.Tables("slList").Rows.Count - 1 < 0 Then
            ans = ""
        Else
            ans = subDS.Tables("slList").Rows(0).Item("sl_Name")
        End If

        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Public Function ThaiBaht(ByVal pamt As Double) As String
        Dim i As Integer, j As Integer
        'Dim v As Integer
        Dim Valstr As String, Vlen As String, Vno As String
        'Dim syslge As String
        Dim wnumber(10) As String, wdigit(10) As String, spcdg(5) As String
        Dim vword(20) As String
        If pamt <= 0 Then
            ThaiBaht = ""
            Exit Function
        End If
        Valstr = Trim(Format$(pamt, "##########0.00"))
        Vlen = Len(Valstr) - 3
        For i = 1 To 20
            vword(i) = ""
        Next i
        wnumber(1) = "˹��" : wnumber(2) = "�ͧ" : wnumber(3) = "���" : wnumber(4) = "���" : wnumber(5) = "���"
        wnumber(6) = "ˡ" : wnumber(7) = "��" : wnumber(8) = "Ỵ" : wnumber(9) = "���" : wdigit(1) = "�ҷ"
        wdigit(2) = "�Ժ" : wdigit(3) = "����" : wdigit(4) = "�ѹ" : wdigit(5) = "����" : wdigit(6) = "�ʹ" : wdigit(7) = "��ҹ"
        spcdg(1) = "ʵҧ��" : spcdg(2) = "���" : spcdg(3) = "���" : spcdg(4) = "��ǹ"
        For i = 1 To Vlen
            Vno = Int(Val(Mid$(Valstr, i, 1)))
            If Vno = 0 Then
                vword(i) = ""

                If (Vlen - i + 1) = 7 Then
                    vword(i) = wdigit(7) '��ҹ
                End If
            Else
                If (Vlen - i + 1) > 7 Then
                    j = Vlen - i - 5 ' �Թ��ѡ��ҹ
                Else
                    j = Vlen - i + 1 '��ѡ�ʹ
                End If
                vword(i) = wnumber(Vno) + wdigit(j)  '30-90
            End If

            If Vno = 1 And j = 2 Then
                vword(i) = wdigit(2) '�Ժ
            End If
            If Vno = 2 And j = 2 Then
                vword(i) = spcdg(3) + wdigit(j) '����Ժ
            End If
            If j = 1 Then
                vword(i) = wnumber(Vno)
                If Vno = 1 And Vlen > 1 Then
                    If Mid$(Valstr, i - 1, 1) <> "0" Then
                        vword(i) = spcdg(2)
                    End If
                End If
            End If
            If j = 7 Then
                vword(i) = wnumber(Vno) + wdigit(j) '��ҹ
                If Vno = 1 And Vlen > 7 Then
                    If Mid$(Valstr, i - 1, 1) <> "0" Then
                        vword(i) = spcdg(2) + wdigit(j)
                    End If
                End If
            End If
        Next i
        If Int(pamt) > 0 Then
            vword(Vlen) = vword(Vlen) + wdigit(1)
        End If


        Valstr = Mid$(Valstr, Vlen + 2, 2) '�ȹ���
        Vlen = Len(Valstr)
        For i = 1 To Vlen
            Vno = Int(Val(Mid$(Valstr, i, 1)))
            If Vno = 0 Then
                vword(i + 10) = ""
            Else
                j = Vlen - i + 1
                vword(i + 10) = wnumber(Vno) + wdigit(j)
                If Vno = 1 And j = 2 Then
                    vword(i + 10) = wdigit(2)
                End If
                If Vno = 2 And j = 2 Then
                    vword(i + 10) = spcdg(3) + wdigit(j)
                End If
                If j = 1 Then
                    If Vno = 1 And Int(Val(Mid$(Valstr, i - 1, 1))) <> 0 Then
                        vword(i + 10) = spcdg(2)
                    Else
                        vword(i + 10) = wnumber(Vno)
                    End If
                End If
            End If

        Next i
        If pamt <> 0 Then
            If Val(Valstr) = 0 Then
                vword(13) = spcdg(4)
            Else
                vword(13) = spcdg(1)
            End If
        End If
        Valstr = ""
        For i = 1 To 20
            Valstr = Valstr + vword(i)
        Next i
        ThaiBaht = (Valstr)
    End Function

    '=====================   Function  ����� ���ͺ�����ҵ�ҧ� � DataBase ============================
    Function getArVat(ByVal ar_Code As String) As String
        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_VAT")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getCusLimit(ByVal ar_Code As String) As Integer

        Dim ans As Integer = 0
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Term")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getArAddr1(ByVal ar_Code As String) As String
        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Addr")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Function getArAddr2(ByVal ar_Code As String) As String
        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Addr_1")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getDocPO(orderNo As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "

        txtSQL = txtSQL & "WHERE (dtl_Type='B')"
        txtSQL = txtSQL & "And  (((TranDataD.Dtl_con_id) = '" & orderNo & "')) "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")
        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Dtl_PO")) Then
                ans = ""
            Else

                ans = subDS.Tables("OrderList").Rows(0).Item("Dtl_PO")
            End If
        Else
            ans = ""
        End If
       

        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getDocPO(orderNo As String, stkCode As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "

        txtSQL = txtSQL & "WHERE (dtl_Type='B')"
        txtSQL = txtSQL & "And TranDataD.Dtl_idTrade='" & stkCode & "' "
        txtSQL = txtSQL & "And  (((TranDataD.Dtl_con_id) = '" & orderNo & "')) "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")
        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Dtl_PO")) Then
                ans = ""
            Else
                ans = subDS.Tables("OrderList").Rows(0).Item("Dtl_PO")
            End If
        Else
            ans = ""
        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Function getOrComment(orderNo As String, stkCode As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "

        txtSQL = txtSQL & "WHERE (dtl_Type='B')"
        txtSQL = txtSQL & "And TranDataD.Dtl_idTrade='" & stkCode & "' "
        txtSQL = txtSQL & "And  (((TranDataD.Dtl_con_id) = '" & orderNo & "')) "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")
        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Dtl_detail")) Then
                ans = ""
            Else
                ans = subDS.Tables("OrderList").Rows(0).Item("Dtl_Detail")
            End If
        Else
            ans = ""
        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getDiscTrH(orderNo As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataH "

        txtSQL = txtSQL & "WHERE (Trh_Type='B')"
        txtSQL = txtSQL & "And TranDataH.Trh_No='" & orderNo & "' "


        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")

        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Trh_Disc")) Then
                ans = 0
            Else
                ans = subDS.Tables("OrderList").Rows(0).Item("Trh_Disc")
            End If
        Else
            ans = 0
        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getDiscProd(orderNo As String, stkCode As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "

        txtSQL = txtSQL & "WHERE (dtl_Type='B')"
        txtSQL = txtSQL & "And TranDataD.Dtl_idTrade='" & stkCode & "' "
        txtSQL = txtSQL & "And  (((TranDataD.Dtl_con_id) = '" & orderNo & "')) "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")

        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Dtl_T_Disc")) Then
                ans = 0
            Else
                ans = subDS.Tables("OrderList").Rows(0).Item("Dtl_T_Disc")
            End If
        Else
            ans = 0
        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getArAddr3(ByVal ar_Code As String) As String
        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Addr_2")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Function getArAddrFull(ByVal ar_Code As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & ar_Code & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Addr") & " " & subDS.Tables("ARList").Rows(0).Item("Ar_Addr_1") & " " & subDS.Tables("ARList").Rows(0).Item("Ar_Addr_2")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getTypeVATbyT(RevNo As String) As String

        Dim txtSQL As String = ""
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataH "
        txtSQL = txtSQL & "Where trh_Type='T' "
        txtSQL = txtSQL & "And trh_NO='" & RevNo & "' "



        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "dbOrder")


        If subDS.Tables("dbOrder").Rows.Count = 0 Then
            Return "N"
        Else
            If IsDBNull(subDS.Tables("dbOrder").Rows(0).Item("trh_NoType")) Then
                Return "V"
            Else

                Return subDS.Tables("dbOrder").Rows(0).Item("trh_NoType")
            End If


        End If

    End Function
    Function getCusName(ByVal cusId As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Ar_Cus_ID,Ar_Name,Ar_C_Term,Ar_Target,Ar_Cre_Lim "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & cusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Name")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getSalesCode(ByVal cusId As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Ar_Cus_ID,Ar_Name,Ar_Sales,Ar_C_Term,Ar_Target,Ar_Cre_Lim "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & cusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Sales")
        subDS = Nothing
        subDA = Nothing

        Return ans

    End Function

    Function getArDisc(ByVal cusId As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Ar_Cus_ID,Ar_Name,Ar_Sales,AR_DISC,Ar_Target,Ar_Cre_Lim "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & cusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("AR_DISC")
        subDS = Nothing
        subDA = Nothing

        Return ans

    End Function

    Function getStkWight(ByVal stkCode As String) As Double
        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If stkCode = "" Then
            ans = 0

        Else

            txtSQL = "Select * "
            txtSQL = txtSQL & "From BaseMast "

            txtSQL = txtSQL & "WHERE (((Stk_Code)='" & stkCode & "'))"

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "command")

            ans = subDS.Tables("command").Rows(0).Item("Stk_Factor")

        End If

      
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Function getDocNumber(ByVal DocNo As String, ByVal DocType As String) As Boolean
        Dim ans As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If Trim(DocNo) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return False
        Else
            DBtools.openDB()
            txtSQL = "Select * "
            txtSQL = txtSQL & "From TranDataH "

            txtSQL = txtSQL & "WHERE ((Trh_Type='" & DocType & "') "
            txtSQL = txtSQL & "And (Trh_No='" & DocNo & "')) "
            'txtSQL=txtSQL & "And Trh_Wh='" & "'"  ' �������� ��ѧ�Թ���

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                ans = True
            Else
                ans = False
            End If

            subDS = Nothing
            subDA = Nothing
            DBtools.closeDB()
            Return ans
        End If

    End Function

    'Public Function GetExcelVersion() As String
    '    Dim excel As Object = Nothing
    '    Dim ver As Integer = 0
    '    Dim build As Integer
    '    Try
    '        excel = CreateObject("Excel.Application")
    '        ver = excel.Version
    '        build = excel.Build
    '    Catch ex As Exception
    '        'Continue to finally sttmt
    '    Finally
    '        Try
    '            Marshal.ReleaseComObject(excel)
    '        Catch
    '        End Try
    '        GC.Collect()
    '    End Try
    '    Return ver
    'End Function
    Function getChkTypeP_State(ByVal strNo As String, idTrade As String) As Boolean

        Dim ans As Boolean = False

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "
        txtSQL = txtSQL & "WHERE Dtl_Type='P' "
        txtSQL = txtSQL & "And Dtl_No='" & strNo & "' "

        txtSQL = txtSQL & "And Dtl_idTrade='" & idTrade & "' "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "Ans")

        If subDS.Tables("Ans").Rows.Count > 0 Then

            If IsDBNull(subDS.Tables("Ans").Rows(0).Item("Dtl_State")) Then
                ans = False
            Else
                ans = subDS.Tables("Ans").Rows(0).Item("Dtl_State")
            End If

        Else
            ans = 0
        End If

        subDS = Nothing
        subDA = Nothing
        Return ans


    End Function
    Function getDocNumberD(ByVal DocNo As String, ByVal DocType As String, ByVal stkCode As String) As Boolean
        Dim ans As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If Trim(DocNo) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return False
        Else
            DBtools.openDB()
            txtSQL = "Select * "
            txtSQL = txtSQL & "From TranDataD "

            txtSQL = txtSQL & "WHERE ((dtl_Type='" & DocType & "') "
            txtSQL = txtSQL & "And (dtl_No='" & DocNo & "')"
            txtSQL = txtSQL & "And (dtl_idTrade='" & stkCode & "') "
            txtSQL = txtSQL & ") "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                ans = True
            Else
                ans = False
            End If

            subDS = Nothing
            subDA = Nothing
            DBtools.closeDB()
            Return ans
        End If

    End Function

    Function getDocType(ByVal docType As String) As String

        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TypeDocMast "

        txtSQL = txtSQL & "WHERE (((Type_ID) Like '%" & docType & "%'))"

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "command")

        ans = subDS.Tables("command").Rows(0).Item("Type_Name")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function


    Function getWhName(ByVal WhCode As String) As String

        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If Trim(WhCode) = "" Then
            Return ""
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From WareHouse "

            txtSQL = txtSQL & "WHERE Wh_id='" & WhCode & "' "
            txtSQL = txtSQL & "And Wh_Type='DC' "


            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "Ans")

            If subDS.Tables("Ans").Rows.Count > 0 Then
                ans = subDS.Tables("Ans").Rows(0).Item("Wh_Name")
            Else
                MsgBox("��辺������ DC ����ͧ���")
                ans = ""
            End If

            subDS = Nothing
            subDA = Nothing
            Return ans

        End If


    End Function

    Function getStkNoDoc(ByVal DocType As String, ByVal DocNo As String, ByVal stkCode As String) As Boolean
        Dim ans As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If Trim(DocNo) = "" Then
            Return False
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From TranDataD "

            txtSQL = txtSQL & "WHERE Dtl_Type='" & DocType & "' "
            txtSQL = txtSQL & "And Dtl_No='" & DocNo & "' "
            txtSQL = txtSQL & "And Dtl_Stk='" & stkCode & "' "


            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                ans = True
            Else
                ans = False
            End If

            subDS = Nothing
            subDA = Nothing
            Return ans
        End If

    End Function

    Function getCostByStk(ByVal stkCode As String, ByVal CSDate As Date, ByVal chkRun As Boolean) As Double
        '  CSDate  ���   �ѹ���  ����ͧ��õ鹷ع
        '  chkRun  ���   
        Dim txtSQL As String = ""
        Dim subDa As SqlClient.SqlDataAdapter
        Dim subDs As DataSet = New DataSet
        Dim Ans As Double


        txtSQL = "Select * "
        txtSQL = txtSQL & "From CostMast "
        txtSQL = txtSQL & "Where CS_Stk_Code='" & stkCode & "' "

        If CSDate = "01/01/2013" Then
        Else
            txtSQL = txtSQL & "And CS_Date='" & Microsoft.VisualBasic.Right(Year(DateAdd(DateInterval.Year, -543, CSDate)), 2) & "/" & Format(Month(CSDate), "00") & "' "
        End If

        txtSQL = txtSQL & "Order by CS_Date desc "


        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subDs, "master")

        If chkRun = False Then
            If subDs.Tables("master").Rows.Count > 0 Then
                Ans = subDs.Tables("master").Rows(0).Item("CS_Cost")
                Return Ans
            Else
                Ans = getCostByStk(stkCode, "01/01/2013", 1)
                Return Ans '100 'getCostByStk(stkCode, "")
            End If
            '===============================================
        Else
            If subDs.Tables("master").Rows.Count > 0 Then
                Ans = subDs.Tables("master").Rows(0).Item("CS_Cost")
                Return Ans
            Else
                Ans = 0
                Return Ans '100 'getCostByStk(stkCode, "")
            End If

        End If


    End Function

    Function getCusTaxCode(ByVal cusCode As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Ar_Cus_ID,Ar_Name,Ar_C_Term,Ar_Target,Ar_Cre_Lim,Ar_Tax_Code "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & cusCode & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")
        If IsDBNull(subDS.Tables("ARList").Rows(0).Item("Ar_Tax_Code")) Then
            ans = ""
        Else
            ans = subDS.Tables("ARList").Rows(0).Item("Ar_Tax_Code")
        End If

        subDS = Nothing
        subDA = Nothing
        Return ans


    End Function

    Function getCostType(ByVal stkCode As String) As Integer

        '  ����觤�ҡ�Ѻ�� 0 �Դ�鹷ع��� ���˹ѡ 
        '  ����觤�ҡ�Ѻ��  1  �Դ�鹷ع��ͪ��



        Dim txtSQL As String = ""
        Dim subDa As SqlClient.SqlDataAdapter
        Dim subDs As DataSet = New DataSet
        Dim Ans As Integer

        txtSQL = "Select * "
        txtSQL = txtSQL & "From CostMast "
        txtSQL = txtSQL & "Where CS_Stk_Code='" & stkCode & "'"

        txtSQL = txtSQL & "Order by CS_Date desc "


        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subDs, "master")

        If subDs.Tables("master").Rows.Count > 0 Then
            Ans = subDs.Tables("master").Rows(0).Item("CS_Type")
            Return Ans
        Else
            Return 9 '  ����դ��

        End If

    End Function

    Function getChkStkDetl(ByVal StrCode As String, ByVal StkCode As String) As Boolean
        Dim ans As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(StkCode) = "" Or Trim(StrCode) = "" Then
            Return False
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From StkDetl "

            txtSQL = txtSQL & "WHERE ((Dtl_Store='" & StrCode & "') "
            txtSQL = txtSQL & "And (Dtl_Code='" & StkCode & "')) "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                ans = True
            Else
                ans = False
            End If

            subDS = Nothing
            subDA = Nothing
            Return ans
        End If

    End Function

    Function getProDCode(ByVal StkCode As String) As String
        'Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(StkCode) = "" Then
            Return ""
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From BaseMast "

            txtSQL = txtSQL & "WHERE  "
            txtSQL = txtSQL & "(Stk_Code='" & StkCode & "') "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                Return subDS.Tables("StkList").Rows(0).Item("Stk_ProD").ToString
            Else
                Return ""
            End If

        End If

    End Function

    Function getChkBaseMast(ByVal StkCode As String) As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(StkCode) = "" Or Trim(StkCode) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return False
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From BaseMast "
            txtSQL = txtSQL & "WHERE (Stk_Code='" & StkCode & "')"

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End If

    End Function

    Function getStkName(ByVal stkId As String) As String

        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        Try
            If Trim(stkId) = "" Then
                subDS = Nothing
                subDA = Nothing
                Return ""
            Else
                txtSQL = "Select * "
                txtSQL = txtSQL & "From BaseMast "

                txtSQL = txtSQL & "WHERE Stk_Code='" & stkId & "'"

                subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
                subDA.Fill(subDS, "stkList")

                ans = subDS.Tables("stkList").Rows(0).Item("stk_Name_1")
                subDS = Nothing
                subDA = Nothing
                Return ans

            End If
        Catch ex As Exception
            Return ""
        End Try



    End Function

    'Function getStock(ByVal stkId As String, ByVal strID As String, ByVal stkWH As String) As Double '����Ѻ�� Stock

    '    Dim ans As String

    '    Dim subDA As SqlClient.SqlDataAdapter
    '    Dim subDS As New DataSet

    '    If Trim(stkId) = "" Then
    '        Return 0 '""
    '        Exit Function
    '    Else
    '        txtSQL = "Select * "
    '        txtSQL = txtSQL & "From StkDetl "

    '        txtSQL = txtSQL & "WHERE (((Dtl_Code) Like '%" & stkId & "%')) "
    '        txtSQL = txtSQL & "And (((Dtl_Store) Like '%" & strID & "%')) "
    '        txtSQL = txtSQL & "And (Dtl_WH='" & stkWH & "')"

    '        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
    '        subDA.Fill(subDS, "stkList")

    '        If subDS.Tables("stkList").Rows.Count > 0 Then
    '            ans = subDS.Tables("stkList").Rows(0).Item("Dtl_Bal_Q1")
    '        Else
    '            ans = 0
    '        End If

    '        subDS = Nothing
    '        subDA = Nothing
    '        Return ans

    '    End If


    'End Function

    Function getStock(ByVal stkId As String, ByVal strID As String, ByVal whCode As String) As Integer
        Dim ans As Integer = 0
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        Try
            If Trim(stkId) = "" Then
                subDS = Nothing
                subDA = Nothing
                Return ""
            Else
                txtSQL = "Select Dtl_Wh,Dtl_Code,Dtl_Bal_Q1 "
                txtSQL = txtSQL & "From StkDetl "

                txtSQL = txtSQL & "WHERE Dtl_Code='" & stkId & "' "
                txtSQL = txtSQL & "And Dtl_Store='110098' "
                txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "' "

                'txtSQL = txtSQL & "group by Dtl_Wh,Dtl_Code "

                subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
                subDA.Fill(subDS, "StkStock")

                If subDS.Tables("StkStock").Rows.Count > 0 Then
                    ans = subDS.Tables("StkStock").Rows(0).Item("Dtl_Bal_Q1")
                Else
                    ans = 0
                End If

                subDS = Nothing
                subDA = Nothing

                Return ans
            End If
        Catch ex As Exception

            Return 0
        End Try

    End Function

    '=====================   Function  ����� ���ͺ�����ҵ�ҧ� � DataBase ============================
    Public Sub rightPrint(ByVal strNum As Double, ByVal PositionX As Integer, ByVal PositionY As Integer)
        Dim txtAns1, txtAns2, txt As String
        Dim i As Integer
        txtAns1 = Str(Format(strNum, "#,###,###.00"))
        txtAns1 = Format(txtAns1, "#,###,###.00")

        For i = 0 To Len(txtAns1) - 1
            txt = (Right(txtAns1, 1))
            If txt = "," Then
                PositionX = PositionX + 50
            End If
            'Printer.CurrentX = PositionX
            'Printer.CurrentY = PositionY
            'Printer.Print(txt)
            txtAns2 = Left(txtAns1, Len(txtAns1) - 1) '������ �Ѵ��Ƿ�����������
            txtAns1 = txtAns2
            PositionX = PositionX - 98
        Next i
    End Sub

    Function getPending(ByVal cusCode As String, ByVal stkCode As String) As Double

        Dim ans As Double

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(stkCode) = "" Or Trim(cusCode) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return False
        Else
            txtSQL = "Select Dtl_idCus,Dtl_idTrade,sum(Dtl_Num-Dtl_Num_2)as penDing "
            txtSQL = txtSQL & "From TranDataD "

            txtSQL = txtSQL & "Where Dtl_idCus='" & cusCode & "'"
            txtSQL = txtSQL & "And Dtl_idTrade='" & stkCode & "'"
            txtSQL = txtSQL & "And Dtl_Type='B' "
            txtSQL = txtSQL & "Group by Dtl_idCus,Dtl_idTrade "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "PendingTB")
            If subDS.Tables("PendingTB").Rows.Count > 0 Then
                ans = subDS.Tables("PendingTB").Rows(0).Item("penDing")
                Return ans
            Else
                subDS = Nothing
                subDA = Nothing
                Return 0
            End If


        End If
    End Function

    ' �� Update  Stkdetl  Ẻ�����
    Sub updateStock(ByVal CusID As String, ByVal whCode As String, ByVal StkCode As String, ByVal OrderQ1 As Double, ByVal RcvQ1 As Double, ByVal IssQ1 As Double, ByVal PenDingUpdate As Boolean)

        Dim subDA3 As New SqlClient.SqlDataAdapter
        Dim subDS3 As New DataSet

        Dim wStk As Double = DBtools.getStkWight(StkCode)

        Dim Dtl_Bal_q2 As Double = 0
        Dim Dtl_Bal_q1 As Double = 0
        Dim Dtl_LS_Q1 As Double = 0
        Dim Dtl_LS_Q2 As Double = 0

        txtSQL = "Select * From StkDetl "
        txtSQL = txtSQL & "Where Dtl_Code='" & StkCode & "' "
        txtSQL = txtSQL & "And  Dtl_Store='" & CusID & "' "
        txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "'"

        subDA3 = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA3.Fill(subDS3, "chkStk")


        If subDS3.Tables("chkStk").Rows.Count > 0 Then
            With subDS3.Tables("chkStk").Rows(0)

                If IsDBNull(.Item("Dtl_Bal_q1")) Then
                    Dtl_Bal_q1 = 0
                Else
                    Dtl_Bal_q1 = .Item("Dtl_Bal_q1")
                End If

                If IsDBNull(.Item("Dtl_Bal_q2")) Then
                    Dtl_Bal_q2 = 0
                Else
                    Dtl_Bal_q2 = .Item("Dtl_Bal_q2")
                End If

                If IsDBNull(.Item("Dtl_LS_Q1")) Then
                    Dtl_LS_Q1 = 0
                Else
                    Dtl_LS_Q1 = .Item("Dtl_LS_Q1")
                End If

                If IsDBNull(.Item("Dtl_LS_Q2")) Then
                    Dtl_LS_Q2 = 0
                Else
                    Dtl_LS_Q2 = .Item("Dtl_LS_Q2")
                End If

                txtSQL = "Update StkDetl Set "
                If IsDBNull(.Item("Dtl_Or_Q1")) Then
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & ((0 + OrderQ1)) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & ((.Item("Dtl_Or_Q1") + OrderQ1)) & ","
                End If

                If IsDBNull(.Item("Dtl_Or_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Or_Q2=" & ((0 + (OrderQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Or_Q2=" & ((.Item("Dtl_Or_Q1") + OrderQ1) * wStk) & ","
                End If

                If IsDBNull(.Item("Dtl_Rcv_Q1")) Then
                    txtSQL = txtSQL & "Dtl_Rcv_Q1=" & ((0 + RcvQ1)) & ","
                    RcvQ1 = (0 + RcvQ1)                    '   ��� 2-1-57  kritpon
                Else
                    txtSQL = txtSQL & "Dtl_Rcv_Q1=" & ((.Item("Dtl_Rcv_Q1") + RcvQ1)) & ","
                    RcvQ1 = (.Item("Dtl_Rcv_Q1") + RcvQ1)  '   ��� 2-1-57  kritpon
                End If

                If IsDBNull(.Item("Dtl_Rcv_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Rcv_Q2=" & ((0 + (RcvQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Rcv_Q2=" & (.Item("Dtl_Rcv_Q1") + RcvQ1) * wStk & ","
                End If

                If IsDBNull(.Item("Dtl_iss_Q1")) Then
                    txtSQL = txtSQL & "Dtl_iSS_Q1=" & ((0 + IssQ1)) & ","
                    IssQ1 = ((0 + IssQ1))        '   ��� 2-1-57  kritpon
                Else
                    txtSQL = txtSQL & "Dtl_iSS_Q1=" & ((.Item("Dtl_iss_Q1") + IssQ1)) & ","
                    IssQ1 = ((.Item("Dtl_iss_Q1") + IssQ1))  '   ��� 2-1-57  kritpon
                End If

                If IsDBNull(.Item("Dtl_iss_Q2")) Then
                    txtSQL = txtSQL & "Dtl_iSS_Q2=" & ((0 + (IssQ1 * wStk))) & ","

                Else
                    txtSQL = txtSQL & "Dtl_iSS_Q2=" & ((.Item("Dtl_iss_Q1") + IssQ1)) * wStk & ","

                End If
                '  �������ǡѺ�ʹ¡��  �ʹ¡������ update 
                'txtSQL = txtSQL & "Dtl_LS_Q1=" & Dtl_LS_Q1 & ","
                'txtSQL = txtSQL & "Dtl_LS_Q2=" & (Dtl_LS_Q1 * wStk) & ","
                'txtSQL = txtSQL & "Dtl_LS_Q1=0,"   '& (((Dtl_LS_Q1 + dtl_Bal_Q1 + RcvQ1) - IssQ1) * -1) & ","
                'txtSQL = txtSQL & "Dtl_LS_Q2=0,"   '& (((Dtl_Bal_q1 + RcvQ1) - IssQ1) * -1) * wStk & ","

                '==================================================================
                If ((Dtl_LS_Q1 + Dtl_Bal_q1 + RcvQ1) - IssQ1) > 0 Then

                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & (((Dtl_LS_Q1 + RcvQ1) - IssQ1)) & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q2=" & (((Dtl_LS_Q1 + RcvQ1) - IssQ1)) * wStk & " "

                Else
                    MsgBox("������ʵ�͡�ջѭ�� �ô�� ICT ��ǹ ", MsgBoxStyle.Critical, "����͹")

                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & (((Dtl_LS_Q1 + RcvQ1) - IssQ1)) & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q2=" & (((Dtl_LS_Q1 + RcvQ1) - IssQ1)) * wStk & " "

                End If

                If PenDingUpdate = True Then
                    txtSQL = txtSQL & ",Dtl_Pen_Q1=" & ((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) & ","
                    txtSQL = txtSQL & "Dtl_Pen_Q2=" & (((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) * wStk) & " "
                End If

                txtSQL = txtSQL & "Where Dtl_Store='110098' "
                txtSQL = txtSQL & "And Dtl_Code='" & StkCode & "' "
                txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "' "

            End With

        Else

            txtSQL = "Insert Into StkDetl "
            txtSQL = txtSQL & "("
            txtSQL = txtSQL & "Dtl_Wh,Dtl_Store,Dtl_Code,"

            txtSQL = txtSQL & "Dtl_Or_Q1,Dtl_Or_Q2,"    ' �ͧ
            txtSQL = txtSQL & "Dtl_Rcv_Q1,Dtl_Rcv_Q2,"  ' ��Ե
            txtSQL = txtSQL & "Dtl_ISS_Q1,Dtl_ISS_Q2,"  ' ���

            txtSQL = txtSQL & "Dtl_LS_Q1,Dtl_LS_Q2, "   ' ¡��
            txtSQL = txtSQL & "Dtl_Bal_Q1,Dtl_Bal_Q2 "

            'txtSQL = txtSQL & "Dtl_Pen_Q1,Dtl_Pen_Q2 "
            txtSQL = txtSQL & ")"
            txtSQL = txtSQL & " Values"
            txtSQL = txtSQL & "('" & whCode & "',"
            txtSQL = txtSQL & "'110098','" & StkCode & "',"

            txtSQL = txtSQL & (OrderQ1) & "," & (((OrderQ1 * wStk))) & ","   '  �ͧ
            txtSQL = txtSQL & (RcvQ1) & "," & (RcvQ1 * wStk) & ","   '  ��Ե
            txtSQL = txtSQL & (IssQ1) & "," & (IssQ1 * wStk) & ","   '  ���

            txtSQL = txtSQL & 0 & "," & 0 & ","     '  ¡��
            txtSQL = txtSQL & (RcvQ1 - IssQ1) & "," & ((RcvQ1 - IssQ1)) * wStk & " " '  Stock

            'If (RcvQ1 - IssQ1) < 0 Then

            'Else
            '    txtSQL = txtSQL & 0 & "," & 0 & ","  '  ¡��
            '    txtSQL = txtSQL & (RcvQ1 - IssQ1) & "," & ((RcvQ1 - IssQ1) * wStk) & ","     '  Stock
            'End If
            'txtSQL = txtSQL & ((dbTools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) & "," 'pen_q1
            'txtSQL = txtSQL & (((dbTools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) * wStk) & " "  '  pen

            txtSQL = txtSQL & ")"

        End If
        Call DBtools.dbSaveSQLsrv(txtSQL, "")
        subDS3 = Nothing
        subDA3 = Nothing

    End Sub

    Sub updateStock2(ByVal CusID As String, ByVal whCode As String, ByVal StkCode As String, ByVal OrderQ1 As Double, ByVal RcvQ1 As Double, ByVal IssQ1 As Double, ByVal PenDingUpdate As Boolean)

        Dim subDA3 As New SqlClient.SqlDataAdapter
        Dim subDS3 As New DataSet

        Dim wStk As Double = DBtools.getStkWight(StkCode)

        Dim Dtl_Bal_q2 As Double = 0
        Dim Dtl_Bal_q1 As Double = 0
        Dim Dtl_LS_Q1 As Double = 0
        Dim Dtl_LS_Q2 As Double = 0



        txtSQL = "Select * From StkDetl "
        txtSQL = txtSQL & "Where Dtl_Code='" & StkCode & "' "
        txtSQL = txtSQL & "And  Dtl_Store='" & CusID & "' "
        txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "'"

        subDA3 = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA3.Fill(subDS3, "chkStk")


        If subDS3.Tables("chkStk").Rows.Count > 0 Then
            With subDS3.Tables("chkStk").Rows(0)

                If IsDBNull(.Item("Dtl_Bal_q1")) Then
                    Dtl_Bal_q1 = 0
                Else
                    Dtl_Bal_q1 = .Item("Dtl_Bal_q1")
                End If
                If IsDBNull(.Item("Dtl_Bal_q2")) Then
                    Dtl_Bal_q2 = 0
                Else
                    Dtl_Bal_q2 = .Item("Dtl_Bal_q2")
                End If

                If IsDBNull(.Item("Dtl_LS_Q1")) Then
                    Dtl_LS_Q1 = 0
                Else
                    Dtl_LS_Q1 = .Item("Dtl_LS_Q1")
                End If
                If IsDBNull(.Item("Dtl_LS_Q2")) Then
                    Dtl_LS_Q2 = 0
                Else
                    Dtl_LS_Q2 = .Item("Dtl_LS_Q2")
                End If

                txtSQL = "Update StkDetl Set "
                If IsDBNull(.Item("Dtl_Or_Q1")) Then
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & ((0 + OrderQ1)) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & ((.Item("Dtl_Or_Q1") + OrderQ1)) & ","
                End If

                If IsDBNull(.Item("Dtl_Or_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Or_Q2=" & ((0 + (OrderQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Or_Q2=" & ((.Item("Dtl_Or_Q2") + (OrderQ1 * wStk))) & ","
                End If

                If IsDBNull(.Item("Dtl_Rcv_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Rcv_Q1=" & ((0 + RcvQ1)) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Rcv_Q1=" & ((.Item("Dtl_Rcv_Q2") + RcvQ1)) & ","
                End If

                If IsDBNull(.Item("Dtl_Rcv_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Rcv_Q2=" & ((0 + (RcvQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Rcv_Q2=" & ((.Item("Dtl_Rcv_Q2") + (RcvQ1 * wStk))) & ","
                End If

                If IsDBNull(.Item("Dtl_iss_Q1")) Then
                    txtSQL = txtSQL & "Dtl_iSS_Q1=" & ((0 + IssQ1)) & ","
                Else
                    txtSQL = txtSQL & "Dtl_iSS_Q1=" & ((.Item("Dtl_iss_Q1") + IssQ1)) & ","
                End If

                If IsDBNull(.Item("Dtl_iss_Q2")) Then
                    txtSQL = txtSQL & "Dtl_iSS_Q2=" & ((0 + (IssQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_iSS_Q2=" & ((.Item("Dtl_iss_Q2") + (IssQ1 * wStk))) & ","
                End If

                If ((Dtl_LS_Q1 + Dtl_Bal_q1 + RcvQ1) - IssQ1) > 0 Then
                    txtSQL = txtSQL & "Dtl_LS_Q1=0" & ","
                    txtSQL = txtSQL & "Dtl_LS_Q2=0" & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & (Dtl_LS_Q1 + Dtl_Bal_q1 + RcvQ1) - IssQ1 & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q2=" & (((Dtl_LS_Q2 + Dtl_Bal_q2 + RcvQ1) - IssQ1) * wStk) & " "
                Else
                    txtSQL = txtSQL & "Dtl_LS_Q1=" & (((.Item("Dtl_LS_Q1") + .Item("Dtl_Bal_q1") + RcvQ1) - IssQ1) * -1) & ","
                    txtSQL = txtSQL & "Dtl_LS_Q2=0" & (((.Item("Dtl_LS_Q1") + .Item("Dtl_Bal_q1") + RcvQ1) - IssQ1) * -1) * wStk & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & ((.Item("Dtl_LS_Q1") + .Item("Dtl_Bal_q1") + RcvQ1) - IssQ1) * -1 & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q2=" & ((((.Item("Dtl_LS_Q2") + .Item("Dtl_Bal_q2") + RcvQ1) - IssQ1)) * wStk) * -1 & " "
                End If

                If PenDingUpdate = True Then
                    txtSQL = txtSQL & ",Dtl_Pen_Q1=" & ((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) & ","
                    txtSQL = txtSQL & "Dtl_Pen_Q2=" & (((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) * wStk) & " "
                End If
                txtSQL = txtSQL & "Where Dtl_Store='" & CusID & "' "
                txtSQL = txtSQL & "And Dtl_Code='" & StkCode & "' "
                txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "' "

            End With

        Else
            txtSQL = "Insert Into StkDetl "
            txtSQL = txtSQL & "("
            txtSQL = txtSQL & "Dtl_Wh,Dtl_Store,Dtl_Code,"

            txtSQL = txtSQL & "Dtl_Or_Q1,Dtl_Or_Q2,"    ' �ͧ
            txtSQL = txtSQL & "Dtl_Rcv_Q1,Dtl_Rcv_Q2,"  ' ��Ե
            txtSQL = txtSQL & "Dtl_ISS_Q1,Dtl_ISS_Q2,"  ' ���

            txtSQL = txtSQL & "Dtl_LS_Q1,Dtl_LS_Q2, "   ' ¡��
            txtSQL = txtSQL & "Dtl_Bal_Q1,Dtl_Bal_Q2,"

            txtSQL = txtSQL & "Dtl_Pen_Q1,Dtl_Pen_Q2 "
            txtSQL = txtSQL & ")"
            txtSQL = txtSQL & " Values"
            txtSQL = txtSQL & "('" & whCode & "',"
            txtSQL = txtSQL & "'" & CusID & "','" & StkCode & "',"

            txtSQL = txtSQL & (OrderQ1) & "," & (((OrderQ1 * wStk))) & ","   '  �ͧ
            txtSQL = txtSQL & (RcvQ1) & "," & (RcvQ1 * wStk) & ","   '  ��Ե
            txtSQL = txtSQL & (IssQ1) & "," & (IssQ1 * wStk) & ","   '  ���

            If (RcvQ1 - IssQ1) < 0 Then
                txtSQL = txtSQL & (RcvQ1 - IssQ1) * -1 & "," & ((RcvQ1 - IssQ1) * -1) * wStk & ","  '  ¡��
                txtSQL = txtSQL & 0 & "," & 0 & ","     '  Stock
            Else
                txtSQL = txtSQL & 0 & "," & 0 & ","  '  ¡��
                txtSQL = txtSQL & (RcvQ1 - IssQ1) & "," & ((RcvQ1 - IssQ1) * wStk) & ","     '  Stock
            End If
            txtSQL = txtSQL & ((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) & "," 'pen_q1
            txtSQL = txtSQL & (((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) * wStk) & " "  '  pen

            txtSQL = txtSQL & ")"

        End If
        Call DBtools.dbSaveSQLsrv(txtSQL, "")
        subDS3 = Nothing
        subDA3 = Nothing

    End Sub

    Sub updateStkdetl(ByVal dtlType As String, ByVal dtlStr As String, ByVal dtlWH As String, ByVal dtlCode As String, ByVal dtlQty As Integer)

        Dim f As String = ""
        'Dim stkCode As String = "010103001230184000001011"
        Dim subDaZ As SqlClient.SqlDataAdapter
        Dim subDsZ As DataSet = New DataSet
        Dim iLoop As Integer = 0
        'Dim strcode As String = "110098"
        'f = dbTools.chkDtlFlag(stkcode, strcode)
        DBtools.openDB()
        txtSQL = "Select * from StkDetl "
        txtSQL = txtSQL & "where Dtl_Code='" & dtlCode & "' "
        txtSQL = txtSQL & "And Dtl_Store='" & dtlStr & "' "
        txtSQL = txtSQL & "And Dtl_Wh='" & dtlWH & "' "

        subDaZ = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDaZ.Fill(subDsZ, "DtlSet")

        Dim issQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_iSS_Q1")
        Dim rcvQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_RCV_Q1")
        Dim lsQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_LS_Q1")
        Dim orQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_Or_Q1")
        Dim balQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_Bal_Q1")
        Dim penQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_Pen_Q1")


        If subDsZ.Tables("DtlSet").Rows.Count > 0 Then

            txtSQL = "Update StkDetl Set "
            Select Case dtlType
                Case "S"
                    txtSQL = txtSQL & "Dtl_iss_Q1=" & issQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Pen_Q1=" & orQ1 + dtlQty - issQ1 & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 - dtlQty & " "
                Case "D"
                    txtSQL = txtSQL & "Dtl_RCV_Q1=" & rcvQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 + dtlQty & " "
                Case "G"
                    txtSQL = txtSQL & "Dtl_RCV_Q1=" & rcvQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 + dtlQty & " "
                Case "B"
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & orQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Pen_Q1=" & orQ1 + dtlQty - issQ1 & " "
                Case "F"
                    txtSQL = txtSQL & "Dtl_RCV_Q1=" & rcvQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 + dtlQty & " "
                Case "Z"
                    txtSQL = txtSQL & "Dtl_iss_Q1=" & issQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 - dtlQty & " "

            End Select

            txtSQL = txtSQL & "Where Dtl_Store='" & dtlStr & "' "
            txtSQL = txtSQL & "And Dtl_Code='" & dtlCode & "' "
            txtSQL = txtSQL & "And Dtl_Wh='" & dtlWH & "' "

            DBtools.dbSaveSQLsrv(txtSQL, "")
        End If
        DBtools.closeDB()

    End Sub

    Function getRunID(trhNo As String) As Integer   '  �红����� �ͧ��Ǻ��  ŧ����� ���ͨ���㹡�úѹ�֡ 

        'CusId = Trim(lbCusID.Text)  '  �����١���
        'DocID = Trim(txtNo.Text)    '  �Ţ����͡��â��
        'whCode = lbWHcode.Text      '  ���ʤ�ѧ�Թ���
        'VatType = cboTypeDoc.SelectedValue    '  ������ VAT

        If Len(trhNo) > 0 Then
            trh_RunID = Microsoft.VisualBasic.Right(trhNo, 4)
        ElseIf Len(trhNo) = 9 Then
            'strDoc = Microsoft.VisualBasic.Right(txtNo.Text, 4)
            trh_RunID = Microsoft.VisualBasic.Right(trhNo, 4)
        End If
        If IsNumeric(trh_RunID) = False Then
            trh_RunID = 0
        End If

        Return trh_RunID

    End Function


    'Public Function ceiling(ByVal strvat As Decimal) As Decimal
    '    Math.Ceiling(strvat)
    '    Return strvat
    'End Function
End Module
