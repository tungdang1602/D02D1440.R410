Imports System
'#-------------------------------------------------------------------------------------
'# Created Date: 25/12/2007 2:49:15 PM
'# Created User: Đoàn Như Thanh
'# Modify Date: 25/12/2007 2:49:15 PM
'# Modify User: Đoàn Như Thanh
'#-------------------------------------------------------------------------------------
Public Class D02F1022
    Private _savedOK As Boolean
    Public ReadOnly Property SavedOK() As Boolean
        Get
            Return _savedOK
        End Get
    End Property

#Region "Const of tdbg1"
    Private Const COL1_AssignmentID As Integer = 0   ' Mã tiêu thức
    Private Const COL1_AssignmentName As Integer = 1 ' Tên tiêu thức
    Private Const COL1_PercentAmount As Integer = 2  ' Tỷ lệ
#End Region

#Region "Const of tdbg2"
    Private Const COL2_ChangeNo As Integer = 0         ' ChangeNo
    Private Const COL2_VoucherTypeID As Integer = 1    ' VoucherTypeID
    Private Const COL2_VoucherDesc As Integer = 2      ' VoucherDesc
    Private Const COL2_Serial As Integer = 3           ' Số Seri
    Private Const COL2_RefNo As Integer = 4            ' Số hóa đơn
    Private Const COL2_TransDesc As Integer = 5        ' Diễn giải 
    Private Const COL2_ObjectTypeID As Integer = 6     ' Loại đối tượng
    Private Const COL2_ObjectID As Integer = 7         ' Đối tượng
    Private Const COL2_CurrencyID As Integer = 8       ' CurrencyID
    Private Const COL2_ExchangeRate As Integer = 9     ' Tỷ giá 
    Private Const COL2_DebitAccountID As Integer = 10  ' TK nợ
    Private Const COL2_CreditAccountID As Integer = 11 ' TK có
    Private Const COL2_Amount As Integer = 12          ' Số Tiền
    Private Const COL2_SourceID As Integer = 13        ' Nguồn vốn
    Private Const COL2_CipNo As Integer = 14          ' Mã XDCB
    Private Const COL2_CipID As Integer = 15           ' CipNo
#End Region

#Region "Const of tdbgInfo"
    Private Const COL3_RefID As Integer = 0        ' RefID
    Private Const COL3_RefName As Integer = 1      ' Thông tin
    Private Const COL3_Caption84 As Integer = 2    ' Diễn giải Tiếng Việt
    Private Const COL3_Caption01 As Integer = 3    ' Diễn giải Tiếng Anh
    Private Const COL3_DataType As Integer = 4     ' DataType
    Private Const COL3_DataTypeName As Integer = 5 ' Tên kiểu dữ liệu
    Private Const COL3_IsUse As Integer = 6        ' Sử dụng
#End Region


    Private dtObjectTypeID As DataTable
    Private dtObjectID As DataTable
    Private dtManagementObjTypeID As DataTable
    Private dtManagementObjID As DataTable
    Private iLastCol1 As Integer
    Private iLastCol2 As Integer
    Private myTabRect As Rectangle
    Private _changeNo As String
    Dim dtGrid2 As DataTable

    Dim clsFilterDropdown As Lemon3.Controls.FilterDropdown

    Public Property ChangeNo() As String
        Get
            Return _changeNo
        End Get
        Set(ByVal Value As String)
            _changeNo = Value
        End Set
    End Property

    Private _createUserID As String
    Public Property CreateUserID() As String
        Get
            Return _createUserID
        End Get
        Set(ByVal Value As String)
            _createUserID = Value
        End Set
    End Property

    Private _createDate As String
    Public Property CreateDate() As String
        Get
            Return _createDate
        End Get
        Set(ByVal Value As String)
            _createDate = Value
        End Set
    End Property

    Dim bLoadFormState As Boolean = False
    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            bLoadFormState = True
            LoadInfoGeneral()
            _FormState = value

            'ID: 143164
            clsFilterDropdown = New Lemon3.Controls.FilterDropdown()
            clsFilterDropdown.CheckD91 = True 'Giá trị mặc định True
            clsFilterDropdown.UseFilterDropdown(tdbg2, COL2_CreditAccountID, COL2_DebitAccountID)

            LoadTDCombo()
            LoadTDBDropDown()
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnNext.Enabled = False
                    LoadTDBGrid1("-1")
                    LoadTDBGrid2("-1")
                    LoadTDBGrid3("", 0)
                    chkDepreciationChange_Click(Nothing, Nothing)
                    chkUsingChange_Click(Nothing, Nothing)
                    chkDepreciationTime_Click(Nothing, Nothing)
                    chkDistributeChange_Click(Nothing, Nothing)
                    chkReceiveChange_Click(Nothing, Nothing)
                    chkManagementChange_Click(Nothing, Nothing) 'ID : 214915
                Case EnumFormState.FormEdit
                    LoadEdit()

                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    btnSave.Top = btnNext.Top
                    txtChangeNo.Enabled = False
                    txtChangeNo.Focus()
                Case EnumFormState.FormView
                    LoadEdit()
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    btnSave.Top = btnNext.Top
                    btnSave.Enabled = False
                    txtChangeNo.Enabled = False
                Case EnumFormState.FormEditOther
                    LoadEdit()
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    btnSave.Top = btnNext.Top
                    txtChangeNo.Enabled = False
                    txtChangeNo.Focus()

                    chkUseAccount.Enabled = False
                    For Each ctrl As Control In TabPage1.Controls
                        ReadOnlyControl(ctrl)
                    Next
                    For Each ctrl As Control In TabPage2.Controls
                        If ctrl.Name = "txtDescription" Then
                        Else
                            ReadOnlyControl(ctrl)
                        End If

                    Next
                    ' txtDescription.ReadOnly = False
            End Select
        End Set
    End Property

    Public Sub ReadOnlyControl(ByVal ParamArray obj() As Control)
        For i As Integer = 0 To obj.Length - 1
            If TypeOf (obj(i)) Is C1.Win.C1Input.C1DateEdit Then
                Dim ctrl As C1.Win.C1Input.C1DateEdit = CType(obj(i), C1.Win.C1Input.C1DateEdit)
                ctrl.ReadOnly = True
            ElseIf TypeOf (obj(i)) Is TextBox Then
                Dim ctrl As TextBox = CType(obj(i), TextBox)
                ctrl.ReadOnly = True
            ElseIf TypeOf (obj(i)) Is C1.Win.C1List.C1Combo Then
                Dim ctrl As C1.Win.C1List.C1Combo = CType(obj(i), C1.Win.C1List.C1Combo)
                ctrl.ReadOnly = True

            ElseIf TypeOf (obj(i)) Is System.Windows.Forms.GroupBox Then
                Dim ctrl As System.Windows.Forms.GroupBox = CType(obj(i), System.Windows.Forms.GroupBox)
                ctrl.Enabled = False
            ElseIf TypeOf (obj(i)) Is System.Windows.Forms.CheckBox Then
                Dim ctrl As System.Windows.Forms.CheckBox = CType(obj(i), System.Windows.Forms.CheckBox)
                ctrl.Enabled = False
            ElseIf TypeOf (obj(i)) Is System.Windows.Forms.TextBox Then
                Dim ctrl As System.Windows.Forms.TextBox = CType(obj(i), System.Windows.Forms.TextBox)
                ctrl.ReadOnly = True
            ElseIf TypeOf (obj(i)) Is C1.Win.C1TrueDBGrid.C1TrueDBGrid Then
                Dim ctrl As C1.Win.C1TrueDBGrid.C1TrueDBGrid = CType(obj(i), C1.Win.C1TrueDBGrid.C1TrueDBGrid)
                ctrl.Enabled = False
            End If

            obj(i).TabStop = False
        Next
    End Sub

    Private Sub D02F1022_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '12/10/2020, id 144622-Tài sản cố định_Lỗi chưa cảnh báo khi lưu
        If _FormState = EnumFormState.FormEdit Then
            If Not _savedOK Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        ElseIf _FormState = EnumFormState.FormAdd Then
            If (txtChangeNo.Text <> "" Or tdbcVoucherTypeID.Text <> "" Or tdbcCurrencyID.Text <> "" Or tdbcDescPriorityID.Text <> "") Then
                If Not _savedOK Then
                    If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub D02F1022_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
            Exit Sub
        End If

        If e.KeyCode = Keys.F11 Then
            If tabMain.SelectedIndex = 0 Then
                HotKeyF11(Me, tdbg1)
                Exit Sub
            Else
                HotKeyF11(Me, tdbg2)
                Exit Sub
            End If

        End If

        If e.Alt Then
            If e.KeyCode = Keys.NumPad1 Or e.KeyCode = Keys.D1 Then
                Application.DoEvents()
                tabMain.SelectedTab = TabPage1
                chkDepreciationChange.Focus()
                Application.DoEvents()
                Exit Sub
            ElseIf e.KeyCode = Keys.NumPad2 Or e.KeyCode = Keys.D2 Then
                Application.DoEvents()
                tabMain.SelectedTab = TabPage2
                tdbcVoucherTypeID.Focus()
                Application.DoEvents()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub D02F1022_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If bLoadFormState = False Then FormState = _FormState
        Me.Cursor = Cursors.WaitCursor
        Loadlanguage()
        SetBackColorObligatory()
        tdbg1_NumberFormat()
        tdbg2_NumberFormat()
        iLastCol1 = CountCol(tdbg1, tdbg1.Splits.Count - 1)
        iLastCol2 = CountCol(tdbg2, tdbg2.Splits.Count - 1)
        tabMain.DrawMode = TabDrawMode.OwnerDrawFixed     'Bắt buộc
        AddHandler tabMain.DrawItem, AddressOf OnDrawItem  'Bắt buộc
        'ID 96202 17.05.2017
        HideControlBySystem()
        '****************************
        CheckIdTextBox(txtChangeNo)
        InputbyUnicode(Me, gbUnicode)
        tdbgInfo_LockedColumns()
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub OnDrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        ' Create pen.
        Dim blackPen As New Pen(tabMain.TabPages(0).BackColor, 3)
        'Get Location tabpage
        myTabRect = tabMain.GetTabRect(tabMain.SelectedIndex)
        ' Create coordinates of points that define line.
        Dim x1 As Integer = myTabRect.X
        Dim y1 As Integer = myTabRect.Bottom
        Dim x2 As Integer = myTabRect.X + myTabRect.Width
        ' Draw line to screen.
        e.Graphics.DrawLine(blackPen, x1, y1, x2, y1)
        '**************
        Dim page As TabPage = tabMain.TabPages(e.Index)
        If Not page.Enabled Then
            Dim brush As New SolidBrush(SystemColors.GrayText)
            e.Graphics.DrawString(page.Text, page.Font, brush, e.Bounds)
        Else
            Dim brush As New SolidBrush(page.ForeColor)
            e.Graphics.DrawString(page.Text, page.Font, brush, e.Bounds)
        End If

    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Thiet_lap_nghiep_vu_tac_dong_-_D02F1022") & UnicodeCaption(gbUnicode) 'ThiÕt lËp nghiÖp vó tÀc ¢èng - D02F1022
        '================================================================ 
        lblChangeNo.Text = rL3("Ma_nghiep_vu") 'Mã nghiệp vụ
        lblChangeName.Text = rL3("Ten_nghiep_vu") 'Tên nghiệp vụ

        lblEmployeeID.Text = rL3("Nguoi_tiep_nhan") 'Người tiếp nhận
        lblServiceLife.Text = rL3("So_ky_khau_hao") 'Số kỳ khấu hao
        lblNote1.Text = rL3("Ghi_chu_1")  'Ghi chú 1
        lblNote2.Text = rL3("Ghi_chu_2") 'Ghi chú 2
        lblNote3.Text = rL3("Ghi_chu_3")  'Ghi chú 3
        lblVoucherTypeID.Text = rL3("Loai_phieu") 'Loại phiếu
        lblCurrencyID.Text = rL3("Loai_tien") 'Loại tiền
        lblDescription.Text = rL3("Dien_giai") 'Diễn giải
        lblLocationID.Text = rL3("Vi_tri") 'Vị trí
        lblSetVoucherInfo.Text = rL3("Thiet_lap_thong_tin_phieu") 'Thiết lập thông tin phiếu
        lblDescPriorityID.Text = rL3("Thu_tu_uu_tien") 'Thứ tự ưu tiên

        '================================================================ 
        lblManagementObjTypeID.Text = rL3("Bo_phan_quan_ly") 'Bộ phận quản lý
        lblObjectTypeID.Text = rL3("Bo_phan_tiep_nhan") 'Bộ phận tiếp nhận

        '================================================================ 
        grpSetDescription.Text = rL3("Thiet_lap_uu_tien_lay_dien_giai_khi_Tap_hop") 'Thiết lập ưu tiên lấy diễn giải khi Tập hợp
        '================================================================ 
        tdbcDescPriorityID.Columns("DescPriorityID").Caption = rL3("Ma") 'Mã
        tdbcDescPriorityID.Columns("DescPriorityName").Caption = rL3("Ten") 'Tên

        '================================================================ 
        btnSave.Text = rL3("_Luu") '&Lưu
        btnNext.Text = rL3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rL3("Do_ng") 'Đó&ng
        '================================================================ 
        chkDisabled.Text = rL3("Khong_su_dung") 'Không sử dụng
        chkUseAccount.Text = rL3("Co_kem_theo_tac_dong_tai_chinh") 'Có kèm theo tác động tài chính
        chkIsEliminated.Text = rL3("Thanh_ly_tai_san") 'Thanh lý tài sản
        chkReceiveChange.Text = rL3("Thay_doi_bo_phan_tiep_nhan") 'Thay đổi bộ phận tiếp nhận
        chkDistributeChange.Text = rL3("Thay_doi_tieu_thuc_phan_bo") 'Thay đổi tiêu thức phân bổ
        chkDepreciationTime.Text = rL3("Thay_doi_thoi_gian_khau_hao") 'Thay đổi thời gian khấu hao
        chkUsingChange.Text = rL3("Thay_doi_tinh_trang_su_dung") 'Thay đổi tình trạng sử dụng
        chkDepreciationChange.Text = rL3("Thay_doi_tinh_trang_khau_hao") 'Thay đổi tình trạng khấu hao
        chkIsChangeAssetAccount.Text = rL3("Thay_doi_tai_khoan_tai_san") 'Thay đổi tài khoản tài sản
        chkIsChangeDepAccount.Text = rL3("Thay_doi_tai_khoan_khau_hao") 'Thay đổi tài khoản khấu hao
        '================================================================ 
        optReUse.Text = rL3("Tai_su_dung") 'Tái sử dụng
        optStopUse.Text = rL3("Ngung_su_dung") 'Ngưng sử dụng
        optReDepreciation.Text = rL3("Tai_khau_hao") 'Tái khấu hao
        optStopDepreciation.Text = rL3("Ngung_khau_hao") 'Ngưng khấu hao        
        '================================================================ 
        TabPage1.Text = "1. " & rL3("Phi_tai_chinh") & " " 'Phi tài chính
        TabPage2.Text = "2. " & rL3("Tai_chinh") & " " 'Tài chính
        TabPage3.Text = "3. " & rL3("Thong_tin_bo_sung") & " "
        '================================================================ 
        tdbcLocationID.Columns("LocationID").Caption = rL3("Ma") 'Mã
        tdbcLocationID.Columns("LocationName").Caption = rL3("Ten") 'Tên
        tdbcEmployeeID.Columns("EmployeeID").Caption = rL3("Ma") 'Mã
        tdbcEmployeeID.Columns("EmployeeName").Caption = rL3("Ten") 'Tên
        tdbcObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        tdbcObjectTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbcObjectTypeID.Columns("ObjectTypeName").Caption = rL3("Ten") 'Tên
        tdbcCurrencyID.Columns("CurrencyID").Caption = rL3("Ma") 'Mã
        tdbcCurrencyID.Columns("CurrencyName").Caption = rL3("Ten") 'Tên
        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rL3("Ma") 'Mã
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rL3("Ten") 'Tên
        tdbcManagementObjID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbcManagementObjID.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        '================================================================ 
        tdbdAssignmentID.Columns("AssignmentID").Caption = rL3("Ma") 'Mã 
        tdbdAssignmentID.Columns("AssignmentName").Caption = rL3("Ten") 'Tên
        tdbdAssignmentID.Columns("DebitAccountID").Caption = rL3("Tai_khoan_no") 'Tài khoản nợ
        tdbdCipID.Columns("CipNo").Caption = rL3("Ma") 'Mã
        tdbdCipID.Columns("CipName").Caption = rL3("Ten") 'Tên
        tdbdSourceID.Columns("SourceID").Caption = rL3("Ma") 'Mã
        tdbdSourceID.Columns("SourceName").Caption = rL3("Ten") 'Tên 
        tdbdAmount.Columns("Code").Caption = rL3("Ma") 'Mã
        tdbdAmount.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbdCreditAccountID.Columns("AccountID").Caption = rL3("Ma") 'Mã
        tdbdCreditAccountID.Columns("AccountName").Caption = rL3("Ten") 'Tên
        tdbdDebitAccountID.Columns("AccountID").Caption = rL3("Ma") 'Mã
        tdbdDebitAccountID.Columns("AccountName").Caption = rL3("Ten") 'Tên
        tdbdObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbdObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        tdbdObjectTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbdObjectTypeID.Columns("ObjectTypeName").Caption = rL3("Ten") 'Tên
        '================================================================ 
        tdbg1.Columns("AssignmentID").Caption = rL3("Ma_tieu_thuc") 'Mã tiêu thức
        tdbg1.Columns("AssignmentName").Caption = rL3("Ten_tieu_thuc") 'Tên tiêu thức
        tdbg1.Columns("PercentAmount").Caption = rL3("Ty_le") 'Tỷ lệ
        tdbg2.Columns("RefNo").Caption = rL3("So_hoa_don") 'Số hóa đơn
        tdbg2.Columns("Serial").Caption = rL3("So_Seri") 'Số Seri
        tdbg2.Columns("TransDesc").Caption = rL3("Dien_giai") 'Diễn giải 
        tdbg2.Columns("ObjectTypeID").Caption = rL3("Loai_doi_tuong") 'Loại đối tượng
        tdbg2.Columns("ObjectID").Caption = rL3("Doi_tuong") 'Đối tượng
        tdbg2.Columns("ExchangeRate").Caption = rL3("Ty_gia") 'Tỷ giá 
        tdbg2.Columns("DebitAccountID").Caption = rL3("TK_no") 'TK nợ
        tdbg2.Columns("CreditAccountID").Caption = rL3("TK_co") 'TK có
        tdbg2.Columns("Amount").Caption = rL3("So_tien") 'Số Tiền
        tdbg2.Columns("SourceID").Caption = rL3("Nguon_von") 'Nguồn vốn
        tdbg2.Columns("CipNo").Caption = rL3("Ma_XDCB") 'Mã XDCB
        '================================================================ 
        tdbgInfo.Columns(COL3_RefName).Caption = rL3("Thong_tin") 'Thông tin
        tdbgInfo.Columns(COL3_Caption84).Caption = rL3("Dien_giai_tieng_Viet") 'Diễn giải Tiếng Việt
        tdbgInfo.Columns(COL3_Caption01).Caption = rL3("Dien_giai_tieng_Anh") 'Diễn giải Tiếng Anh
        tdbgInfo.Columns(COL3_DataTypeName).Caption = rL3("Ten_kieu_du_lieu") 'Tên kiểu dữ liệu
        tdbgInfo.Columns(COL3_IsUse).Caption = rL3("Su_dung") 'Sử dụng

        lblAssetWHID.Text = rL3("Kho")
        tdbcAssetWHID.Columns("WareHouseID").Caption = rL3("Ma")
        tdbcAssetWHID.Columns("WareHouseName").Caption = rL3("Ten")
    End Sub

    Private Sub SetBackColorObligatory()
        txtChangeNo.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCurrencyID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcDescPriorityID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbg1.Splits(SPLIT0).DisplayColumns(COL1_AssignmentID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub tdbgInfo_LockedColumns()
        tdbgInfo.Splits(SPLIT0).DisplayColumns(COL3_RefName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbgInfo.Splits(SPLIT0).DisplayColumns(COL3_DataTypeName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub


    Private Sub tdbg1_NumberFormat()
        tdbg1.Columns(COL1_PercentAmount).NumberFormat = DxxFormat.DefaultNumber2
    End Sub

    Private Sub tdbg2_NumberFormat()
        tdbg2.Columns(COL2_ExchangeRate).NumberFormat = DxxFormat.ExchangeRateDecimals
        tdbg2.Columns(COL2_Amount).NumberFormat = DxxFormat.DecimalPlaces
    End Sub

    Private Sub LoadTDCombo()
        '************** Tab Phi tài chính ************** 

        'Combo ObjectTypeID
        Dim sSQL As String = ""
        dtObjectTypeID = ReturnTableObjectTypeID(gbUnicode)
        LoadDataSource(tdbcObjectTypeID, dtObjectTypeID.Copy, gbUnicode)

        dtManagementObjTypeID = ReturnTableObjectTypeID(gbUnicode)
        LoadDataSource(tdbcManagementObjTypeID, dtManagementObjTypeID.Copy, gbUnicode) 'IDS

        'Combo ObjectID
        sSQL = "Select     ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " As ObjectName, ObjectTypeID " & vbCrLf
        sSQL &= "From       Object WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where      Disabled = 0 " & vbCrLf
        sSQL &= "Order By   ObjectID " & vbCrLf
        dtObjectID = ReturnDataTable(sSQL)
        dtManagementObjID = ReturnDataTable(sSQL)
        LoadDataSource(tdbcObjectID, dtObjectID.Copy, gbUnicode)
        LoadDataSource(tdbcManagementObjID, dtManagementObjID.Copy, gbUnicode) 'IDS


        'Combo EmployeeID
        sSQL = "Select ObjectID As EmployeeID, ObjectName" & UnicodeJoin(gbUnicode) & " As EmployeeName, ObjectTypeID, VATNo " & vbCrLf
        sSQL &= "From Object WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where ObjectTypeID = 'NV' " & vbCrLf
        sSQL &= "Order By ObjectID " & vbCrLf
        LoadDataSource(tdbcEmployeeID, sSQL, gbUnicode)


        '************** Tab tài chính ************** 

        'Combo VoucherTypeID
        LoadVoucherTypeID(tdbcVoucherTypeID, D02, , gbUnicode)

        'Combo CurrencyID
        sSQL = "Select     CurrencyID, CurrencyName" & UnicodeJoin(gbUnicode) & " As CurrencyName, ExchangeRate " & vbCrLf
        sSQL &= "From       D91T0010 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where      Disabled = 0 " & vbCrLf
        sSQL &= "Order By   CurrencyID " & vbCrLf
        LoadDataSource(tdbcCurrencyID, sSQL, gbUnicode)

        sSQL = "Select     WareHouseID, RTrim(LTrim(WareHouseName" & UnicodeJoin(gbUnicode) & ")) As WareHouseName " & vbCrLf
        sSQL &= "From       D07T0007 WITH(NOLOCK)" & vbCrLf
        sSQL &= "Where      DivisionID = " & SQLString(gsDivisionID) & " And Disabled = 0" & vbCrLf
        sSQL &= "           And (DAGroupID = '' Or DAGroupID In" & vbCrLf
        sSQL &= "                   (" & vbCrLf
        sSQL &= "                       Select  DAGroupID " & vbCrLf
        sSQL &= "                       From    LemonSys.Dbo.D00V0080" & vbCrLf
        sSQL &= "                       Where   UserID = " & SQLString(gsUserID) & vbCrLf
        sSQL &= "                               Or 'LEMONADMIN' = " & SQLString(gsUserID) & vbCrLf
        sSQL &= "                   )" & vbCrLf
        sSQL &= "               )" & vbCrLf
        sSQL &= "Order by   WareHouseID" & vbCrLf
        LoadDataSource(tdbcAssetWHID, sSQL, gbUnicode)

        'Load Location
        sSQL = "-- Combo Vi tri " & vbCrLf
        sSQL &= " SELECT		 LookupID As LocationID, Description" & UnicodeJoin(gbUnicode) & " As LocationName"
        sSQL &= " FROM 		D91T0320 WITH(NOLOCK) "
        sSQL &= " WHERE 		LookupType = 'D02_Position' "
        sSQL &= " And (DAGroupID =  ''  Or DAGroupID "
        sSQL &= " IN (Select DAGroupID From lemonsys.dbo.D00V0080 Where UserID= " & SQLString(gsUserID) & " ) Or 'LEMONADMIN' = 'LEMONADMIN')"
        sSQL &= " Order By		 LookupID"
        LoadDataSource(tdbcLocationID, sSQL, gbUnicode)

        'Load tdbcDescPriorityID
        LoadDataSource(tdbcDescPriorityID, SQLStoreD02P2041, gbUnicode)

        '30/10/2019, Lê Thị Phú Hà:id 123376-Bổ sung thay đổi bộ phận quản lý (NV tác động D02)
        'Load tdbcManagementObjID
        'LoadDataSource(tdbcManagementObjID, ReturnTableFilter(dtObjectID, "ObjectTypeID = 'DV'", True), gbUnicode)

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2041
    '# Created User: NGOCTHOAI
    '# Created Date: 20/06/2017 04:45:43
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2041() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon combo Thu tu uu tien " & vbCrLf)
        sSQL &= "Exec D02P2041 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) 'Language, varchar[20], NOT NULL
        Return sSQL
    End Function


    Private Sub LoadTDBDropDown()
        '************** Tab Phi tài chính ************** 

        'DropDown AssignmentID
        Dim sSQL As String = ""
        sSQL = "Select  AssignmentID, AssignmentName" & UnicodeJoin(gbUnicode) & " As AssignmentName, DebitAccountID " & vbCrLf
        sSQL &= "From   D02T0002 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where  Disabled = 0 " & vbCrLf
        LoadDataSource(tdbdAssignmentID, sSQL, gbUnicode)

        '************** Tab tài chính ************** 

        'DropDown ObjectTypeID
        LoadDataSource(tdbdObjectTypeID, dtObjectTypeID.Copy, gbUnicode)

        'DropDown ObjectID
        LoadDataSource(tdbdObjectID, dtObjectID.Copy, gbUnicode)

        'Combo DebitAccountID và CreditAccountID
        Dim dt As DataTable
        sSQL = "Select      '_TKTS_' As AccountID, "
        If gbUnicode Then
            sSQL &= "N" & SQLString(rL3("Tai_khoan_tai_san")) & " As AccountName" & vbCrLf
            sSQL &= "Union All " & vbCrLf
            sSQL &= "Select     '_TKKH_' As AccountID, "
            sSQL &= "N" & SQLString(rL3("Tai_khoan_khau_hao")) & " As AccountName" & vbCrLf
        Else
            sSQL &= SQLString(ConvertUnicodeToVni(rL3("Tai_khoan_tai_san"))) & " As AccountName" & vbCrLf
            sSQL &= "Union All " & vbCrLf
            sSQL &= "Select     '_TKKH_' As AccountID, "
            sSQL &= SQLString(ConvertUnicodeToVni(rL3("Tai_khoan_khau_hao"))) & " As AccountName" & vbCrLf
        End If
        sSQL &= "Union All " & vbCrLf
        sSQL &= "Select     AccountID, " & vbCrLf
        sSQL &= IIf(gsLanguage = "84", "AccountName", "AccountName01").ToString() & UnicodeJoin(gbUnicode) & " As AccountName " & vbCrLf
        sSQL &= "From       D90T0001 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where      AccountStatus=0 And OffAccount = 0  " & vbCrLf
        sSQL &= "Order By   AccountID " & vbCrLf
        dt = ReturnDataTable(sSQL)
        'LoadDataSource(tdbdDebitAccountID, dt.Copy, gbUnicode)
        'LoadDataSource(tdbdCreditAccountID, dt.Copy, gbUnicode)

        If clsFilterDropdown.IsNewFilter Then
            LoadStringSource(tdbdDebitAccountID, sSQL)
            LoadStringSource(tdbdCreditAccountID, sSQL)
        Else ' Nhập liệu dạng cũ
            LoadDataSource(tdbdDebitAccountID, sSQL, gbUnicode)
            LoadDataSource(tdbdCreditAccountID, sSQL, gbUnicode)
        End If

        'DropDown SourceID
        sSQL = "Select      SourceID, SourceName" & UnicodeJoin(gbUnicode) & " As SourceName" & vbCrLf
        sSQL &= "From       D02T0013 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Order By   SourceID " & vbCrLf
        LoadDataSource(tdbdSourceID, sSQL, gbUnicode)

        'DropDown SourceID
        sSQL = "Select      CipNo, CipName" & UnicodeJoin(gbUnicode) & " As CipName, CipID " & vbCrLf
        sSQL &= "From       D02T0100 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where      Disabled = 0 AND Status = 0 " & vbCrLf
        sSQL &= "Order By   CipNo " & vbCrLf
        LoadDataSource(tdbdCipID, sSQL, gbUnicode)

        'DropDown Amount
        dt.Clear()
        Dim dr As DataRow
        dt.Columns.Add("Code", System.Type.GetType("System.String"))
        dt.Columns.Add("Description", System.Type.GetType("System.String"))

        dr = dt.NewRow
        dr.Item("Code") = "-1"
        dr.Item("Description") = IIf(gbUnicode, ConvertVniToUnicode(rL3("Gia_tri_con_lai_V")), rL3("Gia_tri_con_lai_V")).ToString
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "-2"
        dr.Item("Description") = IIf(gbUnicode, ConvertVniToUnicode(rL3("Muc_khau_hao_V")), rL3("Muc_khau_hao_V")).ToString
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "-3"
        dr.Item("Description") = IIf(gbUnicode, ConvertVniToUnicode(rL3("Hao_mon_luy_ke_V")), rL3("Hao_mon_luy_ke_V")).ToString
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "-4"
        dr.Item("Description") = IIf(gbUnicode, ConvertVniToUnicode(rL3("Gia_tri_con_lai_khau_hao")), rL3("Gia_tri_con_lai_khau_hao")).ToString
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "-5"
        dr.Item("Description") = IIf(gbUnicode, ConvertVniToUnicode(rL3("Gia_tri_con_lai_khong_khau_hao")), rL3("Gia_tri_con_lai_khong_khau_hao")).ToString
        dt.Rows.Add(dr)

        LoadDataSource(tdbdAmount, dt, gbUnicode)
    End Sub

    Private Sub LoadEdit()
        LoadMaster()
        LoadTDBGrid1(_changeNo)
        LoadTDBGrid2(_changeNo)
        LoadTDBGrid3(_changeNo, 1)
    End Sub

    Private Sub LoadMaster()
        Dim dt As New DataTable
        Dim sSQL As New StringBuilder(257)
        sSQL.Append(" Select Distinct ChangeNo, ChangeName, ChangeNameU, Notes1, Notes2, Notes3, Notes1U, Notes2U, Notes3U, Disabled, IsEliminated, UseAccount, ")
        sSQL.Append(" CreateUserID, CreateDate ")
        sSQL.Append(" From D02T0201 WITH(NOLOCK) ")
        sSQL.Append(" Where ChangeNo = " & SQLString(_changeNo))

        dt = ReturnDataTable(sSQL.ToString)
        If dt.Rows.Count > 0 Then
            txtChangeNo.Text = dt.Rows(0).Item("ChangeNo").ToString
            txtChangeName.Text = dt.Rows(0).Item("ChangeName" & UnicodeJoin(gbUnicode)).ToString
            txtNote1.Text = dt.Rows(0).Item("Notes1" & UnicodeJoin(gbUnicode)).ToString
            txtNote2.Text = dt.Rows(0).Item("Notes2" & UnicodeJoin(gbUnicode)).ToString
            txtNote3.Text = dt.Rows(0).Item("Notes3" & UnicodeJoin(gbUnicode)).ToString
            chkDisabled.Checked = Convert.ToBoolean(IIf(dt.Rows(0).Item("Disabled").ToString = "1", True, False))
            chkUseAccount.Checked = Convert.ToBoolean(IIf(dt.Rows(0).Item("UseAccount").ToString = "1", True, False))
            chkIsEliminated.Checked = Convert.ToBoolean(IIf(dt.Rows(0).Item("IsEliminated").ToString = "1", True, False))
        End If

        LoadChangeAsset_DepAccount() '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ

        If chkUseAccount.Checked And (chkIsChangeAssetAccount.Checked = False And chkIsChangeDepAccount.Checked = False) Then
            tabMain.SelectedTab = TabPage2
        Else
            tabMain.SelectedTab = TabPage1
        End If

        '****************** Tab phi tài chính ****************** 

        'Check Thay đổi thời gian khấu hao
        Dim sServiceLife As String = ""
        sSQL.Remove(0, sSQL.Length)
        sSQL.Append(" Select ServiceLife ")
        sSQL.Append(" From D02T0201 WITH(NOLOCK)")
        sSQL.Append(" Where ChangeType = 'SL' And ChangeNo = " & SQLString(_changeNo))

        sServiceLife = ReturnScalar(sSQL.ToString)

        chkDepreciationTime.Checked = sServiceLife <> ""
        txtServiceLife.Text = sServiceLife
        If Not chkDepreciationTime.Checked Then
            txtServiceLife.BackColor = Color.White
            txtServiceLife.Enabled = False
        Else
            txtServiceLife.BackColor = COLOR_BACKCOLOROBLIGATORY
        End If

        'Check Thay đổi tình trạng khấu hao
        Dim sStopDepreciation As String = ""
        sSQL.Remove(0, sSQL.Length)
        sSQL.Append(" Select StopDepreciation ")
        sSQL.Append(" From D02T0201 ")
        sSQL.Append(" Where ChangeType = 'SD' And ChangeNo = " & SQLString(_changeNo))

        sStopDepreciation = ReturnScalar(sSQL.ToString)
        chkDepreciationChange.Checked = sStopDepreciation <> ""
        optStopDepreciation.Checked = sStopDepreciation = "1"
        optReDepreciation.Checked = sStopDepreciation = "0"
        If Not chkDepreciationChange.Checked Then
            optStopDepreciation.Enabled = False
            optReDepreciation.Enabled = False
        End If
        'Check Thay đổi tình trạng sử dụng
        sSQL.Remove(0, sSQL.Length)
        sSQL.Append(" Select StopUse, AssetWHID ")
        sSQL.Append(" From D02T0201 WITH(NOLOCK) ")
        sSQL.Append(" Where ChangeType = 'SU' And ChangeNo = " & SQLString(_changeNo))
        Dim dtStopUse As DataTable
        dtStopUse = ReturnDataTable(sSQL.ToString)
        If dtStopUse.Rows.Count > 0 Then
            chkUsingChange.Checked = dtStopUse.Rows(0).Item("StopUse").ToString <> ""
            optStopUse.Checked = dtStopUse.Rows(0).Item("StopUse").ToString <> "1"
            optReUse.Checked = dtStopUse.Rows(0).Item("StopUse").ToString <> "0"
            tdbcAssetWHID.SelectedValue = dtStopUse.Rows(0).Item("AssetWHID").ToString
        End If
        If Not chkUsingChange.Checked Then
            optStopUse.Enabled = False
            optReUse.Enabled = False
            tdbcAssetWHID.Enabled = False
        End If
        'Check Thay đổi bộ phận quản lí & Check Thay đổi bộ phận tiếp nhận
        sSQL.Remove(0, sSQL.Length)
        dt.Clear()
        sSQL.Append(" Select ObjectTypeID, ObjectID,ManagementObjTypeID ,ManagementObjID, EmployeeID, IsManagement, isReceive, FullName" & UnicodeJoin(gbUnicode))
        sSQL.Append(" From D02T0201 WITH(NOLOCK) ")
        sSQL.Append(" Where ChangeType = 'OB' And IsManagement = 0 And ChangeNo = " & SQLString(_changeNo))

        dt = ReturnDataTable(sSQL.ToString)

        If dt.Rows.Count > 0 Then
            tdbcObjectTypeID.Text = dt.Rows(0).Item("ObjectTypeID").ToString
            tdbcObjectID.SelectedValue = dt.Rows(0).Item("ObjectID").ToString
            tdbcEmployeeID.Text = dt.Rows(0).Item("EmployeeID").ToString
            txtEmployeeName.Text = dt.Rows(0).Item("FullName" & UnicodeJoin(gbUnicode)).ToString

            '30/10/2019, Lê Thị Phú Hà:id 123376-Bổ sung thay đổi bộ phận quản lý (NV tác động D02)
            'ID : 214915
            tdbcManagementObjTypeID.SelectedValue = dt.Rows(0).Item("ManagementObjTypeID").ToString
            tdbcManagementObjID.SelectedValue = dt.Rows(0).Item("ManagementObjID").ToString
            'chkManagementChange.Checked = L3Bool(dt.Rows(0).Item("IsManagement").ToString)
            'chkReceiveChange.Checked = L3Bool(dt.Rows(0).Item("isReceive").ToString)

            chkReceiveChange.Checked = True
            tdbcObjectTypeID.Enabled = chkReceiveChange.Checked
            tdbcObjectID.Enabled = chkReceiveChange.Checked
            tdbcEmployeeID.Enabled = chkReceiveChange.Checked
            tdbcLocationID.Enabled = chkReceiveChange.Checked
            tdbcManagementObjID.Enabled = False
            tdbcManagementObjTypeID.Enabled = False

        End If

        sSQL.Remove(0, sSQL.Length)
        dt.Clear()
        sSQL.Append(" Select ObjectTypeID, ObjectID,ManagementObjTypeID ,ManagementObjID, EmployeeID, IsManagement, IsReceive, FullName" & UnicodeJoin(gbUnicode))
        sSQL.Append(" From D02T0201 WITH(NOLOCK) ")
        sSQL.Append(" Where ChangeType = 'OB' And IsManagement = 1 And ChangeNo = " & SQLString(_changeNo))

        dt = ReturnDataTable(sSQL.ToString)
        If dt.Rows.Count > 0 Then
            tdbcObjectTypeID.Text = dt.Rows(0).Item("ObjectTypeID").ToString
            tdbcObjectID.SelectedValue = dt.Rows(0).Item("ObjectID").ToString
            tdbcEmployeeID.Text = dt.Rows(0).Item("EmployeeID").ToString
            txtEmployeeName.Text = dt.Rows(0).Item("FullName" & UnicodeJoin(gbUnicode)).ToString

            '30/10/2019, Lê Thị Phú Hà:id 123376-Bổ sung thay đổi bộ phận quản lý (NV tác động D02)
            'ID : 214915
            tdbcManagementObjTypeID.SelectedValue = dt.Rows(0).Item("ManagementObjTypeID").ToString
            tdbcManagementObjID.SelectedValue = dt.Rows(0).Item("ManagementObjID").ToString
            chkManagementChange.Checked = L3Bool(dt.Rows(0).Item("IsManagement").ToString)
            chkReceiveChange.Checked = L3Bool(dt.Rows(0).Item("isReceive").ToString)
            tdbcObjectTypeID.Enabled = chkReceiveChange.Checked
            tdbcObjectID.Enabled = chkReceiveChange.Checked
            tdbcManagementObjID.Enabled = chkManagementChange.Checked
            tdbcManagementObjTypeID.Enabled = chkManagementChange.Checked
            tdbcEmployeeID.Enabled = chkReceiveChange.Checked
            tdbcLocationID.Enabled = chkReceiveChange.Checked

        End If

        'ID : 214915
       
        '****************** Tab tài chính ****************** 
        'sSQL.Remove(0, sSQL.Length)
        'dt.Clear()
        'sSQL.Append(" Select  A. ChangeNo, A.VoucherTypeID, A.VoucherDesc" & UnicodeJoin(gbUnicode) & " As VoucherDesc, A.CurrencyID ")
        'sSQL.Append(" From D02T0204 A WITH(NOLOCK) Left Join D02T0100 B WITH(NOLOCK) On A.CipID = B.CipID  ")
        'sSQL.Append(" Where ChangeNo = " & SQLString(_changeNo))
        'sSQL.Append(" Order By RefNo")

        dt = ReturnDataTable(SQLStoreD02P1022) '21/6/2017, Nguyễn Thị Hồng Nhị: id 97491-Bổ sung Thiết lập ưu tiên lấy diễn giải
        If dt.Rows.Count > 0 Then
            tdbcVoucherTypeID.SelectedValue = dt.Rows(0).Item("VoucherTypeID").ToString
            txtDescription.Text = dt.Rows(0).Item("VoucherDesc").ToString
            tdbcCurrencyID.Text = dt.Rows(0).Item("CurrencyID").ToString
            tdbcDescPriorityID.Text = dt.Rows(0).Item("DescPriorityID").ToString '21/6/2017, Nguyễn Thị Hồng Nhị: id 97491-Bổ sung Thiết lập ưu tiên lấy diễn giải
        End If

        'Load Location
        sSQL.Remove(0, sSQL.Length)
        sSQL.Append("-- Do nguon Location" & vbCrLf)
        sSQL.Append(" Select LocationID")
        sSQL.Append(" FROM 	D02T0201 WITH(NOLOCK)  ")
        sSQL.Append(" WHERE 	ChangeType = 'OB' And ChangeNo = " & SQLString(txtChangeNo.Text))
        Dim sLocationID As String = ReturnScalar(sSQL.ToString)
        tdbcLocationID.SelectedValue = sLocationID

    End Sub

    Private Sub LoadChangeAsset_DepAccount()
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        Dim sSQL As String = ""
        sSQL = "SELECT IsChangeAssetAccount FROM D02T0201 WITH(NOLOCK) WHERE ChangeType = 'AAC' AND ChangeNo = " & SQLString(_changeNo)
        chkIsChangeAssetAccount.Checked = L3Bool(ReturnScalar(sSQL))
        sSQL = "SELECT IsChangeDepAccount  FROM D02T0201 WITH(NOLOCK) WHERE ChangeType = 'DAC' AND ChangeNo = " & SQLString(_changeNo)
        chkIsChangeDepAccount.Checked = L3Bool(ReturnScalar(sSQL))
    End Sub
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1022
    '# Created User: NGOCTHOAI
    '# Created Date: 21/06/2017 02:31:47
    '21/6/2017, Nguyễn Thị Hồng Nhị: id 97491-Bổ sung Thiết lập ưu tiên lấy diễn giải
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1022() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon tab 2 Tai chinh " & vbCrLf)
        sSQL &= "Exec D02P1022 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString(_changeNo) 'ChangeNo, varchar[50], NOT NULL
        Return sSQL
    End Function


    Private Sub LoadTDBGrid1(ByVal sChangeNo As String)
        Dim sSQL As New StringBuilder(252)
        sSQL.Append(" Select A.AssignmentID, B.AssignmentName" & UnicodeJoin(gbUnicode) & " As AssignmentName, A.PercentAmount ")
        sSQL.Append(" From D02T0201 A WITH(NOLOCK) Left Join D02T0002 B WITH(NOLOCK) On A.AssignmentID = B.AssignmentID ")
        sSQL.Append(" Where ChangeType = 'AS' And ChangeNo = " & SQLString(sChangeNo))

        LoadDataSource(tdbg1, sSQL.ToString, gbUnicode)
        chkDistributeChange.Checked = tdbg1.RowCount > 0
        tdbg1.Enabled = chkDistributeChange.Checked

    End Sub

    Private Sub LoadTDBGrid2(ByVal sChangeNo As String)
        Dim sSQL As New StringBuilder(449)
        sSQL.Append(" Select  A. ChangeNo, A.VoucherTypeID, A.VoucherDesc" & UnicodeJoin(gbUnicode) & " As VoucherDesc, A.RefNo, A.Serial, A.TransDesc" & UnicodeJoin(gbUnicode) & " As TransDesc, ")
        sSQL.Append(" A.ObjectTypeID, A.ObjectID, A.CurrencyID, A.ExchangeRate, A.DebitAccountID, A.CreditAccountID, ")
        sSQL.Append(" A.Amount, A.SourceID , B.CipNo, A.CipID ")
        sSQL.Append(" From D02T0204 A WITH(NOLOCK) Left Join D02T0100 B WITH(NOLOCK) On A.CipID = B.CipID  ")
        sSQL.Append(" Where ChangeNo = " & SQLString(sChangeNo))
        sSQL.Append(" Order By RefNo")
        dtGrid2 = ReturnDataTable(sSQL.ToString)
        LoadDataSource(tdbg2, dtGrid2, gbUnicode)
    End Sub

    Private Sub LoadTDBGrid3(ByVal sChangeNo As String, ByVal iMode As Integer)
        Dim sSQL As String = SQLStoreD02P0205(sChangeNo, iMode)
        Dim dtTemp As DataTable = ReturnDataTable(sSQL.ToString)
        LoadDataSource(tdbgInfo, dtTemp, gbUnicode)
    End Sub

    Private Function AllowSave() As Boolean
        If txtChangeNo.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("Ma_nghiep_vu"))
            txtChangeNo.Focus()
            Return False
        End If

        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        If chkIsChangeAssetAccount.Checked Or chkIsChangeDepAccount.Checked Then
            If chkUseAccount.Checked = False Then
                D99C0008.MsgL3(rL3("Nghiep_vu_thay_doi_tai_khoan_phai_chon_co_kem_tac_dong_tai_chinh")) 'rL3("Nghiep_vu_thay_doi_tai_san_phai_chon_co_kem_tac_dong_tai_chinh")
                chkUseAccount.Focus()
                Return False
            End If
            If chkIsEliminated.Checked = True Then
                D99C0008.MsgL3(rL3("Ban_khong_duoc_phep_thuc_hien_dong_thoi_nghiep_vu_thanh_ly_va_thay_doi_tai_khoan"))
                tabMain.SelectedTab = TabPage1
                chkIsEliminated.Focus()
                Return False
            End If
        End If

        If chkDepreciationTime.Checked Then
            If txtServiceLife.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rL3("So_ky_khau_hao"))
                tabMain.SelectedTab = TabPage1
                txtServiceLife.Focus()
                Return False
            End If
        End If

        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        If chkUseAccount.Checked And (chkIsChangeAssetAccount.Checked = False And chkIsChangeDepAccount.Checked = False) Then
            If tdbcVoucherTypeID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rL3("Loai_phieu"))
                tabMain.SelectedTab = TabPage2
                tdbcVoucherTypeID.Focus()
                Return False
            End If

            If tdbcCurrencyID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rL3("Loai_tien"))
                tabMain.SelectedTab = TabPage2
                tdbcCurrencyID.Focus()
                Return False
            End If

            If tdbg2.RowCount <= 0 Then
                D99C0008.MsgNoDataInGrid()
                tabMain.SelectedTab = TabPage2
                tdbg2.Focus()
                Return False
            End If

            If tdbcDescPriorityID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblDescPriorityID.Text)
                tabMain.SelectedTab = TabPage2
                tdbcDescPriorityID.Focus()
                Return False
            End If
        End If

        If _FormState = EnumFormState.FormAdd Then
            If IsExistKey("D02T0201", "ChangeNo", txtChangeNo.Text) Then
                D99C0008.MsgDuplicatePKey()
                txtChangeNo.Focus()
                Return False
            End If
        End If
        If chkDepreciationChange.Checked = False And chkUsingChange.Checked = False And chkDepreciationTime.Checked = False And chkDistributeChange.Checked = False And chkReceiveChange.Checked = False And chkManagementChange.Checked = False And chkUseAccount.Checked = False Then
            D99C0008.MsgL3(rL3("Ban_phai_chon_it_nhat_1_nghiep_vu"))
            chkDepreciationChange.Focus()
            Return False
        End If
        If (chkDistributeChange.Checked) Then
            If tdbg1.RowCount <= 0 Then
                D99C0008.MsgNotYetEnter(rL3("Ma_tieu_thuc_phan_bo"))
                tabMain.SelectedTab = TabPage1
                tdbg1.Focus()
                Return False
            End If
            For i As Integer = 0 To tdbg1.RowCount - 1
                If tdbg1(i, COL1_AssignmentID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rL3("Ma_tieu_thuc_phan_bo"))
                    tabMain.SelectedTab = TabPage1
                    tdbg1.Focus()
                    tdbg1.SplitIndex = SPLIT0
                    tdbg1.Col = COL1_AssignmentID
                    tdbg1.Bookmark = i
                    Return False
                End If
            Next
            For i As Integer = 0 To tdbg1.RowCount - 1
                If L3Int(tdbg1(i, COL1_PercentAmount).ToString) = 0 Or tdbg1(i, COL1_PercentAmount).ToString = "" Then
                    D99C0008.MsgL3(rL3("Ty_le_phai_lon_hon_0"))
                    tabMain.SelectedTab = TabPage1
                    tdbg1.Focus()
                    tdbg1.SplitIndex = SPLIT0
                    tdbg1.Col = COL1_PercentAmount
                    tdbg1.Bookmark = i
                    Return False
                End If
            Next
        End If

        tdbgInfo.UpdateData()
        If tdbgInfo.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tabMain.SelectedTab = TabPage3
            tdbgInfo.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbgInfo.RowCount - 1
            If L3Bool(tdbgInfo(i, COL3_IsUse).ToString) And tdbgInfo(i, COL3_Caption84).ToString = "" Then
                D99C0008.MsgNotYetEnter(tdbgInfo.Columns(COL3_Caption84).Caption)
                tabMain.SelectedTab = TabPage3
                tdbgInfo.Focus()
                tdbgInfo.SplitIndex = 0
                tdbgInfo.Col = COL3_Caption84
                tdbgInfo.Bookmark = i
                Return False
            End If
            If L3Bool(tdbgInfo(i, COL3_IsUse).ToString) And tdbgInfo(i, COL3_Caption01).ToString = "" Then
                D99C0008.MsgNotYetEnter(tdbgInfo.Columns(COL3_Caption01).Caption)
                tdbgInfo.Focus()
                tabMain.SelectedTab = TabPage3
                tdbgInfo.SplitIndex = 0
                tdbgInfo.Col = COL3_Caption01
                tdbgInfo.Bookmark = i
                Return False
            End If
        Next

        Return True
    End Function


    Private Function SQLInsertD02T0201() As String
        Dim sSQL As New StringBuilder(4000)
        Dim sIGEDetailID As String
        Dim sCreateDateAdd As Date = Now.Date
        'Check thay đổi tình trạng khấu hao
        If chkDepreciationChange.Checked Then
            sIGEDetailID = CreateIGE("D02T0201", "DetailID", "02", "TA", gsStringKey)
            sSQL.Append(" Insert Into D02T0201(")
            sSQL.Append(" DetailID, ChangeNo, ChangeNameU, ChangeType, Disabled, ")
            sSQL.Append(" Notes1U, Notes2U, Notes3U, CreateUserID, CreateDate, ")
            sSQL.Append(" LastModifyDate, LastModifyUserID, SourceID, ConvertedAmount, PercentAmount, ")
            sSQL.Append(" AssignmentID, ObjectTypeID, ObjectID, EmployeeID, FullName, ")
            sSQL.Append(" StopDepreciation, StopUse, ServiceLife, IsEliminated, UseAccount, AssetWHID")
            sSQL.Append(" ) Values ( ")
            sSQL.Append(SQLString(sIGEDetailID) & COMMA) 'DetailID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL")
            sSQL.Append(SQLStringUnicode(txtChangeName, True) & COMMA) 'ChangeNameU, varchar[50], NULL")
            sSQL.Append(SQLString("SD") & COMMA) 'ChangeType, varchar[20], NULL")
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL")
            sSQL.Append(SQLStringUnicode(txtNote1, True) & COMMA) 'Notes1U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote2, True) & COMMA) 'Notes2U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote3, True) & COMMA) 'Notes3U, varchar[250], NULL")
            If _FormState = EnumFormState.FormAdd Then
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDateAdd) & COMMA) 'CreateDate, datetime, NULL")
            Else
                sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL")
            End If
            sSQL.Append(" GetDate()" & COMMA) 'LastModifyDate, datetime, NULL")
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'SourceID, varchar[20], NULL")
            sSQL.Append(SQLMoney("0", DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'PercentAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'AssignmentID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectTypeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'EmployeeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'FullName, varchar[50], NULL")
            sSQL.Append(SQLNumber(optStopDepreciation.Checked) & COMMA) 'StopDepreciation, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'StopUse, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL")
            sSQL.Append(SQLNumber(chkIsEliminated.Checked) & COMMA) 'IsEliminated, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(chkUseAccount.Checked) & COMMA) 'UseAccount, tinyint, NOT NULL")
            sSQL.Append(SQLString(""))


            sSQL.Append(" ) " & vbCrLf)
        End If

        'Check thay đổi tình trạng sử dụng
        If chkUsingChange.Checked Then
            sIGEDetailID = CreateIGE("D02T0201", "DetailID", "02", "TA", gsStringKey)
            sSQL.Append(" Insert Into D02T0201(")
            sSQL.Append(" DetailID, ChangeNo, ChangeNameU, ChangeType, Disabled, ")
            sSQL.Append(" Notes1U, Notes2U, Notes3U, CreateUserID, CreateDate, ")
            sSQL.Append(" LastModifyDate, LastModifyUserID, SourceID, ConvertedAmount, PercentAmount, ")
            sSQL.Append(" AssignmentID, ObjectTypeID, ObjectID, EmployeeID, FullName, ")
            sSQL.Append(" StopDepreciation, StopUse, ServiceLife, IsEliminated, UseAccount, AssetWHID ")
            sSQL.Append(" ) Values ( ")
            sSQL.Append(SQLString(sIGEDetailID) & COMMA) 'DetailID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL")
            sSQL.Append(SQLStringUnicode(txtChangeName, True) & COMMA) 'ChangeNameU, varchar[50], NULL")
            sSQL.Append(SQLString("SU") & COMMA) 'ChangeType, varchar[20], NULL")
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL")
            sSQL.Append(SQLStringUnicode(txtNote1, True) & COMMA) 'Notes1U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote2, True) & COMMA) 'Notes2U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote3, True) & COMMA) 'Notes3U, varchar[250], NULL")
            If _FormState = EnumFormState.FormAdd Then
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDateAdd) & COMMA) 'CreateDate, datetime, NULL")
            Else
                sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL")
            End If
            sSQL.Append(" GetDate()" & COMMA) 'LastModifyDate, datetime, NULL")
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'SourceID, varchar[20], NULL")
            sSQL.Append(SQLMoney("0", DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'PercentAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'AssignmentID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectTypeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'EmployeeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'FullName, varchar[50], NULL")
            sSQL.Append("NULL" & COMMA) 'StopDepreciation, tinyint, NULL")
            sSQL.Append(SQLNumber(optStopUse.Checked) & COMMA) 'StopUse, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL")
            sSQL.Append(SQLNumber(chkIsEliminated.Checked) & COMMA) 'IsEliminated, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(chkUseAccount.Checked) & COMMA) 'UseAccount, tinyint, NOT NULL")
            sSQL.Append(SQLString(tdbcAssetWHID.Text))


            sSQL.Append(" ) " & vbCrLf)
        End If

        'Check thay đổi thời gian khấu hao
        If chkDepreciationTime.Checked Then
            sIGEDetailID = CreateIGE("D02T0201", "DetailID", "02", "TA", gsStringKey)
            sSQL.Append(" Insert Into D02T0201(")
            sSQL.Append(" DetailID, ChangeNo, ChangeNameU, ChangeType, Disabled, ")
            sSQL.Append(" Notes1U, Notes2U, Notes3U, CreateUserID, CreateDate, ")
            sSQL.Append(" LastModifyDate, LastModifyUserID, SourceID, ConvertedAmount, PercentAmount, ")
            sSQL.Append(" AssignmentID, ObjectTypeID, ObjectID, EmployeeID, FullName, ")
            sSQL.Append(" StopDepreciation, StopUse, ServiceLife, IsEliminated, UseAccount, AssetWHID")
            sSQL.Append(" ) Values ( ")
            sSQL.Append(SQLString(sIGEDetailID) & COMMA) 'DetailID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL")
            sSQL.Append(SQLStringUnicode(txtChangeName, True) & COMMA) 'ChangeNameU, varchar[50], NULL")
            sSQL.Append(SQLString("SL") & COMMA) 'ChangeType, varchar[20], NULL")
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL")
            sSQL.Append(SQLStringUnicode(txtNote1, True) & COMMA) 'Notes1U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote2, True) & COMMA) 'Notes2U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote3, True) & COMMA) 'Notes3U, varchar[250], NULL")
            If _FormState = EnumFormState.FormAdd Then
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDateAdd) & COMMA) 'CreateDate, datetime, NULL")
            Else
                sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL")
            End If
            sSQL.Append(" GetDate()" & COMMA) 'LastModifyDate, datetime, NULL")
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'SourceID, varchar[20], NULL")
            sSQL.Append(SQLMoney("0", DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'PercentAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'AssignmentID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectTypeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'EmployeeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'FullName, varchar[50], NULL")
            sSQL.Append("NULL" & COMMA) 'StopDepreciation, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'StopUse, tinyint, NULL")
            sSQL.Append(SQLNumber(txtServiceLife.Text) & COMMA) 'ServiceLife, int, NULL")
            sSQL.Append(SQLNumber(chkIsEliminated.Checked) & COMMA) 'IsEliminated, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(chkUseAccount.Checked) & COMMA) 'UseAccount, tinyint, NOT NULL")
            sSQL.Append(SQLString(""))

            sSQL.Append(" ) " & vbCrLf)
        End If

        'Check thay đổi tiêu thức phân bổ
        If chkDistributeChange.Checked Then
            tdbg1.UpdateData()
            For i As Integer = 0 To tdbg1.RowCount - 1
                sIGEDetailID = CreateIGE("D02T0201", "DetailID", "02", "TA", gsStringKey)
                sSQL.Append(" Insert Into D02T0201(")
                sSQL.Append(" DetailID, ChangeNo, ChangeNameU, ChangeType, Disabled, ")
                sSQL.Append(" Notes1U, Notes2U, Notes3U, CreateUserID, CreateDate, ")
                sSQL.Append(" LastModifyDate, LastModifyUserID, SourceID, ConvertedAmount, PercentAmount, ")
                sSQL.Append(" AssignmentID, ObjectTypeID, ObjectID, EmployeeID, FullName, ")
                sSQL.Append(" StopDepreciation, StopUse, ServiceLife, IsEliminated, UseAccount, AssetWHID")
                sSQL.Append(" ) Values ( ")
                sSQL.Append(SQLString(sIGEDetailID) & COMMA) 'DetailID [KEY], varchar[20], NOT NULL
                sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL")
                sSQL.Append(SQLStringUnicode(txtChangeName, True) & COMMA) 'ChangeNameU, varchar[50], NULL")
                sSQL.Append(SQLString("AS") & COMMA) 'ChangeType, varchar[20], NULL")
                sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL")
                sSQL.Append(SQLStringUnicode(txtNote1, True) & COMMA) 'Notes1U, varchar[250], NULL")
                sSQL.Append(SQLStringUnicode(txtNote2, True) & COMMA) 'Notes2U, varchar[250], NULL")
                sSQL.Append(SQLStringUnicode(txtNote3, True) & COMMA) 'Notes3U, varchar[250], NULL")
                If _FormState = EnumFormState.FormAdd Then
                    sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                    sSQL.Append(SQLDateSave(sCreateDateAdd) & COMMA) 'CreateDate, datetime, NULL")
                Else
                    sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                    sSQL.Append(SQLDateSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL")
                End If
                sSQL.Append(" GetDate()" & COMMA) 'LastModifyDate, datetime, NULL")
                sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL")
                sSQL.Append("NULL" & COMMA) 'SourceID, varchar[20], NULL")
                sSQL.Append(SQLMoney("0", DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL")
                sSQL.Append(SQLMoney(tdbg1(i, COL1_PercentAmount).ToString, DxxFormat.DefaultNumber2) & COMMA) 'PercentAmount, money, NULL")
                sSQL.Append(SQLString(tdbg1(i, COL1_AssignmentID).ToString) & COMMA) 'AssignmentID, varchar[20], NULL")
                sSQL.Append("NULL" & COMMA) 'ObjectTypeID, varchar[20], NULL")
                sSQL.Append("NULL" & COMMA) 'ObjectID, varchar[20], NULL")
                sSQL.Append("NULL" & COMMA) 'EmployeeID, varchar[20], NULL")
                sSQL.Append("NULL" & COMMA) 'FullName, varchar[50], NULL")
                sSQL.Append("NULL" & COMMA) 'StopDepreciation, tinyint, NULL")
                sSQL.Append("NULL" & COMMA) 'StopUse, tinyint, NULL")
                sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL")
                sSQL.Append(SQLNumber(chkIsEliminated.Checked) & COMMA) 'IsEliminated, tinyint, NOT NULL")
                sSQL.Append(SQLNumber(chkUseAccount.Checked) & COMMA) 'UseAccount, tinyint, NOT NULL")
                sSQL.Append(SQLString(""))

                sSQL.Append(" ) " & vbCrLf)
            Next
        End If

        'Check thay đổi bộ phận quản lí
        If chkReceiveChange.Checked Or chkManagementChange.Checked Then 'ID : 214915
            sIGEDetailID = CreateIGE("D02T0201", "DetailID", "02", "TA", gsStringKey)
            sSQL.Append(" Insert Into D02T0201(")
            sSQL.Append(" DetailID, ChangeNo, ChangeNameU, ChangeType, Disabled, ")
            sSQL.Append(" Notes1U, Notes2U, Notes3U, CreateUserID, CreateDate, ")
            sSQL.Append(" LastModifyDate, LastModifyUserID, SourceID, ConvertedAmount, PercentAmount, ")
            sSQL.Append(" AssignmentID, ObjectTypeID, ObjectID, EmployeeID, FullNameU, ")
            sSQL.Append(" StopDepreciation, StopUse, ServiceLife, IsEliminated, UseAccount, AssetWHID,LocationID, ManagementObjTypeID, ManagementObjID, IsManagement , isReceive ") 'ID : 214915 - IsManagemnet, IsReceive lưu giá trị checkbox
            sSQL.Append(" ) Values ( ")
            sSQL.Append(SQLString(sIGEDetailID) & COMMA) 'DetailID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL")
            sSQL.Append(SQLStringUnicode(txtChangeName, True) & COMMA) 'ChangeNameU, varchar[50], NULL")
            sSQL.Append(SQLString("OB") & COMMA) 'ChangeType, varchar[20], NULL")
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL")
            sSQL.Append(SQLStringUnicode(txtNote1, True) & COMMA) 'Notes1U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote2, True) & COMMA) 'Notes2U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote3, True) & COMMA) 'Notes3U, varchar[250], NULL")
            If _FormState = EnumFormState.FormAdd Then
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDateAdd) & COMMA) 'CreateDate, datetime, NULL")
            Else
                sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL")
            End If
            sSQL.Append(" GetDate()" & COMMA) 'LastModifyDate, datetime, NULL")
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'SourceID, varchar[20], NULL")
            sSQL.Append(SQLMoney("0", DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'PercentAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'AssignmentID, varchar[20], NULL")
            sSQL.Append(SQLString(tdbcObjectTypeID.Text) & COMMA) 'ObjectTypeID, varchar[20], NULL")
            sSQL.Append(SQLString(tdbcObjectID.Text) & COMMA) 'ObjectID, varchar[20], NULL")
            sSQL.Append(SQLString(tdbcEmployeeID.Text) & COMMA) 'EmployeeID, varchar[20], NULL")
            sSQL.Append(SQLStringUnicode(txtEmployeeName, True) & COMMA) 'FullNameU, varchar[50], NULL")
            sSQL.Append("NULL" & COMMA) 'StopDepreciation, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'StopUse, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL")
            sSQL.Append(SQLNumber(chkIsEliminated.Checked) & COMMA) 'IsEliminated, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(chkUseAccount.Checked) & COMMA) 'UseAccount, tinyint, NOT NULL")
            sSQL.Append(SQLString("") & COMMA)
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcLocationID)) & COMMA)  'LocationID, varchar[50], NOT NULL
            '30/10/2019, Lê Thị Phú Hà:id 123376-Bổ sung thay đổi bộ phận quản lý (NV tác động D02)
            sSQL.Append(SQLString(tdbcManagementObjTypeID.Text) & COMMA) 'ManagementObjTypeID, varchar[50], NULL")
            sSQL.Append(SQLString(tdbcManagementObjID.Text) & COMMA) 'ManagementObjID, varchar[50], NULL")
            sSQL.Append(SQLNumber(chkManagementChange.Checked) & COMMA) 'IsManagement, tinyint, NOT NULL") 'ID : 214915
            sSQL.Append(SQLNumber(chkReceiveChange.Checked)) 'IsReceive, tinyint, NOT NULL") 'ID : 214915

            sSQL.Append(" ) " & vbCrLf)
        End If
        'Check đính kèm tác động tài chính
        If chkDepreciationChange.Checked = False And chkUsingChange.Checked = False And chkDepreciationTime.Checked = False And chkDistributeChange.Checked = False And chkReceiveChange.Checked = False And chkUseAccount.Checked = True Then
            sIGEDetailID = CreateIGE("D02T0201", "DetailID", "02", "TA", gsStringKey)
            sSQL.Append(" Insert Into D02T0201(")
            sSQL.Append(" DetailID, ChangeNo, ChangeNameU, ChangeType, Disabled, ")
            sSQL.Append(" Notes1U, Notes2U, Notes3U, CreateUserID, CreateDate, ")
            sSQL.Append(" LastModifyDate, LastModifyUserID, SourceID, ConvertedAmount, PercentAmount, ")
            sSQL.Append(" AssignmentID, ObjectTypeID, ObjectID, EmployeeID, FullName, ")
            sSQL.Append(" StopDepreciation, StopUse, ServiceLife, IsEliminated, UseAccount, AssetWHID")
            sSQL.Append(" ) Values ( ")
            sSQL.Append(SQLString(sIGEDetailID) & COMMA) 'DetailID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL")
            sSQL.Append(SQLStringUnicode(txtChangeName, True) & COMMA) 'ChangeNameU, varchar[50], NULL")
            sSQL.Append("NULL" & COMMA) 'ChangeType, varchar[20], NULL")
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL")
            sSQL.Append(SQLStringUnicode(txtNote1, True) & COMMA) 'Notes1U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote2, True) & COMMA) 'Notes2U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote3, True) & COMMA) 'Notes3U, varchar[250], NULL")
            If _FormState = EnumFormState.FormAdd Then
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDateAdd) & COMMA) 'CreateDate, datetime, NULL")
            Else
                sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL")
            End If
            sSQL.Append(" GetDate()" & COMMA) 'LastModifyDate, datetime, NULL")
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'SourceID, varchar[20], NULL")
            sSQL.Append(SQLMoney("0", DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'PercentAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'AssignmentID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectTypeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'EmployeeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'FullName, varchar[50], NULL")
            sSQL.Append("NULL" & COMMA) 'StopDepreciation, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'StopUse, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL")
            sSQL.Append(SQLNumber(chkIsEliminated.Checked) & COMMA) 'IsEliminated, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(chkUseAccount.Checked) & COMMA) 'UseAccount, tinyint, NOT NULL")
            sSQL.Append(SQLString(""))

            sSQL.Append(" ) " & vbCrLf)
        End If

        If chkIsChangeAssetAccount.Checked Then 'Thay đổi tài khoản tài sản
            sIGEDetailID = CreateIGE("D02T0201", "DetailID", "02", "TA", gsStringKey)
            sSQL.Append(" Insert Into D02T0201(")
            sSQL.Append(" DetailID, ChangeNo, ChangeNameU, ChangeType, Disabled, ")
            sSQL.Append(" Notes1U, Notes2U, Notes3U, CreateUserID, CreateDate, ")
            sSQL.Append(" LastModifyDate, LastModifyUserID, SourceID, ConvertedAmount, PercentAmount, ")
            sSQL.Append(" AssignmentID, ObjectTypeID, ObjectID, EmployeeID, FullName, ")
            sSQL.Append(" StopDepreciation, StopUse, ServiceLife, IsEliminated, UseAccount, AssetWHID,IsChangeAssetAccount,IsChangeDepAccount")
            sSQL.Append(" ) Values ( ")
            sSQL.Append(SQLString(sIGEDetailID) & COMMA) 'DetailID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL")
            sSQL.Append(SQLStringUnicode(txtChangeName, True) & COMMA) 'ChangeNameU, varchar[50], NULL")
            sSQL.Append(SQLString("AAC") & COMMA) 'ChangeType, varchar[20], NULL")
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL")
            sSQL.Append(SQLStringUnicode(txtNote1, True) & COMMA) 'Notes1U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote2, True) & COMMA) 'Notes2U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote3, True) & COMMA) 'Notes3U, varchar[250], NULL")
            If _FormState = EnumFormState.FormAdd Then
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDateAdd) & COMMA) 'CreateDate, datetime, NULL")
            Else
                sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL")
            End If
            sSQL.Append(" GetDate()" & COMMA) 'LastModifyDate, datetime, NULL")
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'SourceID, varchar[20], NULL")
            sSQL.Append(SQLMoney("0", DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'PercentAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'AssignmentID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectTypeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'EmployeeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'FullName, varchar[50], NULL")
            sSQL.Append(SQLNumber(0) & COMMA) 'StopDepreciation, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'StopUse, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL")
            sSQL.Append(SQLNumber(chkIsEliminated.Checked) & COMMA) 'IsEliminated, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(chkUseAccount.Checked) & COMMA) 'UseAccount, tinyint, NOT NULL")
            sSQL.Append(SQLString("") & COMMA)
            sSQL.Append(SQLNumber(1) & COMMA) 'IsChangeAssetAccount, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(0)) 'IsChangeDepAccount, tinyint, NOT NULL")


            sSQL.Append(" ) " & vbCrLf)
        End If

        If chkIsChangeDepAccount.Checked Then 'Thay đổi tài khoản khấu hào
            sIGEDetailID = CreateIGE("D02T0201", "DetailID", "02", "TA", gsStringKey)
            sSQL.Append(" Insert Into D02T0201(")
            sSQL.Append(" DetailID, ChangeNo, ChangeNameU, ChangeType, Disabled, ")
            sSQL.Append(" Notes1U, Notes2U, Notes3U, CreateUserID, CreateDate, ")
            sSQL.Append(" LastModifyDate, LastModifyUserID, SourceID, ConvertedAmount, PercentAmount, ")
            sSQL.Append(" AssignmentID, ObjectTypeID, ObjectID, EmployeeID, FullName, ")
            sSQL.Append(" StopDepreciation, StopUse, ServiceLife, IsEliminated, UseAccount, AssetWHID,IsChangeAssetAccount,IsChangeDepAccount")
            sSQL.Append(" ) Values ( ")
            sSQL.Append(SQLString(sIGEDetailID) & COMMA) 'DetailID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL")
            sSQL.Append(SQLStringUnicode(txtChangeName, True) & COMMA) 'ChangeNameU, varchar[50], NULL")
            sSQL.Append(SQLString("DAC") & COMMA) 'ChangeType, varchar[20], NULL")
            sSQL.Append(SQLNumber(chkDisabled.Checked) & COMMA) 'Disabled, tinyint, NOT NULL")
            sSQL.Append(SQLStringUnicode(txtNote1, True) & COMMA) 'Notes1U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote2, True) & COMMA) 'Notes2U, varchar[250], NULL")
            sSQL.Append(SQLStringUnicode(txtNote3, True) & COMMA) 'Notes3U, varchar[250], NULL")
            If _FormState = EnumFormState.FormAdd Then
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDateAdd) & COMMA) 'CreateDate, datetime, NULL")
            Else
                sSQL.Append(SQLString(sCreateUserID) & COMMA) 'CreateUserID, varchar[20], NULL")
                sSQL.Append(SQLDateSave(sCreateDate) & COMMA) 'CreateDate, datetime, NULL")
            End If
            sSQL.Append(" GetDate()" & COMMA) 'LastModifyDate, datetime, NULL")
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'SourceID, varchar[20], NULL")
            sSQL.Append(SQLMoney("0", DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'PercentAmount, money, NULL")
            sSQL.Append("NULL" & COMMA) 'AssignmentID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectTypeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'ObjectID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'EmployeeID, varchar[20], NULL")
            sSQL.Append("NULL" & COMMA) 'FullName, varchar[50], NULL")
            sSQL.Append(SQLNumber(0) & COMMA) 'StopDepreciation, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'StopUse, tinyint, NULL")
            sSQL.Append("NULL" & COMMA) 'ServiceLife, int, NULL")
            sSQL.Append(SQLNumber(chkIsEliminated.Checked) & COMMA) 'IsEliminated, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(chkUseAccount.Checked) & COMMA) 'UseAccount, tinyint, NOT NULL")
            sSQL.Append(SQLString("") & COMMA)
            sSQL.Append(SQLNumber(0) & COMMA) 'IsChangeAssetAccount, tinyint, NOT NULL")
            sSQL.Append(SQLNumber(1)) 'IsChangeDepAccount, tinyint, NOT NULL")
            sSQL.Append(" ) " & vbCrLf)
        End If

        Return sSQL.ToString
    End Function

    Private Function SQLInsertD02T0204() As String
        Dim sSQL As New StringBuilder
        If Not chkUseAccount.Checked Then Return ""
        tdbg2.UpdateData()
        For i As Integer = 0 To tdbg2.RowCount - 1
            sSQL.Append(" Insert Into D02T0204( ")
            sSQL.Append(" ChangeNo, VoucherTypeID, VoucherDescU, RefNo, Serial, ")
            sSQL.Append(" TransDescU, ObjectTypeID, ObjectID, CurrencyID, ExchangeRate, ")
            sSQL.Append(" DebitAccountID, CreditAccountID, Amount, SourceID, CipID, DescPriorityID")
            sSQL.Append(" ) Values( ")
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcVoucherTypeID.Text) & COMMA) 'VoucherTypeID, varchar[20], NOT NULL
            sSQL.Append(SQLStringUnicode(txtDescription, True) & COMMA) 'VoucherDescU, varchar[250], NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_RefNo).ToString) & COMMA) 'RefNo, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Serial).ToString) & COMMA) 'Serial, varchar[20], NOT NULL
            sSQL.Append(SQLStringUnicode(tdbg2(i, COL2_TransDesc).ToString, gbUnicode, True) & COMMA) 'TransDescU, varchar[250], NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_ObjectTypeID).ToString) & COMMA) 'ObjectTypeID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_ObjectID).ToString) & COMMA) 'ObjectID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcCurrencyID.Text) & COMMA) 'CurrencyID, varchar[20], NOT NULL
            sSQL.Append(SQLMoney(tdbg2(i, COL2_ExchangeRate).ToString, DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_DebitAccountID).ToString) & COMMA) 'DebitAccountID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_CreditAccountID).ToString) & COMMA) 'CreditAccountID, varchar[20], NOT NULL
            sSQL.Append(SQLMoney(tdbg2(i, COL2_Amount).ToString, DxxFormat.DecimalPlaces) & COMMA) 'Amount, money, NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_SourceID).ToString) & COMMA) 'SourceID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_CipID).ToString) & COMMA) 'CipID, varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcDescPriorityID.Text)) 'DescPriorityID
            sSQL.Append(" ) ")
        Next
        Return sSQL.ToString
    End Function

    Private Function SQLDelete() As String
        Dim sSQL As New StringBuilder(100)
        sSQL.Append(" Delete D02T0201")
        sSQL.Append(" Where ChangeNo = " & SQLString(txtChangeNo.Text) & vbCrLf)

        sSQL.Append(" Delete D02T0204")
        sSQL.Append(" Where ChangeNo = " & SQLString(txtChangeNo.Text) & vbCrLf)

        Return sSQL.ToString
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0201
    '# Created User: Nguyễn Đức Trọng
    '# Created Date: 29/09/2011 11:25:45
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0201() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0201 Set ")
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'tinyint, NOT NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("ChangeNameU = " & SQLStringUnicode(txtChangeName, True)) 'nvarchar, NOT NULL
        sSQL.Append(" Where ChangeNo = " & SQLString(txtChangeNo.Text) & vbCrLf)

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0204
    '# Created User: Phạm Vanw Vinh
    '# Created Date: 19/07/2012 08:37:45
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0204() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0204 Set ") 'tdbcVoucherTypeID
        sSQL.Append("VoucherDescU  = " & SQLStringUnicode(txtDescription, True)) 'varchar (250), NOT NULL
        sSQL.Append(" Where ChangeNo = " & SQLString(txtChangeNo.Text) & vbCrLf)

        Return sSQL
    End Function

    Dim sCreateUserID As String = ""
    Dim sCreateDate As String = ""
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        tdbgInfo.UpdateData()
        If _FormState <> EnumFormState.FormEditOther Then
            If Not AllowSave() Then Exit Sub
        End If

        btnSave.Enabled = False
        btnClose.Enabled = False

        If _FormState = EnumFormState.FormAdd Then
            sCreateUserID = gsUserID
            sCreateDate = Now.ToString
        Else
            sCreateUserID = _createUserID
            sCreateDate = _createDate
        End If

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                sSQL.Append(SQLInsertD02T0201() & vbCrLf)
                sSQL.Append(SQLInsertD02T0204() & vbCrLf)
                sSQL.Append(SQLInsertD02T0205s().ToString & vbCrLf)
            Case EnumFormState.FormEdit
                sSQL.Append(SQLDelete)
                sSQL.Append(SQLInsertD02T0201() & vbCrLf)
                sSQL.Append(SQLInsertD02T0204() & vbCrLf)
                sSQL.Append(SQLDeleteD02T0205().ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T0205s().ToString & vbCrLf)
            Case EnumFormState.FormEditOther
                sSQL.Append(SQLUpdateD02T0201)
                'Cập nhật ngày 19/7/2012 theo incident  49798 của HIENDUY
                sSQL.Append(SQLUpdateD02T0204)
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            btnClose.Enabled = True
            _savedOK = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnNext.Enabled = True
                    btnNext.Focus()
                    ChangeNo = txtChangeNo.Text
                Case EnumFormState.FormEdit
                    btnSave.Enabled = True
                    btnClose.Focus()
                Case EnumFormState.FormEditOther
                    btnSave.Enabled = True
                    btnClose.Focus()
            End Select
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        chkUseAccount.Checked = False
        'Tab phi tài chính
        txtChangeNo.Text = ""
        txtChangeName.Text = ""
        txtServiceLife.Text = ""
        tdbcObjectTypeID.Text = ""
        tdbcObjectID.Text = ""
        txtObjectName.Text = ""
        tdbcManagementObjTypeID.Text = ""
        tdbcManagementObjID.Text = ""
        tdbcEmployeeID.Text = ""
        txtEmployeeName.Text = ""
        tdbcLocationID.Text = ""
        tdbcLocationID.Enabled = False
        txtLocationName.Text = ""
        txtNote1.Text = ""
        txtNote2.Text = ""
        txtNote3.Text = ""
        chkDepreciationChange.Checked = False
        optStopDepreciation.Enabled = False
        optReDepreciation.Enabled = False
        chkUsingChange.Checked = False
        optStopUse.Enabled = False
        optReUse.Enabled = False
        chkDepreciationTime.Checked = False
        txtServiceLife.Enabled = False
        txtServiceLife.BackColor = Color.White
        chkDistributeChange.Checked = False
        tdbg1.Enabled = False
        chkReceiveChange.Checked = False
        tdbcObjectTypeID.Enabled = False
        tdbcObjectID.Enabled = False
        tdbcManagementObjID.Enabled = False
        tdbcManagementObjTypeID.Enabled = False

        tdbcEmployeeID.Enabled = False
        chkIsEliminated.Checked = False
        chkDisabled.Checked = False
        'Xóa dữ liệu của lưới
        LoadTDBGrid1("-1")
        LoadTDBGrid2("-1")
        LoadTDBGrid3("", 0)
        'Tab tài chính
        tdbcVoucherTypeID.Text = ""
        tdbcCurrencyID.Text = ""
        txtDescription.Text = ""
        tdbcAssetWHID.Text = ""
        'Xóa dữ liệu của lưới

        btnNext.Enabled = False
        btnSave.Enabled = True
        btnClose.Enabled = True
        tabMain.SelectedTab = TabPage1
        txtChangeNo.Focus()
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

#Region "Events tdbcEmployeeID with txtEmployeeName"

    Private Sub tdbcEmployeeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcEmployeeID.Close
        If tdbcEmployeeID.FindStringExact(tdbcEmployeeID.Text) = -1 Then
            tdbcEmployeeID.Text = ""
            txtEmployeeName.Text = ""
            tdbcEmployeeID.Focus()
        End If
    End Sub

    Private Sub tdbcEmployeeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcEmployeeID.SelectedValueChanged
        txtEmployeeName.Text = tdbcEmployeeID.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcEmployeeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcEmployeeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcEmployeeID.Text = ""
            txtEmployeeName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcCurrencyID"

    Private Sub tdbcCurrencyID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCurrencyID.Close
        If tdbcCurrencyID.FindStringExact(tdbcCurrencyID.Text) = -1 Then
            tdbcCurrencyID.Text = ""
            tdbcCurrencyID.Focus()
        End If
    End Sub

    Private Sub tdbcCurrencyID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCurrencyID.SelectedValueChanged
        SetExchangeRateToGrid()
    End Sub

    Private Sub SetExchangeRateToGrid()
        Dim i As Integer
        For i = 0 To tdbg2.RowCount - 1
            tdbg2(i, COL2_ExchangeRate) = Convert.ToDouble(SQLMoney(tdbcCurrencyID.Columns("ExchangeRate").Text, DxxFormat.ExchangeRateDecimals))
        Next
        DxxFormat.DecimalPlaces = GetOriginalDecimal(tdbcCurrencyID.Text)
    End Sub

    Private Sub tdbcCurrencyID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcCurrencyID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcCurrencyID.Text = ""
    End Sub

#End Region

#Region "Events tdbcVoucherTypeID with txtDescription load tdbcObjectTypeID"

    Private Sub tdbcVoucherTypeID_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.Close
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then
            tdbcVoucherTypeID.Text = ""
            tdbcVoucherTypeID.Focus()
        End If

    End Sub

    Private Sub tdbcVoucherTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.SelectedValueChanged
        'txtDescription.Text = tdbcVoucherTypeID.Columns(1).Text
    End Sub

    Private Sub tdbcVoucherTypeID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcVoucherTypeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcVoucherTypeID.Text = ""
        End If
    End Sub

    Private Sub tdbcObjectTypeID_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.Close
        If tdbcObjectTypeID.FindStringExact(tdbcObjectTypeID.Text) = -1 Then
            tdbcObjectTypeID.Text = ""
            tdbcObjectID.Text = ""
            txtObjectName.Text = ""
            tdbcObjectTypeID.Focus()
        End If

    End Sub

    Private Sub tdbcObjectTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcObjectTypeID.SelectedValueChanged
        txtObjectName.Text = ""
        LoadDataSource(tdbcObjectID, ReturnTableFilter(dtObjectID, " ObjectTypeID = " & SQLString(tdbcObjectTypeID.Text), True), gbUnicode)
    End Sub

    Private Sub tdbcObjectTypeID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcObjectTypeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcObjectTypeID.Text = ""
            tdbcObjectID.Text = ""
            txtObjectName.Text = ""
        End If

    End Sub

#End Region

#Region "Events tdbcObjectID with txtObjectName"

    Private Sub tdbcObjectID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcObjectID.Close
        If tdbcObjectID.FindStringExact(tdbcObjectID.Text) = -1 Then
            tdbcObjectID.Text = ""
            txtObjectName.Text = ""
            tdbcObjectID.Focus()
        End If
    End Sub

    Private Sub tdbcObjectID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcObjectID.SelectedValueChanged
        txtObjectName.Text = tdbcObjectID.Columns(1).Text
    End Sub

    Private Sub tdbcObjectID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcObjectID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcObjectID.Text = ""
            txtObjectName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcAssetWHID"

    Private Sub tdbcAssetWHID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetWHID.LostFocus
        If tdbcAssetWHID.FindStringExact(tdbcAssetWHID.Text) = -1 Then tdbcAssetWHID.Text = ""
    End Sub

#End Region

    Private Sub tdbg1_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.ComboSelect
        Select Case e.ColIndex
            Case COL1_AssignmentID
                tdbg1.Columns(COL1_AssignmentName).Text = tdbdAssignmentID.Columns(1).Text
        End Select
    End Sub

    Private Sub tdbg1_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg1.BeforeColUpdate
        Select Case e.ColIndex
            Case COL1_AssignmentID
                If tdbg1.Columns(COL1_AssignmentID).Text <> tdbdAssignmentID.Columns(0).Text Then
                    tdbg1.Columns(COL1_AssignmentID).Text = ""
                End If
            Case COL1_AssignmentName
            Case COL1_PercentAmount
                If tdbg1.Columns(COL1_PercentAmount).Text = "" Then Exit Sub
                If Not IsNumeric(tdbg1.Columns(COL1_PercentAmount).Text) Then
                    tdbg1.Columns(COL1_PercentAmount).Text = ""
                End If
        End Select
    End Sub

    Private Sub tdbg1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg1.KeyPress
        Select Case tdbg1.Col
            Case COL1_PercentAmount
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        If e.KeyCode = Keys.F7 Then
            HotKeyF7(tdbg1)
            Exit Sub
        End If

        If e.KeyCode = Keys.Enter Then
            If tdbg1.Col = iLastCol1 Then HotKeyEnterGrid(tdbg1, COL1_AssignmentID, e)
            Exit Sub
        End If

        HotKeyDownGrid(e, tdbg1, COL1_AssignmentID)

    End Sub

    Private Sub tdbg2_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.ComboSelect
        Select Case e.ColIndex
            Case COL2_ObjectTypeID
            Case COL2_ObjectID
            Case COL2_DebitAccountID
            Case COL2_CreditAccountID
            Case COL2_Amount
            Case COL2_SourceID
            Case COL2_CipNo
                tdbg2.Columns(COL2_CipID).Text = tdbdCipID.Columns(2).Text
        End Select
    End Sub

    Private Sub tdbg2_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg2.RowColChange
        If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        Select Case tdbg2.Col
            Case COL2_ObjectID
                LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObjectID, " ObjectTypeID=" & SQLString(tdbg2.Columns(COL2_ObjectTypeID).Text), True), gbUnicode)
        End Select
    End Sub

    'Private Sub tdbg2_BeforeColEdit(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColEditEventArgs) Handles tdbg2.BeforeColEdit
    '    Select Case e.ColIndex
    '        Case COL2_ObjectTypeID
    '        Case COL2_ObjectID
    '            LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObjectID, " ObjectTypeID=" & SQLString(tdbg2.Columns(COL2_ObjectTypeID).Text), True), gbUnicode)
    '        Case COL2_DebitAccountID
    '        Case COL2_CreditAccountID
    '        Case COL2_Amount
    '        Case COL2_SourceID
    '        Case COL2_CipNo
    '    End Select
    'End Sub

    Private Sub tdbg2_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg2.BeforeColUpdate
        Select Case e.ColIndex
            Case COL2_ChangeNo
            Case COL2_VoucherTypeID
            Case COL2_VoucherDesc
            Case COL2_RefNo
            Case COL2_Serial
            Case COL2_TransDesc
            Case COL2_ObjectTypeID
                If tdbg2.Columns(COL2_ObjectTypeID).Text <> tdbdObjectTypeID.Columns(0).Text Then
                    tdbg2.Columns(COL2_ObjectTypeID).Text = ""
                    tdbg2.Columns(COL2_ObjectID).Text = ""
                End If
            Case COL2_ObjectID
                If tdbg2.Columns(COL2_ObjectID).Text <> tdbdObjectID.Columns(0).Text Then
                    tdbg2.Columns(COL2_ObjectID).Text = ""
                End If
            Case COL2_CurrencyID
            Case COL2_ExchangeRate
                If tdbg2.Columns(COL2_ExchangeRate).Text = "" Then Exit Sub
                If Not IsNumeric(tdbg2.Columns(COL2_ExchangeRate).Text) Then
                    tdbg2.Columns(COL2_ExchangeRate).Text = ""
                End If
            Case COL2_DebitAccountID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg2.Columns(COL2_DebitAccountID).Text <> tdbdDebitAccountID.Columns(0).Text Then
                    tdbg2.Columns(COL2_DebitAccountID).Text = ""
                End If
            Case COL2_CreditAccountID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg2.Columns(COL2_CreditAccountID).Text <> tdbdCreditAccountID.Columns(0).Text Then
                    tdbg2.Columns(COL2_CreditAccountID).Text = ""
                End If
            Case COL2_Amount
                If tdbg2.Columns(COL2_Amount).Text = "" Then Exit Sub
                If Not IsNumeric(tdbg2.Columns(COL2_Amount).Text) Then
                    tdbg2.Columns(COL2_Amount).Text = ""
                End If
            Case COL2_SourceID
                If tdbg2.Columns(COL2_SourceID).Text <> tdbdSourceID.Columns(0).Text Then
                    tdbg2.Columns(COL2_SourceID).Text = ""
                End If
            Case COL2_CipNo
                If tdbg2.Columns(COL2_CipNo).Text <> tdbdCipID.Columns(0).Text Then
                    tdbg2.Columns(COL2_CipNo).Text = ""
                End If
                tdbg2.Columns(COL2_CipID).Text = tdbdCipID.Columns(2).Text
            Case COL2_CipNo
        End Select
    End Sub

    Private Sub tdbg2_ButtonClick(sender As Object, e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.ButtonClick
        If DxxFormat.LoadFormNotINV = 0 Then Exit Sub
        Select Case e.ColIndex
            Case COL2_CreditAccountID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, tdbg2.Columns(tdbg2.Col).DataField)
                If tdbd Is Nothing Then Exit Select
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                If dr Is Nothing Then Exit Sub
                AfterColUpdate(tdbg2.Col, dr)
            Case COL2_DebitAccountID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, tdbg2.Columns(tdbg2.Col).DataField)
                If tdbd Is Nothing Then Exit Select
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                If dr Is Nothing Then Exit Sub
                AfterColUpdate(tdbg2.Col, dr)
        End Select
    End Sub

    Private Sub AfterColUpdate(ByVal iCol As Integer, ByVal dr() As DataRow)
        Dim iRow As Integer = tdbg2.Row
        If dr Is Nothing OrElse dr.Length = 0 Then
            Dim row As DataRow = Nothing
            AfterColUpdate(iCol, row)
        ElseIf dr.Length = 1 Then
            If tdbg2.Bookmark <> tdbg2.Row AndAlso tdbg2.RowCount = tdbg2.Row Then 'Đang đứng dòng mới
                Dim dr1 As DataRow = dtGrid2.NewRow
                dtGrid2.Rows.InsertAt(dr1, tdbg2.Row)
                tdbg2.Bookmark = tdbg2.Row
            End If
            AfterColUpdate(iCol, dr(0))
        Else
            For Each row As DataRow In dr
                tdbg2.Bookmark = iRow
                tdbg2.Row = iRow
                AfterColUpdate(iCol, row, True)
                tdbg2.UpdateData()
                iRow += 1
            Next
            tdbg2.Focus()
        End If
    End Sub

    Private Sub AfterColUpdate(ByVal iCol As Integer, ByVal dr As DataRow, Optional ByVal bMultiRow As Boolean = False)
        Select Case iCol
            Case COL2_CreditAccountID
                If dr Is Nothing OrElse dr.Item("AccountID").ToString = "" Then
                    tdbg2.Columns(iCol).Text = ""
                Else
                    tdbg2.Columns(iCol).Text = dr.Item("AccountID").ToString
                    If tdbg2.Columns(iCol).Value.ToString <> "" And tdbcCurrencyID.Text <> "" Then
                        tdbg2.Columns(COL2_ExchangeRate).Text = Convert.ToDouble(SQLMoney(tdbcCurrencyID.Columns("ExchangeRate").Text, DxxFormat.ExchangeRateDecimals)).ToString
                    End If
                End If
                If clsFilterDropdown.IsNewFilter Then
                    tdbg2.Col = IndexOfColumn(tdbg2, tdbg2.Columns(iCol).DataField) + 1
                    tdbg2.Col = IndexOfColumn(tdbg2, tdbg2.Columns(iCol).DataField)
                End If
            Case COL2_DebitAccountID
                If dr Is Nothing OrElse dr.Item("AccountID").ToString = "" Then
                    tdbg2.Columns(iCol).Text = ""
                Else
                    tdbg2.Columns(iCol).Text = dr.Item("AccountID").ToString
                    If tdbg2.Columns(iCol).Value.ToString <> "" And tdbcCurrencyID.Text <> "" Then
                        tdbg2.Columns(COL2_ExchangeRate).Text = Convert.ToDouble(SQLMoney(tdbcCurrencyID.Columns("ExchangeRate").Text, DxxFormat.ExchangeRateDecimals)).ToString
                    End If
                End If
                If clsFilterDropdown.IsNewFilter Then
                    tdbg2.Col = IndexOfColumn(tdbg2, tdbg2.Columns(iCol).DataField) + 1
                    tdbg2.Col = IndexOfColumn(tdbg2, tdbg2.Columns(iCol).DataField)
                End If
        End Select
    End Sub

    Private Sub tdbg2_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.AfterColUpdate
        Select Case e.ColIndex
            Case COL2_ChangeNo
            Case COL2_VoucherTypeID
            Case COL2_VoucherDesc
            Case COL2_RefNo
            Case COL2_Serial
            Case COL2_TransDesc
            Case COL2_ObjectTypeID
            Case COL2_ObjectID
            Case COL2_CurrencyID
            Case COL2_ExchangeRate
            Case COL2_DebitAccountID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, e.Column.DataColumn.DataField)
                If tdbd Is Nothing Then Exit Select
                If clsFilterDropdown.IsNewFilter Then
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                    AfterColUpdate(e.ColIndex, dr)
                    Exit Sub
                Else ' Nhập liệu dạng cũ (xổ dropdown)
                    If tdbg2.Columns(e.ColIndex).Text = "" Then Exit Sub
                    Dim row As DataRow = ReturnDataRow(tdbdDebitAccountID, "AccountID =" & SQLString(tdbg2.Columns(COL2_DebitAccountID).Text))
                    AfterColUpdate(e.ColIndex, row)
                End If
            Case COL2_CreditAccountID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, e.Column.DataColumn.DataField)
                If tdbd Is Nothing Then Exit Select
                If clsFilterDropdown.IsNewFilter Then
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                    AfterColUpdate(e.ColIndex, dr)
                    Exit Sub
                Else ' Nhập liệu dạng cũ (xổ dropdown)
                    If tdbg2.Columns(e.ColIndex).Text = "" Then Exit Sub
                    Dim row As DataRow = ReturnDataRow(tdbdCreditAccountID, "AccountID =" & SQLString(tdbg2.Columns(COL2_CreditAccountID).Text))
                    AfterColUpdate(e.ColIndex, row)
                End If
            Case COL2_Amount
            Case COL2_SourceID
            Case COL2_CipNo
            Case COL2_CipNo
        End Select
    End Sub

    Private Sub tdbg2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg2.KeyPress
        Select Case tdbg2.Col
            Case COL2_ChangeNo
            Case COL2_VoucherTypeID
            Case COL2_VoucherDesc
            Case COL2_RefNo
            Case COL2_Serial
            Case COL2_TransDesc
            Case COL2_ObjectTypeID
            Case COL2_ObjectID
            Case COL2_CurrencyID
            Case COL2_ExchangeRate
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL2_DebitAccountID
            Case COL2_CreditAccountID
            Case COL2_Amount
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL2_SourceID
            Case COL2_CipNo
            Case COL2_CipNo
        End Select
    End Sub

    Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
        If clsFilterDropdown.CheckKeydownFilterDropdown(tdbg2, e) Then
            Select Case tdbg2.Col
                Case COL2_CreditAccountID
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, tdbg2.Columns(tdbg2.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                    If dr Is Nothing Then Exit Sub
                    AfterColUpdate(tdbg2.Col, dr)
                    Exit Sub
                Case COL2_DebitAccountID
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg2, tdbg2.Columns(tdbg2.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg2, e, tdbd)
                    If dr Is Nothing Then Exit Sub
                    AfterColUpdate(tdbg2.Col, dr)
                    Exit Sub
            End Select
        End If
        If e.KeyCode = Keys.F7 Then
            HotKeyF7(tdbg2)
            Exit Sub
        End If

        If e.KeyCode = Keys.Enter Then
            If tdbg2.Col = iLastCol2 Then HotKeyEnterGrid(tdbg2, COL2_RefNo, e)
            Exit Sub
        End If

        HotKeyDownGrid(e, tdbg2, COL2_RefNo)

    End Sub

    Private Sub tdbg2_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.HeadClick
        Select Case e.ColIndex
            Case COL2_TransDesc
                CopyColumns(tdbg2, e.ColIndex, tdbg2.Columns(e.ColIndex).Value.ToString, tdbg2.Row)
        End Select
    End Sub

    Private Sub chkDepreciationChange_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDepreciationChange.Click
        optStopDepreciation.Enabled = chkDepreciationChange.Checked
        optReDepreciation.Enabled = chkDepreciationChange.Checked
        optStopDepreciation.Checked = chkDepreciationChange.Checked
    End Sub

    Private Sub chkUsingChange_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUsingChange.Click
        optStopUse.Enabled = chkUsingChange.Checked
        optReUse.Enabled = chkUsingChange.Checked
        tdbcAssetWHID.Enabled = chkUsingChange.Checked
        optStopUse.Checked = chkUsingChange.Checked
    End Sub

    Private Sub chkDepreciationTime_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDepreciationTime.Click
        txtServiceLife.Enabled = chkDepreciationTime.Checked
        If chkDepreciationTime.Checked Then
            txtServiceLife.BackColor = COLOR_BACKCOLOROBLIGATORY
        Else
            txtServiceLife.BackColor = Color.White
        End If
    End Sub

    Private Sub chkDistributeChange_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDistributeChange.Click
        tdbg1.Enabled = chkDistributeChange.Checked
    End Sub
    'ID : 214915
    Private Sub chkReceiveChange_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkReceiveChange.Click
        tdbcObjectTypeID.Enabled = chkReceiveChange.Checked
        tdbcObjectID.Enabled = chkReceiveChange.Checked
        tdbcEmployeeID.Enabled = chkReceiveChange.Checked
        tdbcLocationID.Enabled = chkReceiveChange.Checked

    End Sub
    'ID : 214915
    Private Sub chkManagementChange_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkManagementChange.Click
        tdbcManagementObjTypeID.Enabled = chkManagementChange.Checked 'IDS
        tdbcManagementObjID.Enabled = chkManagementChange.Checked
        tdbcEmployeeID.Enabled = chkReceiveChange.Checked
        tdbcLocationID.Enabled = chkReceiveChange.Checked

    End Sub

    Private Sub chkUseAccount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseAccount.Click
        If chkUseAccount.Checked And (chkIsChangeAssetAccount.Checked = False And chkIsChangeDepAccount.Checked = False) Then
            tabMain.SelectedTab = TabPage2
        Else
            tabMain.SelectedTab = TabPage1
        End If
    End Sub

    Private Sub txtServiceLife_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtServiceLife.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
    End Sub

    Private Sub tabMain_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles tabMain.Selecting
        If chkUseAccount.Checked Then
            If (chkIsChangeAssetAccount.Checked Or chkIsChangeDepAccount.Checked) Then '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
                If tabMain.SelectedIndex = 0 Then
                    tabMain.SelectedTab = TabPage1
                Else
                    If tabMain.SelectedIndex = 2 Then
                        tabMain.SelectedTab = TabPage3
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                e.Cancel = False
                If tabMain.SelectedIndex = 2 Then
                    tabMain.SelectedTab = TabPage3
                End If
            End If
        Else

            If tabMain.SelectedIndex = 0 Then
                tabMain.SelectedTab = TabPage1
            Else
                If tabMain.SelectedIndex = 2 Then
                    tabMain.SelectedTab = TabPage3
                Else
                    e.Cancel = True
                End If
            End If

        End If

    End Sub

#Region "Events tdbcLocationID with txtLocationName"

    Private Sub tdbcLocationID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLocationID.SelectedValueChanged
        If tdbcLocationID.SelectedValue Is Nothing Then
            txtLocationName.Text = ""
        Else
            txtLocationName.Text = tdbcLocationID.Columns(0).Value.ToString
        End If
    End Sub

    Private Sub tdbcLocationID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLocationID.LostFocus
        If tdbcLocationID.FindStringExact(tdbcLocationID.Text) = -1 Then
            tdbcLocationID.Text = ""
        End If
    End Sub

#End Region

    Private Sub tdbcName_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLocationID.Close
        tdbcName_Validated(sender, Nothing)
    End Sub

    Private Sub tdbcName_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLocationID.Validated
        Dim tdbc As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        tdbc.Text = tdbc.WillChangeToText
    End Sub

#Region "TdbgInfo"
    Private Sub tdbgInfo_FetchCellStyle(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.FetchCellStyleEventArgs) Handles tdbgInfo.FetchCellStyle
        Select Case e.Col
            Case COL3_Caption84, COL3_Caption01
                If L3Bool(tdbgInfo(e.Row, COL3_IsUse).ToString) Then
                    e.CellStyle.BackColor = COLOR_BACKCOLOROBLIGATORY
                Else
                    e.CellStyle.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
                    e.CellStyle.Locked = True
                End If
        End Select
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0205
    '# Created User: HUỲNH KHANH
    '# Created Date: 17/10/2014 09:35:13
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0205(ByVal sChangeNo As String, ByVal iMode As Integer) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Grid thong tin bo sung" & vbCrLf)
        sSQL &= "Exec D02P0205 "
        sSQL &= SQLString(sChangeNo) & COMMA 'ChangeNo, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(iMode) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0205s
    '# Created User: HUỲNH KHANH
    '# Created Date: 17/10/2014 09:37:58
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0205s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbgInfo.RowCount - 1
            If sSQL.ToString = "" And sRet.ToString = "" Then sSQL.Append("-- Bổ sung lưu thêm bảng D02T0205" & vbCrLf)
            sSQL.Append("Insert Into D02T0205(")
            sSQL.Append("RefID, ChangeNo, Caption84U, ")
            sSQL.Append("Caption01U, DataType, IsUse")
            sSQL.Append(") Values(" & vbCrLf)
            sSQL.Append(SQLString(tdbgInfo(i, COL3_RefID)) & COMMA) 'RefID, varchar[50], NOT NULL
            sSQL.Append(SQLString(txtChangeNo.Text) & COMMA) 'ChangeNo, varchar[50], NOT NULL
            sSQL.Append(SQLStringUnicode(tdbgInfo(i, COL3_Caption84), gbUnicode, True) & COMMA) 'Caption84U, nvarchar[200], NOT NULL
            sSQL.Append(SQLStringUnicode(tdbgInfo(i, COL3_Caption01), gbUnicode, True) & COMMA) 'Caption01U, nvarchar[200], NOT NULL
            sSQL.Append(SQLString(tdbgInfo(i, COL3_DataType)) & COMMA) 'DataType, varchar[10], NOT NULL
            sSQL.Append(SQLNumber(tdbgInfo(i, COL3_IsUse))) 'IsUse, int, NOT NULL
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0205
    '# Created User: HUỲNH KHANH
    '# Created Date: 17/10/2014 09:39:44
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0205() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Xóa trước khi cập nhật lại" & vbCrLf)
        sSQL &= "Delete From D02T0205"
        sSQL &= " Where ChangeNo = " & SQLString(txtChangeNo.Text)
        Return sSQL
    End Function

#End Region

    Private Sub HideControlBySystem()
        If Not D02Systems.IsAllowChangeAccount Then
            pnlNote.Top = chkIsChangeAssetAccount.Top
        End If
    End Sub

#Region "Events tdbcDescPriorityID with txtDescPriorityName"

    Private Sub tdbcDescPriorityID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescPriorityID.SelectedValueChanged
        If tdbcDescPriorityID.SelectedValue Is Nothing Then
            txtDescPriorityName.Text = ""
        Else
            txtDescPriorityName.Text = tdbcDescPriorityID.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcDescPriorityID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcDescPriorityID.LostFocus
        If tdbcDescPriorityID.FindStringExact(tdbcDescPriorityID.Text) = -1 Then
            tdbcDescPriorityID.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcManagementObjTypeID with txtManagementObjName"
    Private Sub tdbcManagementObjTypeID_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObjTypeID.Close
        If tdbcManagementObjTypeID.FindStringExact(tdbcManagementObjTypeID.Text) = -1 Then
            tdbcManagementObjTypeID.Text = ""
            tdbcManagementObjID.Text = ""
            txtManagementObjName.Text = ""
            tdbcManagementObjTypeID.Focus()
        End If

    End Sub

    Private Sub tdbcManagementObjTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcManagementObjTypeID.SelectedValueChanged

        LoadDataSource(tdbcManagementObjID, ReturnTableFilter(dtManagementObjID, " ObjectTypeID = " & SQLString(tdbcManagementObjTypeID.Text), True), gbUnicode)
        txtManagementObjName.Text = ""
    End Sub

    Private Sub tdbcManagementObjTypeID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcManagementObjTypeID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcManagementObjTypeID.Text = ""
            tdbcManagementObjID.Text = ""
            txtManagementObjName.Text = ""
        End If

    End Sub

    Private Sub tdbcManagementObjID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcManagementObjID.SelectedValueChanged
        If tdbcManagementObjID.SelectedValue Is Nothing Then
            txtManagementObjName.Text = ""
        Else
            txtManagementObjName.Text = tdbcManagementObjID.Columns("ObjectName").Value.ToString
        End If
    End Sub

    Private Sub tdbcManagementObjID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcManagementObjID.LostFocus
        If tdbcManagementObjID.FindStringExact(tdbcManagementObjID.Text) = -1 Then
            tdbcManagementObjID.Text = ""
        End If
    End Sub

#End Region

    Private Sub chkIsChangeAssetAccount_CheckedChanged(sender As Object, e As EventArgs) Handles chkIsChangeAssetAccount.CheckedChanged, chkIsChangeDepAccount.CheckedChanged
        '10/11/2021, Phạm Thị Mỹ Tiên:id 191804-[LAF] D02 - Phát triển nghiệp vụ tác động Thay đổi TK Tài sản, TK khấu hao của Mã TSCĐ
        If (chkIsChangeAssetAccount.Checked Or chkIsChangeDepAccount.Checked) And chkUseAccount.Checked Then
            EnabledTabPage(New TabPage() {TabPage2}, False)
        Else
            EnabledTabPage(New TabPage() {TabPage2}, True)
        End If
    End Sub

End Class