'#-------------------------------------------------------------------------------------
'# Created Date: 05/05/2011 8:32:43 AM
'# Created User: Nguyễn Đức Trọng
'# Modify Date: 05/05/2011 8:32:43 AM
'# Modify User: Nguyễn Đức Trọng
'#-------------------------------------------------------------------------------------
Public Class D02F1021
	Dim report As D99C2003
	Private _formIDPermission As String = "D02F1021"
	Public WriteOnly Property FormIDPermission() As String
		Set(ByVal Value As String)
			       _formIDPermission = Value
		   End Set
	End Property


#Region "Const of tdbg"
    Private Const COL_ChangeNo As Integer = 0         ' Mã nghiệp vụ
    Private Const COL_ChangeName As Integer = 1       ' Tên nghiệp vụ
    Private Const COL_Notes1 As Integer = 2           ' Chú ý 1
    Private Const COL_Disabled As Integer = 3         ' Không sử dụng
    Private Const COL_CreateDate As Integer = 4       ' CreateDate
    Private Const COL_CreateUserID As Integer = 5     ' CreateUserID
    Private Const COL_LastModifyDate As Integer = 6   ' LastModifyDate
    Private Const COL_LastModifyUserID As Integer = 7 ' LastModifyUserID
#End Region

    Private dtGrid, dtCaptionCols As DataTable
    Dim bRefreshFilter As Boolean
    Dim sFilter As New System.Text.StringBuilder()

    Private Sub D02F1021_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        Me.Cursor = Cursors.WaitCursor
        gbEnabledUseFind = False
        ResetColorGrid(tdbg)
        Loadlanguage()
        LoadTDBGrid()
        SetShortcutPopupMenu(Me, TableToolStrip, ContextMenuStrip1)
        SetResolutionForm(Me, ContextMenuStrip1)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_muc_nghiep_vu_tac_dong_-_D02F1021") & UnicodeCaption(gbUnicode) 'Danh móc nghiÖp vó tÀc ¢èng - D02F1021
        '================================================================ 
        tdbg.Columns("ChangeNo").Caption = rl3("Ma_nghiep_vu") 'Mã nghiệp vụ
        tdbg.Columns("ChangeName").Caption = rl3("Ten_nghiep_vu") 'Tên nghiệp vụ
        tdbg.Columns("Notes1").Caption = rl3("Ghi_chu_1") 'Ghi chú 1
        tdbg.Columns("Disabled").Caption = rl3("KSD") 'Không sử dụng    
        '================================================================ 
        chkShowDisabled.Text = rl3("Hien_thi_danh_muc_khong_su_dung") 'Hiển thị danh mục không sử dụng
    End Sub

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String
        sSQL = "Select      Distinct ChangeNo, ChangeName" & UnicodeJoin(gbUnicode) & " As ChangeName, " & vbCrLf
        sSQL &= "           Notes1" & UnicodeJoin(gbUnicode) & " As Notes1, " & vbCrLf
        sSQL &= "           Disabled, CreateDate, CreateUserID, LastModifyDate,LastModifyUserID" & vbCrLf
        sSQL &= "From       D02T0201 WITH(NOLOCK)" & vbCrLf
        sSQL &= "Group By   ChangeNo, ChangeName" & UnicodeJoin(gbUnicode) & ", Notes1" & UnicodeJoin(gbUnicode) & ", Disabled, CreateDate, CreateUserID, LastModifyDate,LastModifyUserID" & vbCrLf
        sSQL &= "Order By   ChangeNo"
        dtGrid = ReturnDataTable(sSQL)

        gbEnabledUseFind = dtGrid.Rows.Count > 0
        If FlagAdd Then ' Thêm mới thì set Filter = "" và sFind =""
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFind = ""
        End If

        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()

        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("ChangeNo = " & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0))
        End If

        If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString

        If Not chkShowDisabled.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "Disabled = 0"
        End If
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        CheckMenu(_formIDPermission, TableToolStrip, tdbg.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
        tsmEditOther.Enabled = tsbEdit.Enabled
        mnsEditOther.Enabled = tsbEdit.Enabled
        FooterTotalGrid(tdbg, COL_ChangeName)
    End Sub

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
	Dim gbEnabledUseFind As Boolean = False
    'Cần sửa Tìm kiếm như sau:
	'Bỏ sự kiện Finder_FindClick.
	'Sửa tham số Me.Name -> Me
	'Phải tạo biến properties có tên chính xác strNewFind và strNewServer
	'Sửa gdtCaptionExcel thành dtCaptionCols: biến toàn cục trong form
	'Nếu có F12 dùng D09U1111 thì Sửa dtCaptionCols thành ResetTableByGrid(usrOption, dtCaptionCols.DefaultView.ToTable)
    Private sFind As String = ""
	Public WriteOnly Property strNewFind() As String
		Set(ByVal Value As String)
			sFind = Value
			ReLoadTDBGrid()'Làm giống sự kiện Finder_FindClick. Ví dụ đối với form Báo cáo thường gọi btnPrint_Click(Nothing, Nothing): sFind = "
		End Set
	End Property


    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me, "0", gbUnicode)
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Menu bar"

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim f As New D02F1022
        With f
            .ChangeNo = ""
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If f.SavedOK Then LoadTDBGrid(True, .ChangeNo)
            .Dispose()
        End With
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim f As New D02F1022
        With f
            .ChangeNo = tdbg.Columns(COL_ChangeNo).Text
            .CreateUserID = tdbg.Columns(COL_CreateUserID).Text
            .CreateDate = tdbg.Columns(COL_CreateDate).Text
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        If Not AllowEdit() Then Exit Sub
        Dim f As New D02F1022
        With f
            .ChangeNo = tdbg.Columns(COL_ChangeNo).Text
            .CreateUserID = tdbg.Columns(COL_CreateUserID).Text
            .CreateDate = tdbg.Columns(COL_CreateDate).Text
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            .Dispose()
        End With
        If f.SavedOK Then LoadTDBGrid(False, tdbg.Columns(COL_ChangeNo).Text)
    End Sub

    Private Sub tsbEditOther_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmEditOther.Click, mnsEditOther.Click
        'If Not AllowEdit() Then Exit Sub
        Dim f As New D02F1022
        With f
            .ChangeNo = tdbg.Columns(COL_ChangeNo).Text
            .CreateUserID = tdbg.Columns(COL_CreateUserID).Text
            .CreateDate = tdbg.Columns(COL_CreateDate).Text
            .FormState = EnumFormState.FormEditOther
            .ShowDialog()
            .Dispose()
        End With
        If f.SavedOK Then LoadTDBGrid(False, tdbg.Columns(COL_ChangeNo).Text)
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowDelete() Then Exit Sub

        Dim sSQL As String = ""
        sSQL = "Delete D02T0204 "
        sSQL &= "Where ChangeNo = " & SQLString(tdbg.Columns(COL_ChangeNo).Text) & vbCrLf
        sSQL &= "Delete D02T0201 "
        sSQL &= "Where ChangeNo =  " & SQLString(tdbg.Columns(COL_ChangeNo).Text)
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL)
        If bRunSQL = True Then
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
            ResetGrid()
            DeleteOK()
        Else
            DeleteNotOK()
        End If
    End Sub

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(Me,tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbPrint.Click, tsmPrint.Click, mnsPrint.Click

        Me.Cursor = Cursors.WaitCursor

        'Dim report As New D99C1003
        'Đưa vể đầu tiên hàm In trước khi gọi AllowPrint()
		If Not AllowNewD99C2003(report, Me) Then Exit Sub
		'************************************
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = "D02R0201"
        Dim sSubReportName As String = "D02R0000"
        Dim sReportCaption As String = ""
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        sReportCaption = rl3("Danh_muc_nghiep_vu_tac_dong") & " - " & sReportName
        sPathReport = UnicodeGetReportPath(gbUnicode, D02Options.ReportLanguage, "") & sReportName & ".rpt"

        sSQL = "Select * From D02T0201 WITH(NOLOCK) Order By ChangeNo "
        sSQLSub = "Select Top 1 * From D91T0025 WITH(NOLOCK) "
        UnicodeSubReport(sSubReportName, sSQLSub, , gbUnicode)

        With report
            .OpenConnection(conn)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(sSQL)
            .PrintReport(sPathReport, sReportCaption)
        End With

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub chkShowDisabled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "Grid"

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If tdbg.FilterActive Then Exit Sub
        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg, e)
    End Sub

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            FilterChangeGrid(tdbg, sFilter)
            ReLoadTDBGrid()
        Catch ex As Exception
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

#End Region

    Private Function AllowEdit() As Boolean
        Dim sSQL As New StringBuilder(132)
        sSQL.Append(" Select Top 1 1 ")
        sSQL.Append(" From D02T0202 WITH(NOLOCK) ")
        sSQL.Append(" Where ChangeNo = " & SQLString(tdbg.Columns(COL_ChangeNo).Text))

        If ExistRecord(sSQL.ToString) Then
            D99C0008.MsgL3(rL3("Du_lieu_da_duoc_su_dung_Ban_khong_the_sua_thong_tin_nay"))
            tdbg.Focus()
            Return False
        End If
        Return True
    End Function

    Private Function AllowDelete() As Boolean
        Dim sSQL As New StringBuilder(132)
        sSQL.Append(" Select Top 1 1 ")
        sSQL.Append(" From D02T0202 WITH(NOLOCK) ")
        sSQL.Append(" Where ChangeNo = " & SQLString(tdbg.Columns(COL_ChangeNo).Text))

        If ExistRecord(sSQL.ToString) Then
            D99C0008.MsgL3(rL3("Du_lieu_da_duoc_su_dung_Ban_khong_the_xoa_thong_tin_nay"))
            tdbg.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Col
            Case COL_Disabled 'Chặn Ctrl + V trên cột Check
                e.Handled = CheckKeyPress(e.KeyChar)
        End Select
    End Sub
End Class