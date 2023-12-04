'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong class này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Diễn giải: Hiển thị thanh finder tìm kiếm
'# Ngày cập nhật cuối cùng: 09/11/2012
'# Người cập nhật cuối cùng: Bùi Thị Thanh Huyền
'# Thêm hàm ShowFormFindClientServer
'# Bổ sung đối số FindServer, FindClient
'# Bổ sung tìm kiếm KHÔNG CHỨA
'#Bổ sung font Unicode cho combo Toán tử
'#Bổ sung Tìm kiếm tỷ lệ %
'#######################################################################################

''' <summary>
''' Liên quan đến Finder
''' </summary>
''' 
Public Class D99C1001

    Private m_FormCaption As String
    Private m_FindCommandCaption As String
    Private m_FrameCaption As String
    Private m_HelpCommandCaption As String
    Private m_CloseCommandCaption As String
    Private m_AndCommandCaption As String
    Private m_OrCommandCaption As String
    Private m_SortCommandCaption As String
    Private m_AND As String
    Private m_OR As String
    Private m_OP_STARTWITH As String
    Private m_OP_CONTAIN As String
    Private m_OP_NOTCONTAIN As String
    Private m_OP_ENDWITH As String
    Private m_OP_EQUAL As String
    Private m_OP_GREATEREQUAL As String
    Private m_OP_GREATER As String
    Private m_OP_LESSEQUAL As String
    Private m_OP_LESS As String
    Private m_OP_BETWEEN As String
    Private m_OP_NOTBETWEEN As String
    Private m_OP_NOTEQUAL As String
    Private m_ASC As String
    Private m_DESCAS As String
    Private m_OKCommandCaption As String
    Private m_DELETECommandCaption As String
    Private m_FIELDNAMECaption As String
    Private m_ORDERNOCaption As String
    Private FlagAddField As Boolean 'Kiểm tra có phải là lần đầu không
    Private m_DropHeight As Integer
    Private m_AllowSortOrder As Boolean

    Public Event FindClick(ByVal ResultWhereClause As Object)
    Public Event FindReportClick(ByVal ResultWhereClause As Object, ByVal ResultWhereClientForReport As Object)

    Public Sub AddFieldName(ByVal sFieldName As String, ByVal sFieldDescription As String, ByVal sDataType As FinderTypeEnum, Optional ByVal nLengh As Integer = 20, Optional ByVal sType As String = "")
        Try

            If FlagAddField Then
                ReDim arrAdvanced(0)
                arrAdvanced(UBound(arrAdvanced)) = sFieldDescription & LM_DLM & sFieldName & LM_DLM & sDataType & LM_DLM & CStr(nLengh) & LM_DLM & sType
                FlagAddField = False
            Else
                ReDim Preserve arrAdvanced(UBound(arrAdvanced) + 1)
                arrAdvanced(UBound(arrAdvanced)) = sFieldDescription & LM_DLM & sFieldName & LM_DLM & sDataType & LM_DLM & CStr(nLengh) & LM_DLM & sType
            End If
        Catch
        End Try
    End Sub

    Public Sub ShowFormFindClientServer(ByVal sFormID As String, ByVal sMode As String, Optional ByVal bUseV1234 As Boolean = False)
        'Update 0/10/2010: Tìm kiếm client nhưng trả ra giá trị vừa server vừa client
        strSQLAdvanced = ""
        strSQLAdvancedForReport = ""

        Dim Frm As New D99F0006
        Frm.LM_FindClient = 2 'True
        Frm.dropH = m_DropHeight
        Frm.Text = m_FormCaption
        Frm.ButFind.Text = m_FindCommandCaption
        Frm.GroupBox1.Text = m_FrameCaption
        'Frm.ButHelp.Text = m_HelpCommandCaption
        Frm.ButClose.Text = m_CloseCommandCaption
        Frm.ButAnd.Text = m_AndCommandCaption & " (F9)"
        Frm.ButOr.Text = m_OrCommandCaption & " (F10)"
        Frm.LM_AND = m_AND
        Frm.LM_OR = m_OR
        Frm.LM_OP_STARTWITH = m_OP_STARTWITH
        Frm.LM_OP_CONTAIN = m_OP_CONTAIN
        Frm.LM_OP_NOTCONTAIN = m_OP_NOTCONTAIN
        Frm.LM_OP_ENDWITH = m_OP_ENDWITH
        Frm.LM_OP_EQUAL = m_OP_EQUAL
        Frm.LM_OP_GREATEREQUAL = m_OP_GREATEREQUAL
        Frm.LM_OP_GREATER = m_OP_GREATER
        Frm.LM_OP_LESSEQUAL = m_OP_LESSEQUAL
        Frm.LM_OP_LESS = m_OP_LESS
        Frm.LM_OP_BETWEEN = m_OP_BETWEEN
        Frm.LM_OP_NOTBETWEEN = m_OP_NOTBETWEEN
        Frm.LM_OP_NOTEQUAL = m_OP_NOTEQUAL
        Frm.LM_ASC = m_ASC
        Frm.LM_DESCAS = m_DESCAS
        Frm.LM_OKCommandCaption = m_OKCommandCaption
        Frm.LM_DELETECommandCaption = m_DELETECommandCaption
        Frm.LM_CloseCommandCaption = m_CloseCommandCaption
        Frm.LM_FIELDNAMECaption = m_FIELDNAMECaption
        Frm.LM_ORDERNOCaption = m_ORDERNOCaption

        '********************
        'Update 10/03/2009
        Frm.Mode = sMode
        Frm.FormID = sFormID
        'Update 29/07/2010
        Frm.UseUnicode = _useUnicode
        'Update 04/10/2010
        Frm.FindServer = bUseV1234 'Nếu Nhập liệu Unicode thì kiểm tra có tìm kiếm theo view V1234

        '********************
        Frm.ShowDialog()
        Frm.Dispose()
        FlagAddField = True

        If strSQLAdvanced = Nothing Then strSQLAdvanced = ""
        If strSQLAdvancedForReport = Nothing Then strSQLAdvancedForReport = ""

        If Trim(strSQLAdvanced) = "" Then
            RaiseEvent FindReportClick(strSQLAdvanced, strSQLAdvancedForReport)
        Else
            RaiseEvent FindReportClick("( " & strSQLAdvanced & " )", "( " & strSQLAdvancedForReport & " )")
        End If

    End Sub

    Public Sub ShowFormFindClient(ByVal sFormID As String, ByVal sMode As String, Optional ByVal bUseV1234 As Boolean = False)
        strSQLAdvanced = ""

        Dim Frm As New D99F0006
        Frm.LM_FindClient = 1 'True
        Frm.dropH = m_DropHeight
        Frm.Text = m_FormCaption
        Frm.ButFind.Text = m_FindCommandCaption
        Frm.GroupBox1.Text = m_FrameCaption
        'Frm.ButHelp.Text = m_HelpCommandCaption
        Frm.ButClose.Text = m_CloseCommandCaption
        Frm.ButAnd.Text = m_AndCommandCaption & " (F9)"
        Frm.ButOr.Text = m_OrCommandCaption & " (F10)"
        Frm.LM_AND = m_AND
        Frm.LM_OR = m_OR
        Frm.LM_OP_STARTWITH = m_OP_STARTWITH
        Frm.LM_OP_CONTAIN = m_OP_CONTAIN
        Frm.LM_OP_NOTCONTAIN = m_OP_NOTCONTAIN
        Frm.LM_OP_ENDWITH = m_OP_ENDWITH
        Frm.LM_OP_EQUAL = m_OP_EQUAL
        Frm.LM_OP_GREATEREQUAL = m_OP_GREATEREQUAL
        Frm.LM_OP_GREATER = m_OP_GREATER
        Frm.LM_OP_LESSEQUAL = m_OP_LESSEQUAL
        Frm.LM_OP_LESS = m_OP_LESS
        Frm.LM_OP_BETWEEN = m_OP_BETWEEN
        Frm.LM_OP_NOTBETWEEN = m_OP_NOTBETWEEN
        Frm.LM_OP_NOTEQUAL = m_OP_NOTEQUAL
        Frm.LM_ASC = m_ASC
        Frm.LM_DESCAS = m_DESCAS
        Frm.LM_OKCommandCaption = m_OKCommandCaption
        Frm.LM_DELETECommandCaption = m_DELETECommandCaption
        Frm.LM_CloseCommandCaption = m_CloseCommandCaption
        Frm.LM_FIELDNAMECaption = m_FIELDNAMECaption
        Frm.LM_ORDERNOCaption = m_ORDERNOCaption

        '********************
        'Update 10/03/2009
        Frm.Mode = sMode
        Frm.FormID = sFormID
        'Update 29/07/2010
        Frm.UseUnicode = _useUnicode
        'Update 04/10/2010
        Frm.FindServer = bUseV1234 'Nếu Nhập liệu Unicode thì kiểm tra có tìm kiếm theo view V1234

        '********************
        Frm.ShowDialog()
        Frm.Dispose()
        FlagAddField = True

        If strSQLAdvanced = Nothing Then strSQLAdvanced = ""

        If Trim(strSQLAdvanced) = "" Then
            RaiseEvent FindClick(strSQLAdvanced)
        Else
            RaiseEvent FindClick("( " & strSQLAdvanced & " )")
        End If

        
    End Sub

    Public Sub ShowFormFind(ByVal sFormID As String, ByVal sMode As String)
        Dim Frm As New D99F0006
        Frm.LM_FindClient = 0 'False
        Frm.dropH = m_DropHeight
        Frm.Text = m_FormCaption
        Frm.ButFind.Text = m_FindCommandCaption
        Frm.GroupBox1.Text = m_FrameCaption
        'Frm.ButHelp.Text = m_HelpCommandCaption
        Frm.ButClose.Text = m_CloseCommandCaption
        Frm.ButAnd.Text = m_AndCommandCaption & " (F9)"
        Frm.ButOr.Text = m_OrCommandCaption & " (F10)"
        Frm.LM_AND = m_AND
        Frm.LM_OR = m_OR
        Frm.LM_OP_STARTWITH = m_OP_STARTWITH
        Frm.LM_OP_CONTAIN = m_OP_CONTAIN
        Frm.LM_OP_NOTCONTAIN = m_OP_NOTCONTAIN
        Frm.LM_OP_ENDWITH = m_OP_ENDWITH
        Frm.LM_OP_EQUAL = m_OP_EQUAL
        Frm.LM_OP_GREATEREQUAL = m_OP_GREATEREQUAL
        Frm.LM_OP_GREATER = m_OP_GREATER
        Frm.LM_OP_LESSEQUAL = m_OP_LESSEQUAL
        Frm.LM_OP_LESS = m_OP_LESS
        Frm.LM_OP_BETWEEN = m_OP_BETWEEN
        Frm.LM_OP_NOTBETWEEN = m_OP_NOTBETWEEN
        Frm.LM_OP_NOTEQUAL = m_OP_NOTEQUAL
        Frm.LM_ASC = m_ASC
        Frm.LM_DESCAS = m_DESCAS
        Frm.LM_OKCommandCaption = m_OKCommandCaption
        Frm.LM_DELETECommandCaption = m_DELETECommandCaption
        Frm.LM_CloseCommandCaption = m_CloseCommandCaption
        Frm.LM_FIELDNAMECaption = m_FIELDNAMECaption
        Frm.LM_ORDERNOCaption = m_ORDERNOCaption
        '********************
        'Update 10/03/2009
        Frm.FormID = sFormID
        Frm.Mode = sMode
        Frm.UseUnicode = _useUnicode
        Frm.FindServer = True
        '********************
        Frm.ShowDialog()
        Frm.Dispose()
        FlagAddField = True

        If strSQLAdvanced = Nothing Then strSQLAdvanced = ""

        If Trim(strSQLAdvanced) = "" Then
            RaiseEvent FindClick(strSQLAdvanced)
        Else
            RaiseEvent FindClick("( " & strSQLAdvanced & " )")
        End If

    End Sub

    Private Sub convertValueOperator()
        If _useUnicode Then
            m_OP_STARTWITH = "|~  Bắt đầu là"
            m_OP_CONTAIN = "~   Có chứa"
            m_OP_NOTCONTAIN = "~   Không có chứa"
            m_OP_ENDWITH = "~|  Kết thúc là"
            m_OP_EQUAL = "=  Bằng"
            m_OP_GREATEREQUAL = ">=  Lớn hơn hoặc bằng"
            m_OP_GREATER = ">    Lớn hơn"
            m_OP_LESSEQUAL = "<=  Nhỏ hơn hoặc bằng"
            m_OP_LESS = "<    Nhỏ hơn"
            m_OP_BETWEEN = "[ ]  Trong khoảng"
            m_OP_NOTBETWEEN = "] [  Ngoài khoảng"
            m_OP_NOTEQUAL = "< > Khác"
        Else
            m_OP_STARTWITH = "|~  Baét ñaàu laø"
            m_OP_CONTAIN = "~   Coù chöùa"
            m_OP_NOTCONTAIN = "~   Khoâng coù chöùa"
            m_OP_ENDWITH = "~|  Keát thuùc laø"
            m_OP_EQUAL = "=  Baèng"
            m_OP_GREATEREQUAL = ">=  Lôùn hôn hoaëc baèng"
            m_OP_GREATER = ">    Lôùn hôn"
            m_OP_LESSEQUAL = "<=  Nhoû hôn hoaëc baèng"
            m_OP_LESS = "<    Nhoû hôn"
            m_OP_BETWEEN = "[ ]  Trong khoaûng"
            m_OP_NOTBETWEEN = "] [  Ngoaøi khoaûng"
            m_OP_NOTEQUAL = "< > Khaùc"
        End If

    End Sub

    Public Sub New()
        FlagAddField = True
        m_FormCaption = "Lemon Advanced Finder"
        m_FindCommandCaption = "Tìm &kiếm"
        m_FrameCaption = "Điều kiện tìm kiếm"
        m_HelpCommandCaption = "Trợ &giúp"
        m_CloseCommandCaption = "Đó&ng"
        m_AndCommandCaption = "&Và"
        m_OrCommandCaption = "&Hoặc"
        m_SortCommandCaption = "&Thứ tự"
        m_AND = "Và"
        m_OR = "Hoặc"
        convertValueOperator()
        'm_OP_STARTWITH = "|~  Baét ñaàu laø"
        'm_OP_CONTAIN = "~   Coù chöùa"
        'm_OP_NOTCONTAIN = "~   Khoâng coù chöùa"
        'm_OP_ENDWITH = "~|  Keát thuùc laø"
        'm_OP_EQUAL = "=  Baèng"
        'm_OP_GREATEREQUAL = ">=  Lôùn hôn hoaëc baèng"
        'm_OP_GREATER = ">    Lôùn hôn"
        'm_OP_LESSEQUAL = "<=  Nhoû hôn hoaëc baèng"
        'm_OP_LESS = "<    Nhoû hôn"
        'm_OP_BETWEEN = "[ ]  Trong khoaûng"
        'm_OP_NOTBETWEEN = "] [  Ngoaøi khoaûng"
        'm_OP_NOTEQUAL = "< > Khaùc"
        m_ASC = "Tăng dần"
        m_DESCAS = "Giảm dần"
        m_OKCommandCaption = "&Đồng ý"
        m_DELETECommandCaption = "&Xoá"
        m_CloseCommandCaption = "Đó&ng"
        m_FIELDNAMECaption = "Tên trường"
        m_ORDERNOCaption = "Thứ tự sắp xếp"
        m_DropHeight = 3000
        m_AllowSortOrder = True
    End Sub

    ''' <summary>
    ''' Khởi tạo Finder với biến ngôn ngữ truyền vào
    ''' </summary>
    ''' <param name="language">Ngôn ngữ cần sử dụng</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal language As EnumLanguage)
        _Language = language
        ChangeLanguage()
    End Sub

    Private Sub ChangeLanguage()
        m_FormCaption = r("Tim_kiemF") & UnicodeCaption(_useUnicode)
        If _Language = EnumLanguage.Vietnamese Then
            FlagAddField = True
            ' m_FormCaption = "TØm kiÕm" & UnicodeCaption(_useUnicode)
            m_FindCommandCaption = "Tìm &kiếm"
            m_FrameCaption = "Điều kiện tìm kiếm"
            m_HelpCommandCaption = "Trợ &giúp"
            m_CloseCommandCaption = "Đó&ng"
            m_AndCommandCaption = "&Và"
            m_OrCommandCaption = "&Hoặc"
            m_SortCommandCaption = "&Thứ tự"
            m_AND = "Và"
            m_OR = "Hoặc"
            'm_OP_STARTWITH = "|~  Baét ñaàu laø"
            'm_OP_CONTAIN = "~   Coù chöùa"
            'm_OP_NOTCONTAIN = "~   Khoâng coù chöùa"
            'm_OP_ENDWITH = "~|  Keát thuùc laø"
            'm_OP_EQUAL = "=  Baèng"
            'm_OP_GREATEREQUAL = ">=  Lôùn hôn hoaëc baèng"
            'm_OP_GREATER = ">    Lôùn hôn"
            'm_OP_LESSEQUAL = "<=  Nhoû hôn hoaëc baèng"
            'm_OP_LESS = "<    Nhoû hôn"
            'm_OP_BETWEEN = "[ ]  Trong khoaûng"
            'm_OP_NOTBETWEEN = "] [  Ngoaøi khoaûng"
            'm_OP_NOTEQUAL = "< > Khaùc"
            convertValueOperator()
            m_ASC = "Tăng dần"
            m_DESCAS = "Giảm dần"
            m_OKCommandCaption = "&Đồng ý"
            m_DELETECommandCaption = "&Xoá"
            m_FIELDNAMECaption = "Tên trường"
            m_ORDERNOCaption = "Thứ tự sắp xếp"
            m_DropHeight = 3000
            m_AllowSortOrder = True
        ElseIf _Language = EnumLanguage.English Then
            FlagAddField = True
            ' m_FormCaption = "Advanced Finder" & UnicodeCaption(_useUnicode)
            m_FindCommandCaption = "&Find"
            m_FrameCaption = "Find Information" '"Find Condition"
            m_HelpCommandCaption = "&Help"
            m_CloseCommandCaption = "&Close"
            m_AndCommandCaption = "&And"
            m_OrCommandCaption = "&Or"
            m_SortCommandCaption = "&Order"
            m_AND = "And"
            m_OR = "Or"
            m_OP_STARTWITH = "|~  Beginning with"
            m_OP_CONTAIN = "~   Including"
            m_OP_NOTCONTAIN = "~   Not Including"
            m_OP_ENDWITH = "~|  Ending with"
            m_OP_EQUAL = "=  With"
            m_OP_GREATEREQUAL = ">=  More than or equal"
            m_OP_GREATER = ">    More than"
            m_OP_LESSEQUAL = "<=  Less than or equal"
            m_OP_LESS = "<    Less than"
            m_OP_BETWEEN = "[ ]  Within"
            m_OP_NOTBETWEEN = "] [  Out of"
            m_OP_NOTEQUAL = "< >  Different from"
            m_ASC = "Increase"
            m_DESCAS = "Decease"
            m_OKCommandCaption = "&OK"
            m_DELETECommandCaption = "&Delete"
            m_FIELDNAMECaption = "Field Name"
            m_ORDERNOCaption = "Order"
            m_DropHeight = 3000
            m_AllowSortOrder = True
        End If
    End Sub

    Private _Language As EnumLanguage
    ''' <summary>
    ''' Thay đổi ngôn ngữ sử dụng
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property Language() As EnumLanguage
        Set(ByVal value As EnumLanguage)
            _Language = value
            ChangeLanguage()
        End Set
    End Property

    'Public WriteOnly Property CaptionFormFind() As String
    '    Set(ByVal vNewValue As String)
    '        m_FormCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCommandFind() As String
    '    Set(ByVal vNewValue As String)
    '        m_FindCommandCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionFrameFind() As String
    '    Set(ByVal vNewValue As String)
    '        m_FrameCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCommandHelp() As String
    '    Set(ByVal vNewValue As String)
    '        m_HelpCommandCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCommandClose() As String
    '    Set(ByVal vNewValue As String)
    '        m_CloseCommandCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCommandAnd() As String
    '    Set(ByVal vNewValue As String)
    '        m_AndCommandCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCommandOr() As String
    '    Set(ByVal vNewValue As String)
    '        m_OrCommandCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCommandSort() As String
    '    Set(ByVal vNewValue As String)
    '        m_SortCommandCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionAnd() As String
    '    Set(ByVal vNewValue As String)
    '        m_AND = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionOr() As String
    '    Set(ByVal vNewValue As String)
    '        m_OR = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionSTARTWITH() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_STARTWITH = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCONTAIN() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_CONTAIN = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionNOTCONTAIN() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_NOTCONTAIN = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionENDWITH() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_ENDWITH = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionEQUAL() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_EQUAL = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionGREATEREQUAL() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_GREATEREQUAL = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionGREATER() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_GREATER = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionLESSEQUAL() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_LESSEQUAL = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionLESS() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_LESS = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionBETWEEN() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_BETWEEN = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionNOTBETWEEN() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_NOTBETWEEN = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionNOTEQUAL() As String
    '    Set(ByVal vNewValue As String)
    '        m_OP_NOTEQUAL = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionASC() As String
    '    Set(ByVal vNewValue As String)
    '        m_ASC = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionDESCAS() As String
    '    Set(ByVal vNewValue As String)
    '        m_DESCAS = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCommandOK() As String
    '    Set(ByVal vNewValue As String)
    '        m_OKCommandCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionCommandDELETE() As String
    '    Set(ByVal vNewValue As String)
    '        m_DELETECommandCaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionFIELDNAME() As String
    '    Set(ByVal vNewValue As String)
    '        m_FIELDNAMECaption = vNewValue
    '    End Set
    'End Property

    'Public WriteOnly Property CaptionORDERNO() As String
    '    Set(ByVal vNewValue As String)
    '        m_ORDERNOCaption = vNewValue
    '    End Set
    'End Property

    Public Property AllowSortOrder() As Boolean
        Set(ByVal vNewValue As Boolean)
            m_AllowSortOrder = vNewValue
        End Set
        Get
            AllowSortOrder = m_AllowSortOrder
        End Get
    End Property

    Private _useUnicode As Boolean
    Public WriteOnly Property UseUnicode() As Boolean
        Set(ByVal vNewValue As Boolean)
            _useUnicode = vNewValue
        End Set
    End Property

End Class
