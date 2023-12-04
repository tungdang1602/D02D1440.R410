Imports System.Collections
Imports System.Text
''' <summary>
''' D02E1440: Chứa các màn hình Danh mục
''' Sub Main và các vấn đề liên quan đến việc khởi động exe con
''' </summary>

Module D02X0000

    Public Sub Main()
        SetSysDateTime()
#If DEBUG Then 'Nếu đang ở trạng thái DEBUG thì ...
        'MakeVirtualConnection() 'tạo kết nối ảo
        'CheckDLL() 'Kiểm tra các DLL tương thích và các file Module hợp lệ
        'SaveParameter() 'Gán giá trị các thông số vào Registry
#Else 'Đang trong trạng thái thực thi exe
        If PrevInstance() Then End 'Kiểm tra nếu chương trình đã chạy rồi thì END
        ReadLanguage() 'Đọc biến ngôn ngữ ở đây nhằm mục đích để báo lỗi theo ngôn ngữ cho những phần sau
        If Not CheckSecurity() Then End 'Kiểm tra an toàn cho chương trình, nếu không an toàn thì END
#End If
        GetAllParameter() 'Đọc các giá trị từ Registry lưu vào biến toàn cục
        '   gsConnectionString = "Data Source=" & gsServer & ";Initial Catalog=" & gsCompanyID & ";User ID=" & gsConnectionUser & ";Password=" & gsPassword & ";Connect Timeout = 0" 'Tạo chuỗi kết nối dùng cho toàn bộ project            
        If Not CheckConnection() Then End 'Kiểm tra nối không kết nối được với Server thì END
        'Update 19/11/2010: Kiểm tra đồng bộ exe và fix 
        If Not CheckExeFixSynchronous(My.Application.Info.AssemblyName) Then End

        If Not CheckOther() Then End 'Vì lý do gì đó, có thể kiểm tra một điều kiện không hợp lệ và có thể kết thúc chương trình
        'Tới đây quá trình kiểm tra cho modlue đã hoàn thành, không còn lệnh END để kết thúc chương trình nữa
        'LoadFormats()
        LoadFormatsNew()
        LoadOptions() 'Load các thông số cho phần tùy chọn
        LoadOthers() 'Các lập trình viên có thể load những thứ khác ở đây

        'Xóa Registry
#If DEBUG Then
        gbUnicode = True
        PARA_FormID = "D02F1021"
#Else
        D99C0007.RegDeleteExe(EXECHILD)
#End If

        'Hiển thị form tương ứng

Select Case PARA_FormID
'Gọi form nhận tham số
            Case Else

                Try
                    'Gọi form không nhận tham số. Default 
                    Dim frm As New Form
                    Dim frmName As String = PARA_FormID
                    frmName = System.Reflection.Assembly.GetEntryAssembly.GetName.Name & "." & frmName
                    frm = DirectCast(System.Reflection.Assembly.GetEntryAssembly.CreateInstance(frmName), Form)
                    frm.ShowInTaskbar = True
                    frm.ShowDialog()
                    frm.Dispose()
                Catch ex As Exception
                    D99C0008.MsgL3(ex.Message)
                End Try
        End Select

        KillChildProcess(MODULED02)
    End Sub

    Private Function CheckOther() As Boolean
        Return True
    End Function

    Private Sub LoadOthers()
        LoadUserKey()
        GeneralItems()
    End Sub

    Private Function PrevInstance() As Boolean
        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            Return True
        End If
        Return False
    End Function

    Private Sub ReadLanguage()
        Dim sLanguage As String = GetSetting("Lemon3 System Module", "Caption Setting", "Language", "0")
        If sLanguage = "0" Then
            geLanguage = EnumLanguage.Vietnamese
            gsLanguage = "84"
        Else
            geLanguage = EnumLanguage.English
            gsLanguage = "01"
        End If
        D99C0008.Language = geLanguage
    End Sub

    Private Function CheckSecurity() As Boolean
        Dim D00_CompanyName As String
        Dim D00_LegalCopyright As String
        Dim CompanyName As String
        Dim LegalCopyright As String
        If Not System.IO.File.Exists(Application.StartupPath & "\D00E0030.EXE") Then
            If gsLanguage = "84" Then
                MessageBox.Show("Thï tóc gãi nèi bè bÊt híp lÖ! (10)", "Th¤ng bÀo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Else
                MessageBox.Show("Invalid internal system call! (10)", "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
            Return False
        Else
            Dim D00_FiVerInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.StartupPath & "\D00E0030.EXE")
            Dim FiVerInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.StartupPath & "\" & MODULED02 & ".EXE")
            D00_CompanyName = D00_FiVerInfo.CompanyName
            D00_LegalCopyright = D00_FiVerInfo.LegalCopyright
            CompanyName = FiVerInfo.CompanyName
            LegalCopyright = FiVerInfo.LegalCopyright
            If (D00_CompanyName <> CompanyName) OrElse (D00_LegalCopyright <> LegalCopyright) Then
                If gsLanguage = "84" Then
                    MessageBox.Show("Thï tóc gãi nèi bè bÊt híp lÖ! (10)", "Th¤ng bÀo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Else
                    MessageBox.Show("Invalid internal system call! (10)", "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
                Return False
            End If
        End If
        Dim CommandArgs() As String = Environment.GetCommandLineArgs()
        If CommandArgs.Length <> 3 OrElse CommandArgs(1) <> "/DigiNet" OrElse CommandArgs(2) <> "Corporation" Then
            If gsLanguage = "84" Then
                MessageBox.Show("Thï tóc gãi nèi bè bÊt híp lÖ! (12)", "Th¤ng bÀo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Else
                MessageBox.Show("Invalid internal system call! (12)", "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
            Return False
        End If
        Return True
    End Function

    Private Sub Ngoisaoleloi()

        'Co lan anh hoi tinh yeu em danh cho anh nhieu nhu the nao
        'Em nguoc len troi va noi: 
        'Tinh yeu em danh cho anh nhieu nhu nhung vi sao tren troi
        'Va bay gio thi em da di that xa
        'Em bo lai anh voi nhung vi sao kia
        'Voi ky niem ngay nao va voi noi nho ve em

        'Dem nay anh ngoi day nhin troi sao anh nho em
        'Nho den nhung luc xua ta hay ngoi
        'Dat ban tay em trong tay anh
        'Nguyen tinh yeu ta mai xanh
        'Sao tren cao chung nhan em va anh

        'Ngoi sao anh le loi nhin ve noi phuong xa co em
        'Co chut anh sang noi xa chan troi
        'Gio ban tay em trong tay ai
        'Loi the xua nay da phai
        'Sao tren cao khoc cho duyen tan mau

        'Du troi mua hay bao to
        'Van co anh noi day mai nho
        'Van mong anh sao soi buoc em quay ve
        'Ve ben anh nhu luc xua
        'Du ngan nam cho mong hoa da thoi gian

        'Ngay nguoi di mua phu loi
        'Tham uot noi tim anh moi toi
        'Uoc mong nang mai len xoa tan dem dai
        'Loi the xua nhu van day
        'Nghin trung may hoi em nho chang noi nay

    End Sub

    Private Sub GetAllParameter()
        PARA_Server = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "ServerName", "", CodeOption.lmCode)
        PARA_Database = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "DBName", "", CodeOption.lmCode)
        PARA_ConnectionUser = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "ConnectionUserID", "", CodeOption.lmCode)
        PARA_UserID = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "UserID", "", CodeOption.lmCode)
        PARA_Password = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Password", "", CodeOption.lmCode)
        PARA_DivisionID = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "DivisionID", gsDivisionID)
        PARA_TranMonth = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "TranMonth", giTranMonth.ToString)
        PARA_TranYear = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "TranYear", giTranYear.ToString)
        PARA_Language = CType(D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Language", "0"), EnumLanguage)
        PARA_FormID = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", "")
        PARA_FormIDPermission = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Ctrl03", "")
        '-----------------------------------------------------------------------
        gbUnicode = CType(D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "CodeTable", "False"), Boolean)
        '-----------------------------------------------------------------------
        AssignToPublicVariable()

    End Sub

    Private Sub AssignToPublicVariable()
        gsServer = PARA_Server
        gsCompanyID = PARA_Database
        gsConnectionUser = PARA_ConnectionUser
        gsUserID = PARA_UserID
        gsPassword = PARA_Password
        gsDivisionID = PARA_DivisionID
        giTranMonth = CInt(PARA_TranMonth)
        giTranYear = CInt(PARA_TranYear)
        geLanguage = PARA_Language
        gsLanguage = IIf(geLanguage = EnumLanguage.Vietnamese, "84", "01").ToString
        D99C0008.Language = geLanguage
        PARA_FormID = PARA_FormID
        PARA_FormIDPermission = PARA_FormIDPermission
        '-----------------------------------------------------------------------        
    End Sub

    Private Sub MakeVirtualConnection()
        gsUserID = "LEMONADMIN"
        gsConnectionUser = "sa"
        gsPassword = ""
        gsServer = "drd247"
        gsCompanyID = "DRD02"


        'gsPassword = "123"
        'gsServer = "DRD40"
        'gsCompanyID = "TRANG"
    End Sub

    Private Sub SaveParameter()
        Dim sFormID As String = "D02F0200"
        'sFormID = "D02F1030"
        'sFormID = "D02F2000"
        sFormID = "D02F1021"

        Dim sDivisionID As String = "HANOI"
        Dim sTranMonth As String = "09"
        Dim sTranYear As String = "2011"

        'sDivisionID = "THHA"
        'sTranMonth = "05"
        'sTranYear = "2011"

        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ServerName", gsServer, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DBName", gsCompanyID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ConnectionUserID", gsConnectionUser, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "UserID", gsUserID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Password", gsPassword, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DivisionID", sDivisionID)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranMonth", sTranMonth)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranYear", sTranYear)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Language", "0")
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", sFormID) 'PARA_FormID
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl03", sFormID) 'PARA_FormIDPermission
    End Sub

End Module
