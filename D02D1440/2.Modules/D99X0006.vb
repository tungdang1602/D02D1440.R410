'#######################################################################################
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 13/08/2013
'# Diễn giải: Các hàm liên quan đến sinh IGE cho Khóa chính (theo kiểu mới)
'# Sửa lại ghi LogFile cho sinh số phiếu
'# Kiểm tra chuỗi con Key1, Key2 phải nằm trong chuỗi Key3
'# Sửa lỗi hàm CreateIGENewS
'# Người cập nhật cuối cùng: Nguyễn Thị Minh Hòa
'#######################################################################################

Module D99X0006
#Region "Tạo IGE cho khóa chính của bảng khi lưu dữ liệu, không lấy trong DLL"

    ''' <summary>
    ''' Sinh IGE cho khóa chính Master 
    ''' </summary>
    ''' <param name="Table">Bảng sinh khóa</param>
    ''' <param name="Field">Tên trường khóa chính</param>
    ''' <param name="Key1">Giá trị chuỗi 1</param>
    ''' <param name="Key2">Giá trị chuỗi 2</param>
    ''' <param name="Key3">Giá trị chuỗi 3</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function CreateIGE(ByVal Table As String, ByVal Field As String, ByVal Key1 As String, ByVal Key2 As String, ByVal Key3 As String) As String
        'Hàm mới
        Dim ret As String = ""
        Dim lLastKey As Long = 0
        ret = IGEKeyPrimary(Table, Field, Key1, Key2, Key3, lLastKey, 1, OutOrderEnum.lmSSSN, 15, "")
        Return ret

    End Function

    ''' <summary>
    ''' Sinh IGE cho khóa chính Master, có truyền chiều dài và thứ tự hiển thị
    ''' </summary>
    ''' <param name="Table">Bảng sinh khóa</param>
    ''' <param name="Field">Tên trường khóa chính</param>
    ''' <param name="Key1">Giá trị chuỗi 1</param>
    ''' <param name="Key2">Giá trị chuỗi 2</param>
    ''' <param name="Key3">Giá trị chuỗi 3</param>
    ''' <param name="OutOrder"></param>
    ''' <param name="Length"></param>
    ''' <param name="Seperator"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function CreateIGE(ByVal Table As String, ByVal Field As String, ByVal Key1 As String, ByVal Key2 As String, ByVal Key3 As String, ByVal OutOrder As D99D0041.OutOrderEnum, ByVal Length As Integer, ByVal Seperator As String) As String
        Dim ret As String = ""
        Dim iLastKey As Long = 0
        ret = IGEKeyPrimary(Table, Field, Key1, Key2, Key3, iLastKey, 1, OutOrder, Length, Seperator.Trim)
        Return ret
    End Function

    ''' <summary>
    ''' Sinh IGE cho khóa chính Detail
    ''' </summary>
    ''' <param name="Table">Bảng sinh khóa</param>
    ''' <param name="Field">Tên trường khóa chính</param>
    ''' <param name="Key1">Giá trị chuỗi 1</param>
    ''' <param name="Key2">Giá trị chuỗi 2</param>
    ''' <param name="Key3">Giá trị chuỗi 3</param>
    ''' <param name="OldIGE">IGE lần đầu sinh</param>
    ''' <param name="NumberIGE">Số dòng cần sinh IGE </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function CreateIGEs(ByVal Table As String, ByVal Field As String, ByVal Key1 As String, ByVal Key2 As String, ByVal Key3 As String, ByVal OldIGE As String, ByVal NumberIGE As Long) As String
        Dim ret As String = ""
        'Update 12/06/2013: Minh Hòa update Sua loi cho TH sinh IGE kieu moi va cu dung chung. Kiem tra TH chieu dai chuoi vuot qua gioi han thi End
        Dim iLength As Integer = Key1.Length + Key2.Length + Key3.Length
        If iLength > 6 Then
            D99C0008.MsgL3("Error: Length of string." & " Exit module.", L3MessageBoxIcon.Err)
            End
        End If


        If OldIGE = "" Then
            Dim iLastKey As Long = 0
            ret = IGEKeyPrimary(Table, Field, Key1, Key2, Key3, iLastKey, NumberIGE, OutOrderEnum.lmSSSN, 15, "", True)
            Return ret
        Else
            Dim iNo As Long = CLng(OldIGE.Substring(iLength)) + 1
            ret = Key1 & Key2 & Key3 & iNo.ToString(Strings.StrDup(15 - iLength, "0"))
            If ret.Length > 15 Then
                D99C0008.MsgL3("Error IGE of " & Table & " (Length)." & " Exit module.", L3MessageBoxIcon.Err)
                WriteLogFile("Loi sinh IGE (chieu dai qua gioi han) cua table " & Table, "LogIGE.log")
                End
            End If
            Return ret
        End If

    End Function

    ''' <summary>
    ''' Sinh IGE cho khóa chính Detail(Hàm mới)
    ''' </summary>
    ''' <param name="Table">Bảng sinh khóa</param>
    ''' <param name="Field">Tên trường khóa chính</param>
    ''' <param name="Key1">Giá trị chuỗi 1</param>
    ''' <param name="Key2">Giá trị chuỗi 2</param>
    ''' <param name="Key3">Giá trị chuỗi 3</param>
    ''' <param name="OldIGE">IGE lần đầu sinh</param>
    ''' <param name="NumberIGE">Số dòng cần sinh IGE </param>
    ''' <param name="iFirstIGE"> Lastkey của dòng sinh IGE đầu tiên </param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function CreateIGENewS(ByVal Table As String, ByVal Field As String, ByVal Key1 As String, ByVal Key2 As String, ByRef Key3 As String, ByVal OldIGE As String, ByVal NumberIGE As Long, ByRef iFirstIGE As Long) As String

        Dim ret As String = ""
        If OldIGE = "" Then
            ret = IGEKeyPrimary(Table, Field, Key1, Key2, Key3, iFirstIGE, NumberIGE, OutOrderEnum.lmSSSN, 15, "")
            Return ret
        Else
            'Update 13/08/2013
            Dim sKeystring As String = Key1 & Key2 & Key3 'Key3
            Dim iLength As Integer = sKeystring.Length
            Dim iNo As Long = CLng(OldIGE.Substring(iLength)) + 1
            'Update 16/12/2009: Kiểm tra Số dòng sinh tăng lên không > số dòng truyền vào để sinh IGE
            If iNo > iFirstIGE + NumberIGE - 1 Then
                D99C0008.MsgL3("Lỗi sinh mã tự động cho khóa chính (Detail) của " & Table & " (Số tăng tự động)." & " Kết thúc chương trình.", L3MessageBoxIcon.Err)
                WriteLogFile("Loi sinh IGE (Detail) cua table " & Table & vbCrLf & "So dong tang len = " & iNo & vbCrLf & "So dong truyen vao = " & iFirstIGE + NumberIGE - 1, "LogIGEDetailNew.log")
                End
            End If

            ret = sKeystring & iNo.ToString(Strings.StrDup(15 - iLength, "0"))
            If ret.Length > 15 Then
                D99C0008.MsgL3("Lỗi sinh mã tự động cho khóa chính (Detail) của " & Table & " (Chiều dài vược quá giới hạn)." & " Kết thúc chương trình.", L3MessageBoxIcon.Err)
                WriteLogFile("Loi sinh IGE (chieu dai qua gioi han) cua table " & Table & vbCrLf & "Chieu dai tang len = " & ret.Length & ">15", "LogIGEDetailNew.log")
                End
            End If

            'Dòng cuối cùng sinh IGE được ghi vào LogFile để khi test biết được sinh IGE theo kiểu mới (17/12/2009)
            If iNo = iFirstIGE + NumberIGE - 1 Then
                WriteLogFile("Sinh IGE kieu moi cua table " & Table & " (Thanh cong)", "LogIGEDetailNew.log")
            End If

            Return ret
        End If

    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9119
    '# Created User: Nguyễn Thị Minh Hòa
    '# Created Date: 03/11/2009 04:08:21
    '# Modified User: Nguyễn Thị Minh Hòa
    '# Modified Date:  03/11/2009 04:08:21
    '# Description: Store sinh IGE cho khóa chính
    '#---------------------------------------------------------------------------------------------------
    <DebuggerStepThrough()> _
    Private Function SQLStoreD91P9119(ByVal sTableName As String, ByVal sStringKey As String, ByVal iCountIGE As Long) As String
        Dim sSQL As String = ""
        sSQL &= "SET NOCOUNT ON " & vbCrLf
        sSQL &= "DECLARE @KeyString AS VARCHAR(20), " & vbCrLf
        sSQL &= "@KeyFrom AS INT, " & vbCrLf
        sSQL &= "@KeyTo AS INT" & vbCrLf
        sSQL &= "Exec D91P9119 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(sStringKey) & COMMA 'IGEID, varchar[20], NOT NULL
        sSQL &= SQLString(sTableName) & COMMA 'TableName, varchar[20], NOT NULL
        sSQL &= SQLNumber(iCountIGE) & COMMA 'TotalKeys, int, NOT NULL
        sSQL &= " @KeyString  OUTPUT " & COMMA 'KeyString, varchar[20], NOT NULL
        sSQL &= "@KeyFrom  OUTPUT " & COMMA 'KeyFrom, int, NOT NULL
        sSQL &= "@KeyTo OUTPUT " & vbCrLf 'KeyTo, int, NOT NULL
        sSQL &= "SELECT @KeyString AS KeyString, @KeyFrom AS KeyFrom "
        Return sSQL
    End Function

    Private Function IGEKeyPrimary(ByVal sTableName As String, _
                                                            ByVal sFieldID As String, _
                                                            ByVal sStringKey1 As String, _
                                                            ByVal sStringKey2 As String, _
                                                            ByRef sStringKey3 As String, _
                                                            ByRef nOutLastKey As Long, _
                                                            Optional ByVal nRowIGE As Long = 1, _
                                                            Optional ByVal nOutputOrder As OutOrderEnum = OutOrderEnum.lmSSSN, _
                                                            Optional ByVal nOutputLength As Integer = 15, _
                                                            Optional ByVal sSeperatorCharacter As String = "", _
                                                            Optional ByVal bUseOldIGE As Boolean = False) As String

        'Chú ý: Chuẩn sinh IGE luôn luôn có dạng SSSN
        Dim ret As String = ""

        Try
            If bUseOldIGE = False Then 'Sinh IGE kiểu mới
                '**********************************************************
                'Update 03/11/2009: viết theo cơ chế mới dùng Store
                Dim sStringKey As String
                sStringKey = sStringKey1.Trim & sStringKey2.Trim
                Dim sSQL As String = SQLStoreD91P9119(sTableName, sStringKey, nRowIGE)
                Dim dtKey As DataTable
                dtKey = ReturnDataTable(sSQL)
                If dtKey.Rows.Count > 0 Then
                    Dim sKeyString As String = dtKey.Rows(0).Item("KeyString").ToString
                    'Update 13/08/2013: Minh Hòa update dùng biến sStringKey3 để trả về chuỗi SQL sinh IGE cho lần sau
                    sStringKey3 = sKeyString.Substring(sStringKey1.Length + sStringKey2.Length)
                    Dim iKeyFrom As Long = CLng(dtKey.Rows(0).Item("KeyFrom"))
                    nOutLastKey = iKeyFrom
                    ret = sKeyString & iKeyFrom.ToString(Strings.StrDup(15 - sKeyString.Length, "0"))
                End If

                If ret = "" Then
                    D99C0008.MsgL3("Lỗi sinh mã tự động cho khóa chính của " & sTableName & "." & " Kết thúc chương trình.", L3MessageBoxIcon.Err)
                    WriteLogFile("Loi sinh IGE cua table " & sTableName & vbCrLf & sSQL, "LogIGENew.log")
                    End
                Else
                    If ret.Length > nOutputLength Then
                        D99C0008.MsgL3("Lỗi sinh mã tự động (chiều dài vượt quá giới hạn) cho khóa chính của " & sTableName & " (Length = " & ret.Length & ")." & " Kết thúc chương trình.", L3MessageBoxIcon.Err)
                        WriteLogFile("Loi sinh IGE chieu dai qua gioi han cua table " & sTableName & " (Length = " & ret.Length & ")" & vbCrLf & sSQL, "LogIGENew.log")
                        End
                    End If
                End If
                Return ret

                '**********************************************************
            Else 'Sinh IGE kiểu cũ
                Dim bKey As Boolean
                Dim KeyString As String
                Dim sIGEKeyPrimary As String = ""

                bKey = False
                KeyString = sStringKey1 & sStringKey2 & sStringKey3
                Dim LastKey As Long
                Do
                    'Lấy LastKey
                    LastKey = GetLastKey(KeyString, sTableName)
                    '-----------------------------------------------------------
                    'Kiem tra chieu dai và lấy chuỗi string của Lastkey
                    Dim LastKeyString As String
                    LastKeyString = CheckLengthKey(LastKey, sStringKey1, sStringKey2, sStringKey3, sSeperatorCharacter, nOutputLength)
                    If LastKeyString <> "" Then
                        'Hop le thi sinh IGE
                        sIGEKeyPrimary = Generate(sStringKey1, sStringKey2, sStringKey3, nOutputOrder, sSeperatorCharacter, LastKeyString)
                    End If

                    If sIGEKeyPrimary = "" Then
                        If LastKeyString <> "" Then
                            D99C0008.MsgL3("Lỗi sinh mã tự động cho khóa chính của " & sTableName & "." & " Kết thúc chương trình.", L3MessageBoxIcon.Err)
                            WriteLogFile("Loi sinh IGE cua table " & sTableName & vbCrLf & "KeyString = " & KeyString & "LastKey = " & LastKeyString, "LogIGEOld.log")
                        Else
                            D99C0008.MsgL3("Lỗi sinh mã tự động (chiều dài vượt quá giới hạn) cho khóa chính của " & sTableName & ". Kết thúc chương trình.", L3MessageBoxIcon.Err)
                            WriteLogFile("Loi sinh IGE chieu dai qua gioi han cua table " & sTableName & vbCrLf & "KeyString = " & KeyString & "LastKey = " & LastKeyString, "LogIGEOld.log")
                        End If

                        End
                    End If

                    'Luu Last key
                    SaveLastKey(sTableName, KeyString, LastKey - 1 + nRowIGE)

                    'Kiem tra trung khoa
                    Dim sKeyFrom As String, sKeyTo As String
                    Dim intZeroLen As Integer
                    Dim StringLastKey As String
                    Dim nNewLastKey As Long

                    sKeyFrom = sIGEKeyPrimary
                    If nRowIGE = 1 Then
                        sKeyTo = sKeyFrom
                    Else
                        nNewLastKey = (LastKey - 1) + nRowIGE
                        intZeroLen = CType(nOutputLength, Integer) - nNewLastKey.ToString.Length - (sStringKey1.Length + sStringKey2.Length + sStringKey3.Length)
                        '----------------------------
                        If sSeperatorCharacter <> "" Then
                            If sStringKey1 <> "" Then intZeroLen = intZeroLen - 1
                            If sStringKey2 <> "" Then intZeroLen = intZeroLen - 1
                            If sStringKey3 <> "" Then intZeroLen = intZeroLen - 1
                        End If

                        If intZeroLen < 0 Then
                            AnnouncementLength()
                            Return ""
                        Else
                            StringLastKey = Strings.StrDup(intZeroLen, "0") & nNewLastKey
                        End If
                        '----------------------------
                        Select Case nOutputOrder
                            Case OutOrderEnum.lmNSSS
                                sKeyTo = StringLastKey & sStringKey1 & sStringKey2 & sStringKey3
                            Case OutOrderEnum.lmSNSS
                                sKeyTo = sStringKey1 & StringLastKey & sStringKey2 & sStringKey3
                            Case OutOrderEnum.lmSSNS
                                sKeyTo = sStringKey1 & sStringKey2 & StringLastKey & sStringKey3
                            Case Else
                                sKeyTo = sStringKey1 & sStringKey2 & sStringKey3 & StringLastKey
                        End Select
                    End If

                    bKey = CheckDupKeyPrimary(sTableName, sFieldID, sKeyFrom, sKeyTo)

                    'Hop le thi lay du lieu va thoat
                    If Not bKey Then
                        nOutLastKey = LastKey
                        Return sIGEKeyPrimary
                    End If

                Loop Until bKey = False

                'Lỗi sinh IGE
                D99C0008.MsgL3("Lỗi sinh mã tự động cho khóa chính của " & sTableName & " (End)." & " Kết thúc chương trình.", L3MessageBoxIcon.Err)
                WriteLogFile("Loi sinh IGE (End) cua table " & sTableName & vbCrLf & "KeyString = " & KeyString, "LogIGEOld.log")

                End

            End If
        Catch ex As Exception
            MessageBox.Show("Error IGE: " & vbCrLf & ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            WriteLogFile("Loi sinh IGE (End)", "LogIGE.log")
            End
        End Try

    End Function

#End Region

End Module
