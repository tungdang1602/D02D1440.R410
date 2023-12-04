'#######################################################################################
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 18/10/2010 
'# Diễn giải: Các hàm liên quan đến sinh IGE cho Khóa chính (theo kiểu cũ)
'# Người cập nhật cuối cùng: Nguyễn Thị Minh Hòa
'#######################################################################################

Module D99X0016

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
        Dim lLastKey As Long = 0
        ret = IGEKeyPrimary(Table, Field, Key1, Key2, Key3, lLastKey, 1, OutOrder, Length, Seperator.Trim)
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
    ''' <param name="NumberIGE">Số dòng cần sinh IGE</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function CreateIGEs(ByVal Table As String, ByVal Field As String, ByVal Key1 As String, ByVal Key2 As String, ByVal Key3 As String, ByVal OldIGE As String, ByVal NumberIGE As Long) As String
        Dim ret As String = ""
        If OldIGE = "" Then
            Dim lLastKey As Long = 0
            ret = IGEKeyPrimary(Table, Field, Key1, Key2, Key3, lLastKey, NumberIGE, OutOrderEnum.lmSSSN, 15, "")
            Return ret

        Else
            Dim iLength As Integer = Key1.Length + Key2.Length + Key3.Length
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

    <DebuggerStepThrough()> _
    Private Function IGEKeyPrimary(ByVal sTableName As String, _
                                                            ByVal sFieldID As String, _
                                                            ByVal sStringKey1 As String, _
                                                            ByVal sStringKey2 As String, _
                                                            ByVal sStringKey3 As String, _
                                                            ByRef nOutLastKey As Long, _
                                                            Optional ByVal nRowIGE As Long = 1, _
                                                            Optional ByVal nOutputOrder As OutOrderEnum = OutOrderEnum.lmSSSN, _
                                                            Optional ByVal nOutputLength As Integer = 15, _
                                                            Optional ByVal sSeperatorCharacter As String = "") As String


        Dim sIGEKeyPrimary As String = ""

        Try
            Dim bKey As Boolean
            Dim KeyString As String

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
                        D99C0008.MsgL3("Lỗi sinh mã tự động cho khóa chính", L3MessageBoxIcon.Err)
                        WriteLogFile("Loi sinh IGE cua table " & sTableName)
                    Else
                        WriteLogFile("Loi sinh IGE (Chieu dai qua gioi han) cua table " & sTableName)
                    End If
                    Return ""
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

            Return ""
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return ""
        End Try

    End Function

#End Region

End Module
