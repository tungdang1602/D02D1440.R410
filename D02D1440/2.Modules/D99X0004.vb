'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 23/09/2013
'# Người cập nhật cuối cùng: Thị Ánh
'# Diễn giải: 
'# Bổ sung hàm ConvertFontUnicode và Gán font 1 lần cho control của hàm InputbyUnicode và bỏ hàm AdjustFontChildControl(): 25/07/2012
'# Sửa lỗi hàm UnicodeConvertFont() khi có 2 cột có DataField ="" 02/10/2012
'# Bổ sung thêm biến gbPrintVNI, sửa lại hàm GetCodeTable 01/03/2013
'# Sửa lỗi Font cho C1Combo trong hàm UnicodeConvertFont vì khi design có thể là font Microsoft Sans Serif
'# Gán font 1 lần cho control của hàm ConvertFontUnicode 25/06/2013
'# Gán font cho Groupby của hàm ConvertFontUnicode 27/06/2013
'# Fix lỗi font hàm ConvertFontUnicode 09/07/2013
'# Xét font tất cả các cột trong ConvertFontTDBGrid 09/08/2013
'# Kiểm tra font split 0 trong ConvertFontTDBGrid: TH lưới groupby có nhiều split 16/08/2013
'# Bổ sung WITH (NOLOCK) vào table, trong bảng D91T0000 23/9/2013
'#######################################################################################


''' <summary>
''' Module quản lý các vấn đề về Convert Font
''' </summary>
''' <remarks></remarks>
Module D99X0004
    Public gbUnicode As Boolean = False
    Public Const sUnicodeFontName As String = "Microsoft Sans Serif"
    Public gbPrintVNI As Boolean = False

#Region "Các hàm liên quan đến ConvertFont"

    ''' <summary>
    ''' Convert chuỗi từ Unicode sang VietwareF
    ''' </summary>
    ''' <param name="sText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvertUnicodeToVietwareF(ByVal sText As String) As String
        Dim sRet As String = sText

        sRet = sRet.Replace("Ó", "[O1]")
        sRet = sRet.Replace("Ò", CStr(ChrW(223)))

        sRet = sRet.Replace("Ỏ", "[O3]")
        sRet = sRet.Replace("Õ", "[O4]")
        sRet = sRet.Replace("Ọ", "[O5]")
        sRet = sRet.Replace("Ô", CStr(ChrW(353)))
        sRet = sRet.Replace("Ố", CStr(ChrW(231)))
        sRet = sRet.Replace("Ồ", CStr(ChrW(228)))
        sRet = sRet.Replace("Ổ", CStr(ChrW(229)))
        sRet = sRet.Replace("Ỗ", CStr(ChrW(230)))

        sRet = sRet.Replace("Ộ", "[O65]")
        sRet = sRet.Replace("Ơ", CStr(ChrW(8250)))

        sRet = sRet.Replace("Ớ", "[O71]")
        sRet = sRet.Replace("Ờ", "[O72]")
        sRet = sRet.Replace("Ở", "[O73]")
        sRet = sRet.Replace("Ỡ", "[O74]")
        sRet = sRet.Replace("Ợ", "[O75]")

        sRet = sRet.Replace("Í", "[I1]") 'CStr(ChrW(219)))
        sRet = sRet.Replace("Ì", "[I2]") 'CStr(ChrW(216)))
        sRet = sRet.Replace("Ỉ", "[I3]") 'CStr(ChrW(217)))
        sRet = sRet.Replace("Ĩ", "[I4]") 'CStr(ChrW(218)))
        sRet = sRet.Replace("Ị", "[I5]") 'CStr(ChrW(220)))

        sRet = sRet.Replace("É", CStr(ChrW(207)))
        sRet = sRet.Replace("È", CStr(ChrW(204)))
        sRet = sRet.Replace("Ẻ", CStr(ChrW(205)))
        sRet = sRet.Replace("Ẽ", CStr(ChrW(206)))
        sRet = sRet.Replace("Ẹ", CStr(ChrW(209)))
        sRet = sRet.Replace("Ê", CStr(ChrW(8482)))
        sRet = sRet.Replace("Ế", CStr(ChrW(213)))
        sRet = sRet.Replace("Ề", CStr(ChrW(210)))
        sRet = sRet.Replace("Ể", CStr(ChrW(211)))
        sRet = sRet.Replace("Ễ", CStr(ChrW(212)))
        sRet = sRet.Replace("Ệ", CStr(ChrW(214)))

        sRet = sRet.Replace("Ằ", "[A82]")
        sRet = sRet.Replace("À", CStr(ChrW(170)))
        sRet = sRet.Replace("Á", CStr(ChrW(192)))
        sRet = sRet.Replace("Ả", CStr(ChrW(182)))
        sRet = sRet.Replace("Ã", CStr(ChrW(186)))
        sRet = sRet.Replace("Ạ", CStr(ChrW(193)))
        sRet = sRet.Replace("Ă", CStr(ChrW(8211)))
        sRet = sRet.Replace("Ắ", CStr(ChrW(197)))
        sRet = sRet.Replace("Ẳ", CStr(ChrW(195)))
        sRet = sRet.Replace("Ẵ", CStr(ChrW(196)))
        sRet = sRet.Replace("Ặ", CStr(ChrW(198)))
        sRet = sRet.Replace("Â", CStr(ChrW(8212)))
        sRet = sRet.Replace("Ấ", CStr(ChrW(202)))
        sRet = sRet.Replace("Ầ", CStr(ChrW(199)))
        sRet = sRet.Replace("Ẩ", CStr(ChrW(200)))
        sRet = sRet.Replace("Ẫ", CStr(ChrW(201)))
        sRet = sRet.Replace("Ậ", CStr(ChrW(203)))

        sRet = sRet.Replace("Ú", "[U1]") 'CStr(ChrW(242)))
        sRet = sRet.Replace("Ù", "[U2]") 'CStr(ChrW(238)))
        sRet = sRet.Replace("Ủ", "[U3]") 'CStr(ChrW(239)))
        sRet = sRet.Replace("Ũ", "[U4]") 'CStr(ChrW(241)))
        sRet = sRet.Replace("Ụ", "[U5]") 'CStr(ChrW(243)))
        sRet = sRet.Replace("Ư", "[U7]") 'CStr(ChrW(339)))
        sRet = sRet.Replace("Ứ", "[U71]") 'CStr(ChrW(247)))
        sRet = sRet.Replace("Ừ", "[U72]") 'CStr(ChrW(244)))
        sRet = sRet.Replace("Ử", "[U73]") 'CStr(ChrW(245)))
        sRet = sRet.Replace("Ữ", "[U74]") 'CStr(ChrW(246)))
        sRet = sRet.Replace("Ự", "[U75]") 'CStr(ChrW(248)))

        sRet = sRet.Replace("Ý", "[Y1]") ' CStr(ChrW(252)))
        sRet = sRet.Replace("Ỳ", "[Y2]") 'CStr(ChrW(249)))
        sRet = sRet.Replace("Ỷ", "[Y3]") 'CStr(ChrW(250)))
        sRet = sRet.Replace("Ỹ", "[Y4]") 'CStr(ChrW(251)))
        sRet = sRet.Replace("Ỵ", "[Y5]") 'CStr(ChrW(255)))

        sRet = sRet.Replace("Đ", CStr(ChrW(732)))

        sRet = sRet.Replace("á", CStr(ChrW(192)))
        sRet = sRet.Replace("à", CStr(ChrW(170)))
        sRet = sRet.Replace("ả", CStr(ChrW(182)))
        sRet = sRet.Replace("ã", CStr(ChrW(186)))
        sRet = sRet.Replace("ạ", CStr(ChrW(193)))
        sRet = sRet.Replace("ă", CStr(ChrW(376)))
        sRet = sRet.Replace("ắ", CStr(ChrW(197)))
        sRet = sRet.Replace("ằ", CStr(ChrW(194)))
        sRet = sRet.Replace("ẳ", CStr(ChrW(195)))
        sRet = sRet.Replace("ẵ", CStr(ChrW(196)))
        sRet = sRet.Replace("ặ", CStr(ChrW(198)))
        sRet = sRet.Replace("â", CStr(ChrW(161)))
        sRet = sRet.Replace("ấ", CStr(ChrW(202)))
        sRet = sRet.Replace("ầ", CStr(ChrW(199)))
        sRet = sRet.Replace("ẩ", CStr(ChrW(200)))
        sRet = sRet.Replace("ẫ", CStr(ChrW(201)))
        sRet = sRet.Replace("ậ", CStr(ChrW(203)))

        sRet = sRet.Replace("é", CStr(ChrW(207)))
        sRet = sRet.Replace("è", CStr(ChrW(204)))
        sRet = sRet.Replace("ẻ", CStr(ChrW(205)))
        sRet = sRet.Replace("ẽ", CStr(ChrW(206)))
        sRet = sRet.Replace("ẹ", CStr(ChrW(209)))
        sRet = sRet.Replace("ê", CStr(ChrW(163)))
        sRet = sRet.Replace("ế", CStr(ChrW(213)))
        sRet = sRet.Replace("ề", CStr(ChrW(210)))
        sRet = sRet.Replace("ể", CStr(ChrW(211)))
        sRet = sRet.Replace("ễ", CStr(ChrW(212)))
        sRet = sRet.Replace("ệ", CStr(ChrW(214)))

        sRet = sRet.Replace("í", CStr(ChrW(219)))
        sRet = sRet.Replace("ì", CStr(ChrW(216)))
        sRet = sRet.Replace("ỉ", CStr(ChrW(217)))
        sRet = sRet.Replace("ĩ", CStr(ChrW(218)))
        sRet = sRet.Replace("ị", CStr(ChrW(220)))

        sRet = sRet.Replace("ó", CStr(ChrW(226)))
        sRet = sRet.Replace("ò", CStr(ChrW(223)))
        sRet = sRet.Replace("ỏ", CStr(ChrW(224)))
        sRet = sRet.Replace("õ", CStr(ChrW(225)))
        sRet = sRet.Replace("ọ", CStr(ChrW(227)))
        sRet = sRet.Replace("ô", CStr(ChrW(164)))
        sRet = sRet.Replace("ố", CStr(ChrW(231)))
        sRet = sRet.Replace("ồ", CStr(ChrW(228)))
        sRet = sRet.Replace("ổ", CStr(ChrW(229)))
        sRet = sRet.Replace("ỗ", CStr(ChrW(230)))
        sRet = sRet.Replace("ộ", CStr(ChrW(232)))
        sRet = sRet.Replace("ơ", CStr(ChrW(165)))
        sRet = sRet.Replace("ớ", CStr(ChrW(236)))
        sRet = sRet.Replace("ờ", CStr(ChrW(233)))
        sRet = sRet.Replace("ở", CStr(ChrW(234)))
        sRet = sRet.Replace("ỡ", CStr(ChrW(235)))
        sRet = sRet.Replace("ợ", CStr(ChrW(237)))

        sRet = sRet.Replace("ú", CStr(ChrW(242)))
        sRet = sRet.Replace("ù", CStr(ChrW(238)))
        sRet = sRet.Replace("ủ", CStr(ChrW(239)))
        sRet = sRet.Replace("ũ", CStr(ChrW(241)))
        sRet = sRet.Replace("ụ", CStr(ChrW(243)))
        sRet = sRet.Replace("ư", CStr(ChrW(167)))
        sRet = sRet.Replace("ứ", CStr(ChrW(247)))
        sRet = sRet.Replace("ừ", CStr(ChrW(244)))
        sRet = sRet.Replace("ử", CStr(ChrW(245)))
        sRet = sRet.Replace("ữ", CStr(ChrW(246)))
        sRet = sRet.Replace("ự", CStr(ChrW(248)))

        sRet = sRet.Replace("ý", CStr(ChrW(252)))
        sRet = sRet.Replace("ỳ", CStr(ChrW(249)))
        sRet = sRet.Replace("ỷ", CStr(ChrW(250)))
        sRet = sRet.Replace("ỹ", CStr(ChrW(251)))
        sRet = sRet.Replace("ỵ", CStr(ChrW(255)))

        sRet = sRet.Replace("đ", CStr(ChrW(162)))

        sRet = sRet.Replace("[A82]", CStr(ChrW(194)))
        sRet = sRet.Replace("[O1]", CStr(ChrW(226)))
        sRet = sRet.Replace("[O3]", CStr(ChrW(224)))
        sRet = sRet.Replace("[O4]", CStr(ChrW(225)))
        sRet = sRet.Replace("[O5]", CStr(ChrW(227)))
        sRet = sRet.Replace("[O65]", CStr(ChrW(232)))
        sRet = sRet.Replace("[O71]", CStr(ChrW(236)))
        sRet = sRet.Replace("[O72]", CStr(ChrW(233)))
        sRet = sRet.Replace("[O73]", CStr(ChrW(234)))
        sRet = sRet.Replace("[O74]", CStr(ChrW(235)))
        sRet = sRet.Replace("[O75]", CStr(ChrW(237)))

        sRet = sRet.Replace("[I1]", CStr(ChrW(219)))
        sRet = sRet.Replace("[I2]", CStr(ChrW(216)))
        sRet = sRet.Replace("[I3]", CStr(ChrW(217)))
        sRet = sRet.Replace("[I4]", CStr(ChrW(218)))
        sRet = sRet.Replace("[I5]", CStr(ChrW(220)))

        sRet = sRet.Replace("[U1]", CStr(ChrW(242)))
        sRet = sRet.Replace("[U2]", CStr(ChrW(238)))
        sRet = sRet.Replace("[U3]", CStr(ChrW(239)))
        sRet = sRet.Replace("[U4]", CStr(ChrW(241)))
        sRet = sRet.Replace("[U5]", CStr(ChrW(243)))
        sRet = sRet.Replace("[U7]", CStr(ChrW(339)))
        sRet = sRet.Replace("[U71]", CStr(ChrW(247)))
        sRet = sRet.Replace("[U72]", CStr(ChrW(244)))
        sRet = sRet.Replace("[U73]", CStr(ChrW(245)))
        sRet = sRet.Replace("[U74]", CStr(ChrW(246)))
        sRet = sRet.Replace("[U75]", CStr(ChrW(248)))
        sRet = sRet.Replace("[Y1]", CStr(ChrW(252)))
        sRet = sRet.Replace("[Y2]", CStr(ChrW(249)))
        sRet = sRet.Replace("[Y3]", CStr(ChrW(250)))
        sRet = sRet.Replace("[Y4]", CStr(ChrW(251)))
        sRet = sRet.Replace("[Y5]", CStr(ChrW(255)))

        Return sRet
    End Function

    ''' <summary>
    ''' Convert chuỗi từ VietwareF (MS Sans Serif, DG Sans Serif) sang Unicode
    ''' </summary>
    ''' <param name="sText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function ConvertVietwareFToUnicode(ByVal sText As String) As String
        Dim sRet As String = sText

        sRet = sRet.Replace(CStr(ChrW(249)), "ỳ")
        sRet = sRet.Replace(CStr(ChrW(252)), "ý")
        sRet = sRet.Replace(CStr(ChrW(255)), "ỵ")
        sRet = sRet.Replace(CStr(ChrW(250)), "ỷ")
        sRet = sRet.Replace(CStr(ChrW(251)), "ỹ")

        sRet = sRet.Replace(CStr(ChrW(238)), "ù")
        sRet = sRet.Replace(CStr(ChrW(242)), "ú")
        sRet = sRet.Replace(CStr(ChrW(243)), "ụ")
        sRet = sRet.Replace(CStr(ChrW(239)), "ủ")
        sRet = sRet.Replace(CStr(ChrW(241)), "ũ")

        sRet = sRet.Replace(CStr(ChrW(167)), "ư")
        sRet = sRet.Replace(CStr(ChrW(244)), "ừ")
        sRet = sRet.Replace(CStr(ChrW(247)), "ứ")
        sRet = sRet.Replace(CStr(ChrW(248)), "ự")
        sRet = sRet.Replace(CStr(ChrW(245)), "ử")
        sRet = sRet.Replace(CStr(ChrW(246)), "ữ")

        sRet = sRet.Replace(CStr(ChrW(223)), "ò")
        sRet = sRet.Replace(CStr(ChrW(226)), "ó")
        sRet = sRet.Replace(CStr(ChrW(227)), "ọ")
        sRet = sRet.Replace(CStr(ChrW(224)), "ỏ")
        sRet = sRet.Replace(CStr(ChrW(225)), "õ")

        sRet = sRet.Replace(CStr(ChrW(228)), "ồ")
        sRet = sRet.Replace(CStr(ChrW(231)), "ố")
        sRet = sRet.Replace(CStr(ChrW(232)), "ộ")
        sRet = sRet.Replace(CStr(ChrW(229)), "ổ")
        sRet = sRet.Replace(CStr(ChrW(230)), "ỗ")

        sRet = sRet.Replace(CStr(ChrW(233)), "ờ")
        sRet = sRet.Replace(CStr(ChrW(236)), "ớ")
        sRet = sRet.Replace(CStr(ChrW(237)), "ợ")
        sRet = sRet.Replace(CStr(ChrW(234)), "ở")
        sRet = sRet.Replace(CStr(ChrW(235)), "ỡ")

        sRet = sRet.Replace(CStr(ChrW(216)), "ì")
        sRet = sRet.Replace(CStr(ChrW(219)), "í")
        sRet = sRet.Replace(CStr(ChrW(220)), "ị")
        sRet = sRet.Replace(CStr(ChrW(217)), "ỉ")
        sRet = sRet.Replace(CStr(ChrW(218)), "ĩ")

        sRet = sRet.Replace(CStr(ChrW(164)), "ô")
        sRet = sRet.Replace(CStr(ChrW(165)), "ơ")

        sRet = sRet.Replace(CStr(ChrW(170)), "à")
        sRet = sRet.Replace(CStr(ChrW(192)), "á")
        sRet = sRet.Replace(CStr(ChrW(193)), "ạ")
        sRet = sRet.Replace(CStr(ChrW(182)), "ả")
        sRet = sRet.Replace(CStr(ChrW(186)), "ã")

        sRet = sRet.Replace(CStr(ChrW(376)), "ă")
        sRet = sRet.Replace(CStr(ChrW(194)), "ằ")
        sRet = sRet.Replace(CStr(ChrW(197)), "ắ")
        sRet = sRet.Replace(CStr(ChrW(198)), "ặ")
        sRet = sRet.Replace(CStr(ChrW(195)), "ẳ")
        sRet = sRet.Replace(CStr(ChrW(196)), "ẵ")

        sRet = sRet.Replace(CStr(ChrW(161)), "â")
        sRet = sRet.Replace(CStr(ChrW(199)), "ầ")
        sRet = sRet.Replace(CStr(ChrW(202)), "ấ")
        sRet = sRet.Replace(CStr(ChrW(203)), "ậ")
        sRet = sRet.Replace(CStr(ChrW(200)), "ẩ")
        sRet = sRet.Replace(CStr(ChrW(201)), "ẫ")

        sRet = sRet.Replace(CStr(ChrW(204)), "è")
        sRet = sRet.Replace(CStr(ChrW(207)), "é")
        sRet = sRet.Replace(CStr(ChrW(209)), "ẹ")
        sRet = sRet.Replace(CStr(ChrW(205)), "ẻ")
        sRet = sRet.Replace(CStr(ChrW(206)), "ẽ")

        sRet = sRet.Replace(CStr(ChrW(163)), "ê")
        sRet = sRet.Replace(CStr(ChrW(210)), "ề")
        sRet = sRet.Replace(CStr(ChrW(213)), "ế")
        sRet = sRet.Replace(CStr(ChrW(214)), "ệ")
        sRet = sRet.Replace(CStr(ChrW(211)), "ể")
        sRet = sRet.Replace(CStr(ChrW(212)), "ễ")

        sRet = sRet.Replace(CStr(ChrW(162)), "đ")

        sRet = sRet.Replace(CStr(ChrW(8211)), "Ă")
        sRet = sRet.Replace(CStr(ChrW(8212)), "Â")
        sRet = sRet.Replace(CStr(ChrW(732)), "Đ")
        sRet = sRet.Replace(CStr(ChrW(8482)), "Ê")
        sRet = sRet.Replace(CStr(ChrW(353)), "Ô")
        sRet = sRet.Replace(CStr(ChrW(8250)), "Ơ")
        sRet = sRet.Replace(CStr(ChrW(339)), "Ư")

        Return sRet
    End Function

    ''' <summary>
    ''' Convert chuỗi từ VNI sang Unicode
    ''' </summary>
    ''' <param name="szInput"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function ConvertVniToUnicode(ByVal szInput As String) As String
        Dim sChar, sNextChar As Char
        Dim sInput, sOutput As String

        If IsDBNull(szInput) Then
            Return ""
        End If

        If szInput = "" Then
            Return ""
        End If

        sInput = szInput.Trim

        If sInput.Length = 0 Then
            Return ""
        End If

        sOutput = ""

        'check the first char
        Do
            sChar = Convert.ToChar(sInput.Substring(0, 1))

            If CheckVNIVowel(sChar) Then
                If Len(sInput) > 1 Then
                    sNextChar = Convert.ToChar(sInput.Substring(1, 1))
                    Dim code As Integer = Asc(sNextChar)
                    If CheckVNIAccent(sNextChar) Then
                        sOutput &= CDoubleVNI(sChar, sNextChar)

                        If sInput.Length > 2 Then
                            sInput = sInput.Substring(2)
                        Else
                            Exit Do
                        End If

                    Else
                        sOutput &= CSingleVNI(sChar)
                        sInput = sInput.Substring(1)
                    End If
                Else
                    sOutput &= CSingleVNI(sChar)
                    Exit Do
                End If
            Else
                'Response.Write "<p>" & sChar & ": not a vowel"
                sOutput &= CSingleVNI(sChar)
                If sInput.Length > 1 Then
                    sInput = sInput.Substring(1)
                Else
                    Exit Do
                End If
            End If
        Loop

        Return sOutput
    End Function

    Function CSingleVNI(ByVal szInput As Char) As String
        'input: single character (vowel and consonant)
        'output Unicode character
        Dim sResult As String
        'Dim sInput As String
        'sInput = szInput

        Select Case Asc(szInput)
            Case 209 'D with dash, 
                sResult = "&#272;"
                sResult = "Đ"

            Case 241 'd with dash
                sResult = "&#273;"
                sResult = "đ"

            Case 212 'O)
                sResult = "&#416;"
                sResult = "Ơ"

            Case 244 'o)
                sResult = "&#417;"
                sResult = "ơ"

            Case 214 'U)
                sResult = "&#431;"
                sResult = "Ư"

            Case 246 'u)		
                sResult = "&#432;"
                sResult = "ư"

            Case 204 'I`	
                sResult = "&Igrave;"
                sResult = "Ì"

            Case 198 'I?
                sResult = "&#7880;"
                sResult = "Ỉ"

            Case 211 'I~
                sResult = "&#296;"
                sResult = "Ĩ"

            Case 210 'I.
                sResult = "&#7882;"
                sResult = "Ị"

            Case 236 'i`
                sResult = "&igrave;"
                sResult = "ì"

            Case 230 'i?
                sResult = "&#7881;"
                sResult = "ỉ"

            Case 243 'i~
                sResult = "&#297;"
                sResult = "ĩ"

            Case 242 'i.
                sResult = "&#7883;"
                sResult = "ị"

            Case 206 'Y.			
                sResult = "&#7924;"
                sResult = "Ỵ"

            Case 238 'y.			
                sResult = "&#7925;"
                sResult = "ỵ"

            Case Else
                sResult = szInput
        End Select

        Return sResult
    End Function

    Function CDoubleVNI(ByVal szVowel As Char, ByVal szAccent As Char) As String
        'input : a vowel and its accent
        'output: Unicode equivalent

        Dim sChar, sNextChar, sOutput As String

        sChar = szVowel
        sNextChar = szAccent
        sOutput = ""

        Select Case Asc(sChar)
            Case 65
                'A
                Select Case Asc(sNextChar)
                    Case 216
                        'grave À
                        sOutput = Chr(192)
                    Case 217
                        'acute Á
                        sOutput = Chr(193)
                    Case 219
                        'question Ả
                        sOutput = "&#7842;"
                        sOutput = "Ả"
                    Case 213
                        'tilde Ã
                        sOutput = Chr(195)

                    Case 207
                        'dot
                        sOutput = "&#7840;"
                        sOutput = "Ạ"
                    Case 194
                        'A^ (shared code 65 vith A)
                        'A^
                        sOutput = Chr(194)
                    Case 192
                        'grave Ầ
                        sOutput = "&#7846;"
                        sOutput = "Ầ"
                    Case 193
                        'acute
                        sOutput = "&#7844;"
                        sOutput = "Ấ"
                    Case 197
                        'question
                        sOutput = "&#7848;"
                        sOutput = "Ẩ"
                    Case 195
                        'tilde
                        sOutput = "&#7850;"
                        sOutput = "Ẫ"
                    Case 196
                        'dot
                        sOutput = "&#7852;"
                        sOutput = "Ậ"
                    Case 202
                        'A( note: sharing code with A
                        sOutput = "&#258;"
                        sOutput = "Ă"
                    Case 200
                        'grave dấu huyền
                        sOutput = "&#7856;"
                        sOutput = "Ằ"

                    Case 201
                        'acute sắc
                        sOutput = "&#7854;"
                        sOutput = "Ắ"

                    Case 218
                        'question
                        sOutput = "&#7858;"
                        sOutput = "Ẳ"

                    Case 220
                        'tilde
                        sOutput = "&#7860;"
                        sOutput = "Ẵ"

                    Case 203
                        'dot
                        sOutput = "&#7862;"
                        sOutput = "Ặ"
                End Select

            Case 97 'a
                Select Case Asc(sNextChar)
                    Case 248
                        'grave
                        sOutput = Chr(224)

                    Case 249
                        'acute
                        sOutput = Chr(225)

                    Case 251
                        'question
                        sOutput = "&#7843;"
                        sOutput = "ả"

                    Case 245
                        'tilde
                        sOutput = "&atilde;"
                        sOutput = "ã"

                    Case 239
                        'dot
                        sOutput = "&#7841;"
                        sOutput = "ạ"

                    Case 226
                        'a^ (sharing code with a)
                        sOutput = Chr(226)
                        sOutput = "â"

                    Case 224
                        'grave
                        sOutput = "&#7847;"
                        sOutput = "ầ"

                    Case 225
                        'acute
                        sOutput = "&#7845;"
                        sOutput = "ấ"

                    Case 229
                        'question
                        sOutput = "&#7849;"
                        sOutput = "ẩ"

                    Case 227
                        'tilde
                        sOutput = "&#7851;"
                        sOutput = "ẫ"

                    Case 228
                        'dot
                        sOutput = "&#7853;"
                        sOutput = "ậ"

                    Case 234
                        'a( note: sharing code with a
                        sOutput = "&#259;"
                        sOutput = "ă"

                    Case 232
                        'grave
                        sOutput = "&#7857;"
                        sOutput = "ằ"

                    Case 233
                        'acute
                        sOutput = "&#7855;"
                        sOutput = "ắ"

                    Case 250
                        'question
                        sOutput = "&#7859;"
                        sOutput = "ẳ"

                    Case 252
                        'tilde
                        sOutput = "&#7861;"
                        sOutput = "ẵ"

                    Case 235
                        'dot
                        sOutput = "&#7863;"
                        sOutput = "ặ"
                End Select

            Case 69 'E
                Select Case Asc(sNextChar)
                    Case 216
                        'grave
                        sOutput = Chr(200)

                    Case 217
                        'acute
                        sOutput = Chr(201)

                    Case 219
                        'question
                        sOutput = "&#7866;"
                        sOutput = "Ẻ"

                    Case 213
                        'tilde
                        sOutput = "&#7868;"
                        sOutput = "Ẽ"

                    Case 207
                        'dot
                        sOutput = "&#7864;"
                        sOutput = "Ẹ"

                    Case 194
                        'E^ note: sharing code with E
                        sOutput = Chr(202)

                    Case 192
                        'grave
                        sOutput = "&#7872;"
                        sOutput = "Ề"

                    Case 193
                        'acute
                        sOutput = "&#7870;"
                        sOutput = "Ế"

                    Case 197
                        'question
                        sOutput = "&#7874;"
                        sOutput = "Ể"

                    Case 195
                        'tilde
                        sOutput = "&#7876;"
                        sOutput = "Ễ"

                    Case 196
                        'dot
                        sOutput = "&#7878;"
                        sOutput = "Ệ"
                End Select

            Case 101 'e
                Select Case Asc(sNextChar)
                    Case 248
                        'grave
                        sOutput = Chr(232)
                        sOutput = "è"

                    Case 249
                        'acute
                        sOutput = Chr(233)
                        sOutput = "é"

                    Case 251
                        'question
                        sOutput = "&#7867;"
                        sOutput = "ẻ"

                    Case 245
                        'tilde
                        sOutput = "&#7869;"
                        sOutput = "ẽ"

                    Case 239
                        'dot
                        sOutput = "&#7865;"
                        sOutput = "ẹ"

                    Case 226
                        'e^ note: sharing code with e
                        sOutput = Chr(234)

                    Case 224
                        'grave
                        sOutput = "&#7873;"
                        sOutput = "ề"

                    Case 225
                        'acute
                        sOutput = "&#7871;"
                        sOutput = "ế"

                    Case 229
                        'question
                        sOutput = "&#7875;"
                        sOutput = "ể"

                    Case 227
                        'tilde
                        sOutput = "&#7877;"
                        sOutput = "ễ"

                    Case 228
                        'dot
                        sOutput = "&#7879;"
                        sOutput = "ệ"
                End Select

                'I(Case 73) not applicable
                'i (Case 105) not applicable

            Case 79 'O
                Select Case Asc(sNextChar)
                    Case 216
                        'grave
                        sOutput = "&Ograve;"
                        sOutput = "Ò"

                    Case 217
                        'acute
                        sOutput = Chr(211)

                    Case 219
                        'question
                        sOutput = "&#7886;"
                        sOutput = "Ỏ"

                    Case 213
                        'tilde
                        sOutput = "&Otilde;"
                        sOutput = "Õ"

                    Case 207
                        'dot
                        sOutput = "&#7884;"
                        sOutput = "Ọ"

                        'O^ note: sharing code with O
                    Case 194
                        sOutput = Chr(212)

                    Case 192
                        'grave
                        sOutput = "&#7890;"
                        sOutput = "Ồ"

                    Case 193
                        'acute
                        sOutput = "&#7888;"
                        sOutput = "Ố"

                    Case 197
                        'question
                        sOutput = "&#7892;"
                        sOutput = "Ổ"

                    Case 195
                        'tilde
                        sOutput = "&#7894;"
                        sOutput = "Ỗ"

                    Case 196
                        'dot
                        sOutput = "&#7896;"
                        sOutput = "Ộ"
                End Select

                'o
            Case 111
                Select Case Asc(sNextChar)
                    Case 248
                        'grave
                        sOutput = "&ograve;"
                        sOutput = "ò"

                    Case 249
                        'acute
                        sOutput = Chr(243)

                    Case 251
                        'question
                        sOutput = "&#7887;"
                        sOutput = "ỏ"

                    Case 245
                        'tilde
                        sOutput = "&otilde;"
                        sOutput = "õ"

                    Case 239
                        'dot
                        sOutput = "&#7885;"
                        sOutput = "ọ"

                        'o^ note: sharing code with o
                    Case 226
                        sOutput = Chr(244)
                        'sOutput = "ộ"

                    Case 224
                        'grave
                        sOutput = "&#7891;"
                        sOutput = "ồ"

                    Case 225
                        'acute
                        sOutput = "&#7889;"
                        sOutput = "ố"

                    Case 229
                        'question
                        sOutput = "&#7893;"
                        sOutput = "ổ"

                    Case 227
                        'tilde
                        sOutput = "&#7895;"
                        sOutput = "ỗ"

                    Case 228
                        'dot
                        sOutput = "&#7897;"
                        sOutput = "ộ"
                End Select

                'ký tự gì vậy nè
                'Ơ
            Case 212
                Select Case Asc(sNextChar)
                    Case 216
                        'grave
                        sOutput = "&#7900;"
                        sOutput = "Ờ"

                    Case 217
                        'acute
                        sOutput = "&#7898;"
                        sOutput = "Ớ"

                    Case 219
                        'question
                        sOutput = "&#7902;"
                        sOutput = "Ở"

                    Case 213
                        'tilde
                        sOutput = "&#7904;"
                        sOutput = "Ỡ"

                    Case 207
                        'dot
                        sOutput = "&#7906;"
                        sOutput = "Ợ"
                End Select

                'ơ
            Case 244
                Select Case Asc(sNextChar)
                    Case 248
                        'grave
                        sOutput = "&#7901;"
                        sOutput = "ờ"

                    Case 249
                        'acute
                        sOutput = "&#7899;"
                        sOutput = "ớ"

                    Case 251
                        'question
                        sOutput = "&#7903;"
                        sOutput = "ở"

                    Case 245
                        'tilde
                        sOutput = "&#7905;"
                        sOutput = "ỡ"

                    Case 239
                        'dot
                        sOutput = "&#7907;"
                        sOutput = "ợ"
                End Select

                'U
            Case 85
                Select Case Asc(sNextChar)
                    Case 216
                        'grave
                        sOutput = Chr(217)
                    Case 217
                        'acute
                        sOutput = Chr(218)

                    Case 219
                        'question
                        sOutput = "&#7910;"
                        sOutput = "Ủ"

                    Case 213
                        'tilde
                        sOutput = "&#360;"
                        sOutput = "Ũ"

                    Case 207
                        'dot
                        sOutput = "&#7908;"
                        sOutput = "Ụ"
                End Select

            Case 117 'u
                Select Case Asc(sNextChar)
                    Case 248
                        'grave
                        sOutput = Chr(249)
                    Case 249
                        'acute
                        sOutput = Chr(250)

                    Case 251
                        'question
                        sOutput = "&#7911;"
                        sOutput = "ủ"

                    Case 245
                        'tilde
                        sOutput = "&#361;"
                        sOutput = "ũ"

                    Case 239
                        'dot
                        sOutput = "&#7909;"
                        sOutput = "ụ"
                End Select

                'U)
            Case 214
                Select Case Asc(sNextChar)
                    Case 216
                        'grave
                        sOutput = "&#7914;"
                        sOutput = "Ừ"

                    Case 217
                        'acute
                        sOutput = "&#7912;"
                        sOutput = "Ứ"

                    Case 219
                        'question
                        sOutput = "&#7916;"
                        sOutput = "Ử"

                    Case 213
                        'tilde
                        sOutput = "&#7918;"
                        sOutput = "Ữ"

                    Case 207
                        'dot
                        sOutput = "&#7920;"
                        sOutput = "Ự"
                End Select

                'u)
            Case 246
                Select Case Asc(sNextChar)
                    Case 248
                        'grave
                        sOutput = "&#7915;"
                        sOutput = "ừ"
                    Case 249
                        'acute
                        sOutput = "&#7913;"
                        sOutput = "ứ"

                    Case 251
                        'question
                        sOutput = "&#7917;"
                        sOutput = "ử"

                    Case 245
                        'tilde
                        sOutput = "&#7919;"
                        sOutput = "ữ"

                    Case 239
                        'dot
                        sOutput = "&#7921;"
                        sOutput = "ự"
                    Case Else
                        sOutput = "&#432;" & sNextChar
                        sOutput = "ư" & sNextChar
                End Select

                'Y
            Case 89
                Select Case Asc(sNextChar)
                    Case 216
                        'grave
                        sOutput = "&#7922;"
                        sOutput = "Ỳ"

                    Case 217
                        'acute
                        sOutput = "&Yacute;"
                        sOutput = "Ý"

                    Case 219
                        'question
                        sOutput = "&#7926;"
                        sOutput = "Ỷ"

                    Case 213
                        'tilde dấu ngã
                        sOutput = "&#7928;"
                        sOutput = "Ỹ"

                        'dot: already handled by CSingleVNI()
                End Select

                'y
            Case 121
                Select Case Asc(sNextChar)
                    Case 248
                        'grave
                        sOutput = "&#7923;"
                        sOutput = "ỳ"

                    Case 249
                        'acute
                        sOutput = "&yacute;"
                        sOutput = "ý"

                    Case 251
                        'question
                        sOutput = "&#7927;"
                        sOutput = "ỷ"

                    Case 245
                        'tilde
                        sOutput = "&#7929;"
                        sOutput = "ỹ"

                        'dot: already handled by CSingleVNI()
                End Select
        End Select

        Return sOutput
    End Function

    Function CheckVNIAccent(ByVal szInput As Char) As Boolean
        'check if an accent character is there
        Select Case Asc(szInput)
            Case 248, 249, 251, 245, 239
                'Small: ` ' ? ~ .
                Return True

            Case 234, 232, 233, 250, 252, 235
                'Small (, (`, (', (?, (~, (.
                Return True

            Case 226, 224, 225, 229, 227, 228
                'Small ^, ^`, ^', ^?, ^~, ^.	
                Return True

            Case 216, 217, 219, 213, 207
                'Big ` ' ? ~ .
                Return True

            Case 202, 200, 201, 218, 220, 203
                'Big (, (`, (', (?, (~, (.
                Return True

            Case 194, 192, 193, 197, 195, 196
                'Big ^, ^`, ^', ^?, ^~, ^.	
                Return True

            Case Else
                Return False
        End Select

    End Function

    Function CheckVNIVowel(ByVal szInput As Char) As Boolean

        Select Case Asc(szInput)
            Case 65, 69, 79, 212, 85, 214
                'A, E, O, O), U, U)
                Return True

            Case 97, 101, 111, 244, 117, 246
                'a, e, o, o), u, u)
                Return True

            Case 73, 105, 89, 121
                'I, i, Y, y
                Return True
            Case Else
                Return False
        End Select


    End Function


    'Public Function VNItoUNICODE(ByVal vnstr As String) As String
    '    Dim Result As String = ""
    '    Dim c As String = "", i As Integer = 0
    '    Dim db As Boolean
    '    For i = 1 To Len(vnstr)
    '        db = False
    '        If i < Len(vnstr) Then ' Không phải ký tự cuối cùng của chuỗi 
    '            c = Mid(vnstr, i + 1, 1)
    '            If c = "ù" Or c = "ø" Or c = "û" Or c = "õ" Or c = "ï" Or _
    '            c = "ê" Or c = "é" Or c = "è" Or c = "ú" Or c = "ü" Or c = "ë" Or _
    '            c = "â" Or c = "á" Or c = "à" Or c = "å" Or c = "ã" Or c = "ä" Or _
    '            c = "Ù" Or c = "Ø" Or c = "Û" Or c = "Õ" Or c = "Ï" Or _
    '            c = "Ê" Or c = "É" Or c = "È" Or c = "Ú" Or c = "Ü" Or c = "Ë" Or _
    '            c = "Â" Or c = "Á" Or c = "À" Or c = "Å" Or c = "Ã" Or c = "Ä" Then db = True
    '        End If
    '        If db Then
    '            c = Mid(vnstr, i, 2)
    '            Select Case c
    '                Case "aù" : c = ChrW(225)
    '                Case "aø" : c = ChrW(224)
    '                Case "aû" : c = ChrW(7843)
    '                Case "aõ" : c = ChrW(227)
    '                Case "aï" : c = ChrW(7841)
    '                Case "aê" : c = ChrW(259)
    '                Case "aé" : c = ChrW(7855)
    '                Case "aè" : c = ChrW(7857)
    '                Case "aú" : c = ChrW(7859)
    '                Case "aü" : c = ChrW(7861)
    '                Case "aë" : c = ChrW(7863)
    '                Case "aâ" : c = ChrW(226)
    '                Case "aá" : c = ChrW(7845)
    '                Case "aà" : c = ChrW(7847)
    '                Case "aå" : c = ChrW(7849)
    '                Case "aã" : c = ChrW(7851)
    '                Case "aä" : c = ChrW(7853)
    '                Case "eù" : c = ChrW(233)
    '                Case "eø" : c = ChrW(232)
    '                Case "eû" : c = ChrW(7867)
    '                Case "eõ" : c = ChrW(7869)
    '                Case "eï" : c = ChrW(7865)
    '                Case "eâ" : c = ChrW(234)
    '                Case "eá" : c = ChrW(7871)
    '                Case "eà" : c = ChrW(7873)
    '                Case "eå" : c = ChrW(7875)
    '                Case "eã" : c = ChrW(7877)
    '                Case "eä" : c = ChrW(7879)
    '                Case "où" : c = ChrW(243)
    '                Case "oø" : c = ChrW(242)
    '                Case "oû" : c = ChrW(7887)
    '                Case "oõ" : c = ChrW(245)
    '                Case "oï" : c = ChrW(7885)
    '                Case "oâ" : c = ChrW(244)
    '                Case "oá" : c = ChrW(7889)
    '                Case "oà" : c = ChrW(7891)
    '                Case "oå" : c = ChrW(7893)
    '                Case "oã" : c = ChrW(7895)
    '                Case "oä" : c = ChrW(7897)
    '                Case "ôù" : c = ChrW(7899)
    '                Case "ôø" : c = ChrW(7901)
    '                Case "ôû" : c = ChrW(7903)
    '                Case "ôõ" : c = ChrW(7905)
    '                Case "ôï" : c = ChrW(7907)
    '                Case "uù" : c = ChrW(250)
    '                Case "uø" : c = ChrW(249)
    '                Case "uû" : c = ChrW(7911)
    '                Case "uõ" : c = ChrW(361)
    '                Case "uï" : c = ChrW(7909)
    '                Case "öù" : c = ChrW(7913)
    '                Case "öø" : c = ChrW(7915)
    '                Case "öû" : c = ChrW(7917)
    '                Case "öõ" : c = ChrW(7919)
    '                Case "öï" : c = ChrW(7921)
    '                Case "yù" : c = ChrW(253)
    '                Case "yø" : c = ChrW(7923)
    '                Case "yû" : c = ChrW(7927)
    '                Case "yõ" : c = ChrW(7929)
    '                Case "AÙ" : c = ChrW(193)
    '                Case "AØ" : c = ChrW(192)
    '                Case "AÛ" : c = ChrW(7842)
    '                Case "AÕ" : c = ChrW(195)
    '                Case "AÏ" : c = ChrW(7840)
    '                Case "AÊ" : c = ChrW(258)
    '                Case "AÉ" : c = ChrW(7854)
    '                Case "AÈ" : c = ChrW(7856)
    '                Case "AÚ" : c = ChrW(7858)
    '                Case "AÜ" : c = ChrW(7860)
    '                Case "AË" : c = ChrW(7862)
    '                Case "AÂ" : c = ChrW(194)
    '                Case "AÁ" : c = ChrW(7844)
    '                Case "AÀ" : c = ChrW(7846)
    '                Case "AÅ" : c = ChrW(7848)
    '                Case "AÃ" : c = ChrW(7850)
    '                Case "AÄ" : c = ChrW(7852)
    '                Case "EÙ" : c = ChrW(201)
    '                Case "EØ" : c = ChrW(200)
    '                Case "EÛ" : c = ChrW(7866)
    '                Case "EÕ" : c = ChrW(7868)
    '                Case "EÏ" : c = ChrW(7864)
    '                Case "EÂ" : c = ChrW(202)
    '                Case "EÁ" : c = ChrW(7870)
    '                Case "EÀ" : c = ChrW(7872)
    '                Case "EÅ" : c = ChrW(7874)
    '                Case "EÃ" : c = ChrW(7876)
    '                Case "EÄ" : c = ChrW(7878)
    '                Case "OÙ" : c = ChrW(211)
    '                Case "OØ" : c = ChrW(210)
    '                Case "OÛ" : c = ChrW(7886)
    '                Case "OÕ" : c = ChrW(213)
    '                Case "OÏ" : c = ChrW(7884)
    '                Case "OÂ" : c = ChrW(212)
    '                Case "OÁ" : c = ChrW(7888)
    '                Case "OÀ" : c = ChrW(7890)
    '                Case "OÅ" : c = ChrW(7892)
    '                Case "OÃ" : c = ChrW(7894)
    '                Case "OÄ" : c = ChrW(7896)
    '                Case "ÔÙ" : c = ChrW(7898)
    '                Case "ÔØ" : c = ChrW(7900)
    '                Case "ÔÛ" : c = ChrW(7902)
    '                Case "ÔÕ" : c = ChrW(7904)
    '                Case "ÔÏ" : c = ChrW(7906)
    '                Case "UÙ" : c = ChrW(218)
    '                Case "UØ" : c = ChrW(217)
    '                Case "UÛ" : c = ChrW(7910)
    '                Case "UÕ" : c = ChrW(360)
    '                Case "UÏ" : c = ChrW(7908)
    '                Case "ÖÙ" : c = ChrW(7912)
    '                Case "ÖØ" : c = ChrW(7914)
    '                Case "ÖÛ" : c = ChrW(7916)
    '                Case "ÖÕ" : c = ChrW(7918)
    '                Case "ÖÏ" : c = ChrW(7920)
    '                Case "YÙ" : c = ChrW(221)
    '                Case "YØ" : c = ChrW(7922)
    '                Case "YÛ" : c = ChrW(7926)
    '                Case "YÕ" : c = ChrW(7928)
    '            End Select
    '        Else
    '            c = Mid(vnstr, i, 1)
    '            Select Case c
    '                Case "ô" : c = ChrW(417)
    '                Case "í" : c = ChrW(237)
    '                Case "ì" : c = ChrW(236)
    '                Case "æ" : c = ChrW(7881)
    '                Case "ó" : c = ChrW(297)
    '                Case "ò" : c = ChrW(7883)
    '                Case "ö" : c = ChrW(432)
    '                Case "î" : c = ChrW(7925)
    '                Case "ñ" : c = ChrW(273)
    '                Case "Ô" : c = ChrW(416)
    '                Case "Í" : c = ChrW(205)
    '                Case "Ì" : c = ChrW(204)
    '                Case "Æ" : c = ChrW(7880)
    '                Case "Ó" : c = ChrW(296)
    '                Case "Ò" : c = ChrW(7882)
    '                Case "Ö" : c = ChrW(431)
    '                Case "Î" : c = ChrW(7924)
    '                Case "Ñ" : c = ChrW(272)
    '            End Select
    '        End If
    '        Result &= c
    '        If db Then i = i + 1
    '    Next i
    '    Return Result
    'End Function

    ''' <summary>
    ''' Chuyển đổi chuyển dạng Unicode sang VNI
    ''' </summary>
    ''' <param name="vnstr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function ConvertUnicodeToVni(ByVal vnstr As String) As String
        Dim Result As String = ""
        Dim c As String = "", i As Integer = 0
        For i = 1 To Len(vnstr)
            c = Mid(vnstr, i, 1)
            Select Case c
                Case ChrW(97) : c = "a"
                Case ChrW(225) : c = "aù"
                Case ChrW(224) : c = "aø"
                Case ChrW(7843) : c = "aû"
                Case ChrW(227) : c = "aõ"
                Case ChrW(7841) : c = "aï"
                Case ChrW(259) : c = "aê"
                Case ChrW(7855) : c = "aé"
                Case ChrW(7857) : c = "aè"
                Case ChrW(7859) : c = "aú"
                Case ChrW(7861) : c = "aü"
                Case ChrW(7863) : c = "aë"
                Case ChrW(226) : c = "aâ"
                Case ChrW(7845) : c = "aá"
                Case ChrW(7847) : c = "aà"
                Case ChrW(7849) : c = "aå"
                Case ChrW(7851) : c = "aã"
                Case ChrW(7853) : c = "aä"
                Case ChrW(101) : c = "e"
                Case ChrW(233) : c = "eù"
                Case ChrW(232) : c = "eø"
                Case ChrW(7867) : c = "eû"
                Case ChrW(7869) : c = "eõ"
                Case ChrW(7865) : c = "eï"
                Case ChrW(234) : c = "eâ"
                Case ChrW(7871) : c = "eá"
                Case ChrW(7873) : c = "eà"
                Case ChrW(7875) : c = "eå"
                Case ChrW(7877) : c = "eã"
                Case ChrW(7879) : c = "eä"
                Case ChrW(111) : c = "o"
                Case ChrW(243) : c = "où"
                Case ChrW(242) : c = "oø"
                Case ChrW(7887) : c = "oû"
                Case ChrW(245) : c = "oõ"
                Case ChrW(7885) : c = "oï"
                Case ChrW(244) : c = "oâ"
                Case ChrW(7889) : c = "oá"
                Case ChrW(7891) : c = "oà"
                Case ChrW(7893) : c = "oå"
                Case ChrW(7895) : c = "oã"
                Case ChrW(7897) : c = "oä"
                Case ChrW(417) : c = "ô"
                Case ChrW(7899) : c = "ôù"
                Case ChrW(7901) : c = "ôø"
                Case ChrW(7903) : c = "ôû"
                Case ChrW(7905) : c = "ôõ"
                Case ChrW(7907) : c = "ôï"
                Case ChrW(105) : c = "i"
                Case ChrW(237) : c = "í"
                Case ChrW(236) : c = "ì"
                Case ChrW(7881) : c = "æ"
                Case ChrW(297) : c = "ó"
                Case ChrW(7883) : c = "ò"
                Case ChrW(117) : c = "u"
                Case ChrW(250) : c = "uù"
                Case ChrW(249) : c = "uø"
                Case ChrW(7911) : c = "uû"
                Case ChrW(361) : c = "uõ"
                Case ChrW(7909) : c = "uï"
                Case ChrW(432) : c = "ö"
                Case ChrW(7913) : c = "öù"
                Case ChrW(7915) : c = "öø" '"uø"
                Case ChrW(7917) : c = "öû"
                Case ChrW(7919) : c = "öõ"
                Case ChrW(7921) : c = "öï"
                Case ChrW(121) : c = "y"
                Case ChrW(253) : c = "yù"
                Case ChrW(7923) : c = "yø"
                Case ChrW(7927) : c = "yû"
                Case ChrW(7929) : c = "yõ"
                Case ChrW(7925) : c = "î"
                Case ChrW(273) : c = "ñ"
                Case ChrW(65) : c = "A"
                Case ChrW(193) : c = "AÙ"
                Case ChrW(192) : c = "AØ"
                Case ChrW(7842) : c = "AÛ"
                Case ChrW(195) : c = "AÕ"
                Case ChrW(7840) : c = "AÏ"
                Case ChrW(258) : c = "AÊ"
                Case ChrW(7854) : c = "AÉ"
                Case ChrW(7856) : c = "AÈ"
                Case ChrW(7858) : c = "AÚ"
                Case ChrW(7860) : c = "AÜ"
                Case ChrW(7862) : c = "AË"
                Case ChrW(194) : c = "AÂ"
                Case ChrW(7844) : c = "AÁ"
                Case ChrW(7846) : c = "AÀ"
                Case ChrW(7848) : c = "AÅ"
                Case ChrW(7850) : c = "AÃ"
                Case ChrW(7852) : c = "AÄ"
                Case ChrW(69) : c = "E"
                Case ChrW(201) : c = "EÙ"
                Case ChrW(200) : c = "EØ"
                Case ChrW(7866) : c = "EÛ"
                Case ChrW(7868) : c = "EÕ"
                Case ChrW(7864) : c = "EÏ"
                Case ChrW(202) : c = "EÂ"
                Case ChrW(7870) : c = "EÁ"
                Case ChrW(7872) : c = "EÀ"
                Case ChrW(7874) : c = "EÅ"
                Case ChrW(7876) : c = "EÃ"
                Case ChrW(7878) : c = "EÄ"
                Case ChrW(79) : c = "O"
                Case ChrW(211) : c = "OÙ"
                Case ChrW(210) : c = "OØ"
                Case ChrW(7886) : c = "OÛ"
                Case ChrW(213) : c = "OÕ"
                Case ChrW(7884) : c = "OÏ"
                Case ChrW(212) : c = "OÂ"
                Case ChrW(7888) : c = "OÁ"
                Case ChrW(7890) : c = "OÀ"
                Case ChrW(7892) : c = "OÅ"
                Case ChrW(7894) : c = "OÃ"
                Case ChrW(7896) : c = "OÄ"
                Case ChrW(416) : c = "Ô"
                Case ChrW(7898) : c = "ÔÙ"
                Case ChrW(7900) : c = "ÔØ"
                Case ChrW(7902) : c = "ÔÛ"
                Case ChrW(7904) : c = "ÔÕ"
                Case ChrW(7906) : c = "ÔÏ"
                Case ChrW(73) : c = "I"
                Case ChrW(205) : c = "Í"
                Case ChrW(204) : c = "Ì"
                Case ChrW(7880) : c = "Æ"
                Case ChrW(296) : c = "Ó"
                Case ChrW(7882) : c = "Ò"
                Case ChrW(85) : c = "U"
                Case ChrW(218) : c = "UÙ"
                Case ChrW(217) : c = "UØ"
                Case ChrW(7910) : c = "UÛ"
                Case ChrW(360) : c = "UÕ"
                Case ChrW(7908) : c = "UÏ"
                Case ChrW(431) : c = "Ö"
                Case ChrW(7912) : c = "ÖÙ"
                Case ChrW(7914) : c = "ÖØ"
                Case ChrW(7916) : c = "ÖÛ"
                Case ChrW(7918) : c = "ÖÕ"
                Case ChrW(7920) : c = "ÖÏ"
                Case ChrW(89) : c = "Y"
                Case ChrW(221) : c = "YÙ"
                Case ChrW(7922) : c = "YØ"
                Case ChrW(7926) : c = "YÛ"
                Case ChrW(7928) : c = "YÕ"
                Case ChrW(7924) : c = "Î"
                Case ChrW(272) : c = "Ñ"
                Case ChrW(208) : c = "Ñ"
            End Select
            Result &= c
        Next i
        Return Result
    End Function

    ''' <summary>
    ''' Chuyển dữ liệu của Table sang Unicode
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub ConvertVniToUnicode(ByRef dt As DataTable)
        For i As Integer = 0 To dt.Rows.Count - 1 'Số dòng
            For j As Integer = 0 To dt.Columns.Count - 1 'Số cột
                If Not dt.Rows(i).IsNull(j) Then
                    dt.Rows(i).Item(j) = ConvertVniToUnicode(dt.Rows(i).Item(j).ToString)
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Chuyển Vni sang Unicode cho các cột của table
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="myStrArray"> Dim myStrArray() As String = {"RelationName", "RelativeName'} </param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub ConvertDataTable(ByRef dt As DataTable, ByVal myStrArray() As String)
        For i As Integer = 0 To dt.Rows.Count - 1 'Số dòng
            For j As Integer = 0 To myStrArray.Length - 1 'Số cột
                If Not dt.Rows(i).IsNull(j) Then
                    dt.Rows(i).Item(myStrArray(j)) = ConvertVniToUnicode(dt.Rows(i).Item(myStrArray(j)).ToString)
                End If
            Next
        Next
    End Sub

    <DebuggerStepThrough()> _
    Public Function ConvertDataTable(ByVal dt As DataTable, ByVal arr As ArrayList) As Boolean
        If dt Is Nothing Then
            Exit Function
        End If

        Dim j, k As Integer

        For j = 0 To dt.Rows.Count - 1
            For k = 0 To dt.Columns.Count - 1
                If arr.Contains(dt.Columns(k).ColumnName) Then
                    If Not dt.Rows(j).IsNull(k) Then
                        dt.Rows(j).Item(k) = ConvertVniToUnicode(dt.Rows(j).Item(k).ToString)
                    End If
                End If
            Next
        Next

    End Function

    ''' <summary>
    ''' Chuyển Vni sang Unicode cho các cột của datarow
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <param name="myStrArray"> Dim myStrArray() As String = {"RelationName", "RelativeName'} </param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub ConvertDataRow(ByRef dr As DataRow, ByVal myStrArray() As String)
        For i As Integer = 0 To myStrArray.Length - 1 'Số cột
            If Not dr.IsNull(i) Then
                dr.Item(myStrArray(i)) = ConvertVniToUnicode(dr.Item(myStrArray(i)).ToString)
            End If
        Next
    End Sub

    <DebuggerStepThrough()> _
    Public Function ConvertDataView(ByVal dtView As System.Data.DataView, ByVal Column As String) As Boolean
        If dtView Is Nothing Then
            Exit Function
        End If

        Dim j As Integer
        Try
            For j = 0 To dtView.Count - 1
                If IsDBNull(dtView(j).Item(Column)) Then
                Else
                    dtView(j).Item(Column) = ConvertVniToUnicode(dtView(j).Item(Column).ToString)
                End If
            Next
        Catch ex As Exception
        End Try

    End Function

    ''' <summary>
    ''' Convert Heading font of Grid have GroupBy
    ''' </summary>
    ''' <param name="C1Grid"></param>
    ''' <param name="bUseUnicode"></param>
    ''' <remarks>This function must set after LoadLanguage function</remarks>
    Public Sub UnicodeGroupBy(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal bUseUnicode As Boolean)

        If bUseUnicode = False Then Exit Sub

        'Modify date: 17/03/2009: Set Font nh?p Unicode

        For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In C1Grid.Columns

            For i As Integer = 0 To C1Grid.Splits.ColCount - 1

                C1Grid.Splits(i).DisplayColumns(dc.DataField).HeadingStyle.Font = FontUnicode(bUseUnicode)

            Next

            C1Grid.Columns(dc.DataField).Caption = ConvertVniToUnicode(C1Grid.Columns(dc.DataField).Caption)

        Next

    End Sub

#End Region

#Region "Lấy biến CodeTable theo sản phẩm"

    ''' <summary>
    ''' Lấy biến CodeTable theo sản phẩm Lemon3
    ''' </summary>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub GetCodeTable()
        Dim iDisplayProduct As Integer = CInt(GetSetting("Lemon3", "Settings", "DisplayProduct", "0"))
        Dim sSQL As String = SQLStoreD91P3001(iDisplayProduct)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            gbUnicode = L3Bool(dt.Rows(0).Item("IsUnicode"))
        Else
            gbUnicode = False
        End If
        'Bổ sung in báo cáo VNI - Tạm thời dùng câu Select, sẽ bổ sung store sau.
        gbPrintVNI = L3Bool(ReturnScalar("Select IsPrintVNI from D91T0025  WITH(NOLOCK) "))
        dt.Dispose()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P3001
    '# Created User: Nguyễn Thị Minh Hòa
    '# Created Date: 15/11/2010 02:27:29
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD91P3001(ByVal iProduct As Integer) As String
        Dim sSQL As String = ""
        'Kiểm tra có tồn tại store D91P3001 chưa?
        sSQL = " SELECT TOP 1 1 FROM DBO.SYSOBJECTS WHERE ID = OBJECT_ID(N'[DBO].[D91P3001]') AND OBJECTPROPERTY(ID, N'IsProcedure') = 1"
        If Not ExistRecord(sSQL) Then ' chưa tồn tại thì tạo store
            CreateStoreD91P3001(iProduct)
        End If
        sSQL = "Exec D91P3001 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLNumber(iProduct) 'Product, tinyint, NOT NULL
        Return sSQL
    End Function

    Private Sub CreateStoreD91P3001(ByVal iProduct As Integer)
        Dim sSQL As String = ""
        'B1: Kiểm tra table F91T0025 có 3 cột IsUnicodeLemon3ERP, IsUnicodeLemonHR, IsUnicodeLemonFinance, nếu chưa có thì tạo
        sSQL = "IF EXISTS(SELECT * FROM sysobjects WHERE name = 'D91T0025' and xtype = 'U') " & vbCrLf
        sSQL &= "BEGIN	" & vbCrLf
        sSQL &= "IF NOT EXISTS( " & vbCrLf
        sSQL &= "SELECT 		* " & vbCrLf
        sSQL &= "FROM 		syscolumns col " & vbCrLf
        sSQL &= "INNER JOIN	sysobjects tab On col.id = tab.id " & vbCrLf
        sSQL &= "WHERE 		tab.name = 'D91T0025' " & vbCrLf
        sSQL &= "AND	col.name = 'IsUnicodeLemon3ERP' )" & vbCrLf
        sSQL &= " ALTER TABLE D91T0025 ADD IsUnicodeLemon3ERP tinyint NOT NULL DEFAULT(0) " & vbCrLf
        sSQL &= "End " & vbCrLf
        'sSQL &= "Go " & vbCrLf
        sSQL &= "IF EXISTS(SELECT * FROM sysobjects WHERE name = 'D91T0025' and xtype = 'U') " & vbCrLf
        sSQL &= "BEGIN " & vbCrLf
        sSQL &= "IF NOT EXISTS( " & vbCrLf
        sSQL &= "SELECT 		* " & vbCrLf
        sSQL &= "FROM 		syscolumns col " & vbCrLf
        sSQL &= "INNER JOIN	sysobjects tab On col.id = tab.id " & vbCrLf
        sSQL &= "WHERE 		tab.name = 'D91T0025' " & vbCrLf
        sSQL &= "AND	col.name = 'IsUnicodeLemonHR')" & vbCrLf
        sSQL &= " ALTER TABLE D91T0025 ADD IsUnicodeLemonHR tinyint NOT NULL DEFAULT(0)" & vbCrLf
        sSQL &= "End" & vbCrLf
        'sSQL &= "Go " & vbCrLf
        sSQL &= "IF EXISTS(SELECT * FROM sysobjects WHERE name = 'D91T0025' and xtype = 'U')" & vbCrLf
        sSQL &= "BEGIN " & vbCrLf
        sSQL &= "IF NOT EXISTS( " & vbCrLf
        sSQL &= "SELECT 		* " & vbCrLf
        sSQL &= "FROM 		syscolumns col " & vbCrLf
        sSQL &= "INNER JOIN	sysobjects tab On col.id = tab.id " & vbCrLf
        sSQL &= "WHERE 		tab.name = 'D91T0025' " & vbCrLf
        sSQL &= "AND	col.name = 'IsUnicodeLemonFinance')" & vbCrLf
        sSQL &= "	ALTER TABLE D91T0025 ADD IsUnicodeLemonFinance tinyint NOT NULL DEFAULT(0)" & vbCrLf
        sSQL &= "End" & vbCrLf
        ExecuteSQLNoTransaction(sSQL)

        'B2: Kiểm tra table D09T5550 có tồn tại chưa, nếu chưa có thì tạo bảng
        sSQL = "IF  NOT EXISTS (SELECT * FROM DBO.SYSOBJECTS WHERE ID = OBJECT_ID(N'[DBO].[D09T5550]') AND OBJECTPROPERTY(ID, N'IsTable') = 1) "
        sSQL &= "Create Table D09T5550 (UserID varchar(20), CodeTable tinyint Not Null Default(0) "
        sSQL &= "PRIMARY KEY(UserID)) "
        ExecuteSQLNoTransaction(sSQL)

        'B3: Kiểm tra store D91P3001 có chưa, nếu chưa có thì tạo store
        sSQL = "CREATE PROCEDURE D91P3001"
        sSQL &= "(   @UserID As Varchar(50), 	@Product As Tinyint ) AS " & vbCrLf
        sSQL &= "DECLARE @IsUnicode As Tinyint " & vbCrLf
        sSQL &= "SET @IsUnicode=0 " & vbCrLf
        sSQL &= " IF NOT EXISTS (SELECT TOP 1 1 FROM D09T5550 )  " & vbCrLf
        sSQL &= "BEGIN " & vbCrLf
        sSQL &= "IF @Product=0 " & vbCrLf
        sSQL &= "BEGIN" & vbCrLf
        sSQL &= "		SELECT @IsUnicode=IsUnicodeLemon3ERP" & vbCrLf
        sSQL &= "   FROM D91T0025 " & vbCrLf
        sSQL &= "End" & vbCrLf
        sSQL &= "IF @Product=1" & vbCrLf
        sSQL &= "            BEGIN " & vbCrLf
        sSQL &= "SELECT @IsUnicode=IsUnicodeLemonHR" & vbCrLf
        sSQL &= "FROM D91T0025 " & vbCrLf
        sSQL &= "End" & vbCrLf
        sSQL &= "IF @Product=2" & vbCrLf
        sSQL &= "BEGIN " & vbCrLf
        sSQL &= "SELECT @IsUnicode=IsUnicodeLemonFinance" & vbCrLf
        sSQL &= "FROM D91T0025 " & vbCrLf
        sSQL &= "End" & vbCrLf
        sSQL &= "End " & vbCrLf
        sSQL &= "Else " & vbCrLf
        sSQL &= "BEGIN " & vbCrLf
        sSQL &= "SELECT @IsUnicode=ISNULL(CodeTable,0)" & vbCrLf
        sSQL &= "FROM D09T5550 " & vbCrLf
        sSQL &= "WHERE UserID=@UserID" & vbCrLf
        sSQL &= "End" & vbCrLf
        sSQL &= "SELECT  @IsUnicode As IsUnicode"

        ExecuteSQLNoTransaction(sSQL)

    End Sub

#End Region

#Region "Các hàm liên quan nhập bằng Unicode"

    ''' <summary>
    ''' Chuyển Textbox lưu dạng Unicode hay Vni
    ''' </summary>
    ''' <param name="txtName"></param>
    ''' <param name="bSaveUnicode">True: lưu chuỗi dạng Unicode, False: Lưu chuỗi dạng Vni</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SQLStringUnicode(ByVal txtName As System.Windows.Forms.TextBox, ByVal bSaveUnicode As Boolean) As String
        Return SQLStringUnicode(txtName.Text, txtName.Font.Name = "Microsoft Sans Serif", bSaveUnicode)
    End Function

    ''' <summary>
    ''' Chuyển giá trị lưu dạng Unicode hay Vni
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="bSaveUnicode">True: lưu chuỗi dạng Unicode, False: Lưu chuỗi dạng Vni</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SQLStringUnicode(ByVal text As String, ByVal bUseUnicode As Boolean, ByVal bSaveUnicode As Boolean) As String
        If bUseUnicode Then ' Uniocde
            If bSaveUnicode Then
                Return "N" & SQLString(text)
            Else
                Return SQLStringVni(text)
            End If
        Else 'VNI
            If bSaveUnicode Then
                Return SQLStringUnicode(text)
            Else
                Return SQLString(text)
            End If
        End If
    End Function

    ''' <summary>
    ''' Chuyển giá trị lưu dạng Unicode hay Vni
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="bSaveUnicode">True: lưu chuỗi dạng Unicode, False: Lưu chuỗi dạng Vni</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SQLStringUnicode(ByVal text As Object, ByVal bUseUnicode As Boolean, ByVal bSaveUnicode As Boolean) As String
        If text Is Nothing Then Return "N''"
        If IsDBNull(text) Then Return "N''"
        Return SQLStringUnicode(text.ToString, bUseUnicode, bSaveUnicode)
    End Function


    ''' <summary>
    ''' Chuyển chuỗi Text (Vni) sang chuỗi Unicode
    ''' </summary>
    ''' <param name="Text">Chuỗi Vni</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SQLStringUnicode(ByVal Text As String) As String
        Return "N" & SQLString(ConvertVniToUnicode(Text.Trim))
    End Function


    ''' <summary>
    ''' Chuyển chuỗi Text (Unicode) sang chuỗi Vni
    ''' </summary>
    ''' <param name="Text">Chuỗi Unicode</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SQLStringVni(ByVal Text As String) As String
        Return SQLString(ConvertUnicodeToVni(Text.Trim))
        'Update 20/08/2010: Khi đang nhập liệu ở Unicode thì không lưu cho trường VNI
        'Return "''"
    End Function

    ''' <summary>
    ''' Trả về chuỗi Tất cả cho Unicode hay Vni
    ''' </summary>
    ''' <param name="sUnicode">Giá trị chữ Unicode</param>
    ''' <param name="sAll">Giá trị chữ Tất cả</param>
    ''' <param name="bUseUnicode"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub UnicodeAllString(ByRef sUnicode As String, ByRef sAll As String, ByVal bUseUnicode As Boolean)
        If bUseUnicode Then
            sUnicode = "U"
            sAll = "N'" & r("Tat_caU") & "'"
        Else
            sUnicode = ""
            sAll = "'" & r("Tat_caV") & "'"
        End If
    End Sub

    ''' <summary>
    ''' Nối chữ U vào FieldName để load dữ liệu cho form dùng Unicode
    ''' </summary>
    ''' <param name="bUnicode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UnicodeJoin(ByVal bUnicode As Boolean) As String
        Return IIf(bUnicode, "U", "").ToString
    End Function

    ''' <summary>
    ''' Nối chữ Unicode vào tiêu đề cho form Nhập liệu bằng Unicode
    ''' </summary>
    ''' <param name="bUnicode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UnicodeCaption(ByVal bUnicode As Boolean) As String
        Return IIf(bUnicode, Space(1) & "(Unicode)", "").ToString
    End Function


    ''' <summary>
    ''' Chuyển đổi DataField và cột COL_ của lưới (định nghĩa cột bằng string) cho load Unicode
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="ArrCol_Name"></param>
    ''' <param name="bUnicode"></param>
    ''' <remarks></remarks>
    Public Sub UnicodeGridDataField(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ArrCol_Name() As String, ByVal bUnicode As Boolean)
        If Not bUnicode Then Exit Sub
        For i As Integer = 0 To ArrCol_Name.Length - 1
            tdbg.Columns(ArrCol_Name(i)).DataField = ArrCol_Name(i) & "U"
        Next
    End Sub

    ''' <summary>
    ''' Chuyển đổi DataField và cột COL_ của lưới (định nghĩa cột bằng string) cho load Unicode
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="Col_Name"></param>
    ''' <param name="bUnicode"></param>
    ''' <remarks></remarks>
    Public Sub UnicodeGridDataField(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef Col_Name As String, ByVal bUnicode As Boolean)
        If Not bUnicode Then Exit Sub
        tdbg.Columns(Col_Name).DataField = Col_Name & "U"
        Col_Name &= "U"
    End Sub

    ''' <summary>
    ''' Chuyển đổi DataField của lưới (định nghĩa cột bằng Integer) cho load Unicode
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="Col_Name"></param>
    ''' <param name="bUnicode"></param>
    ''' <remarks></remarks>
    Public Sub UnicodeGridDataField(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal Col_Name As Integer, ByVal bUnicode As Boolean)
        If Not bUnicode Then Exit Sub
        tdbg.Columns(Col_Name).DataField = tdbg.Columns(Col_Name).DataField.ToString & "U"
    End Sub

    ''' <summary>
    ''' Chuyển đổi DataField của lưới (định nghĩa cột bằng Integer) cho load Unicode
    ''' </summary>
    ''' <param name="tdbg"></param>
    ''' <param name="ArrCol_Name"></param>
    ''' <param name="bUnicode"></param>
    ''' <remarks></remarks>
    Public Sub UnicodeGridDataField(ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ArrCol_Name() As Integer, ByVal bUnicode As Boolean)
        If Not bUnicode Then Exit Sub
        For i As Integer = 0 To ArrCol_Name.Length - 1
            tdbg.Columns(ArrCol_Name(i)).DataField = tdbg.Columns(ArrCol_Name(i)).DataField & "U"
        Next
    End Sub

    'Public Function FontUnicode(Optional ByVal bUseUnicode As Boolean = True) As System.Drawing.Font
    '    If bUseUnicode Then
    '        Return (New System.Drawing.Font("Microsoft Sans Serif", 8.25))
    '    Else
    '        Return (New System.Drawing.Font("Lemon3", 8.25))
    '    End If
    'End Function
    ''' <summary>
    ''' Font của control
    ''' </summary>
    ''' <param name="bUseUnicode">Tùy chọn Unicode</param>
    ''' <param name="fontStyle">Kiểu font: Bold, Underline, ... (default Normal)</param>
    ''' <returns>Trả về font</returns>
    ''' <remarks>fontStyle: Underline phải truyền vào nếu label hoặc caption của lưới có Tìm kiếm F2</remarks>
    Public Function FontUnicode(Optional ByVal bUseUnicode As Boolean = True, Optional ByVal fontStyle As System.Drawing.FontStyle = FontStyle.Regular, Optional ByVal size As Single = 8.25) As System.Drawing.Font
        If size = 8.25 Then
            size = iSizeFont
        Else 'Lấy size + độ phân giải màn hình
            size += CSng(size - 8.25)
        End If
        Return (New System.Drawing.Font(IIf(bUseUnicode, "Microsoft Sans Serif", "Lemon3").ToString, size, fontStyle, GraphicsUnit.Point))
    End Function


    ''' <summary>
    ''' Lấy đường dẫn báo cáo
    ''' </summary>
    ''' <param name="bUnicode">Tùy chọn Unicode</param>
    ''' <param name="iReportLanguage">Tùy chọn ngôn ngữ bóa cáo</param>
    ''' <param name="sValueCustom">giá trị của combo Đặc thù</param>
    ''' <returns>Đường dẫn báo cáo</returns>
    ''' <remarks>Đối số sValueCustom: nếu form không có combo Đặc thù thì truyền ""</remarks>
    Public Function UnicodeGetReportPath(ByVal bUnicode As Boolean, ByVal iReportLanguage As Byte, ByVal sValueCustom As String) As String
        If sValueCustom <> "" Then Return Application.StartupPath & "\XCustom\"
        If bUnicode Then
            Return Application.StartupPath & IIf(iReportLanguage = 0, "\Reports\", IIf(iReportLanguage = 1, "\Reports\VE-Reports\", "\Reports\E-Reports\")).ToString
        Else
            Return Application.StartupPath & IIf(iReportLanguage = 0, "\XReports\", IIf(iReportLanguage = 1, "\XReports\VE-XReports\", "\XReports\E-XReports\")).ToString
        End If
    End Function

    ''' <summary>
    ''' Trả về SubReport và câu đổ nguồn của SubReport khi Unicode
    ''' </summary>
    ''' <param name="sSubReportID">Tên SubReport</param>
    ''' <param name="sSQLSubReportID">Câu đổ nguồn SubReport</param>
    ''' <param name="sDivisionID">Đơn vị có trên form - Mặc định là "%" nếu form không có combo Đơn vị</param>
    ''' <param name="bUnicode">Tùy chọn Unicode</param>
    ''' <remarks>Form có combo Đơn vị không?</remarks>
    Public Sub UnicodeSubReport(ByRef sSubReportID As String, ByRef sSQLSubReportID As String, Optional ByVal sDivisionID As String = "%", Optional ByVal bUnicode As Boolean = False)
        If bUnicode = False Then Exit Sub

        sSubReportID = "D91R0000"

        sSQLSubReportID = "Select *" & vbCrLf
        sSQLSubReportID &= " FROM D91V0016 " & vbCrLf
        sSQLSubReportID &= " WHERE   DivisionID = " & SQLString(sDivisionID)
    End Sub

    ''' <summary>
    ''' Chuyển font Unicode cho combo, dropdown, lưới, textbox
    ''' </summary>
    ''' <param name="ctrl"></param>
    ''' <param name="bUseUnicode"></param>
    ''' <returns>có convert không?</returns>
    ''' <remarks></remarks>
    Public Function UnicodeConvertFont(ByVal ctrl As Control, ByVal bUseUnicode As Boolean) As Boolean
        'Update 19/06/2013 bỏ đoạn code này vì khi design có thể là font Microsoft Sans Serif
        'If Not bUseUnicode OrElse ctrl.Font.Name = sUnicodeFontName Then Return False
        If ctrl.Font.Name = FontUnicode(gbUnicode).Name Then
            If TypeOf (ctrl) Is C1.Win.C1TrueDBGrid.C1TrueDBGrid Then ConvertFontTDBGrid(CType(ctrl, C1.Win.C1TrueDBGrid.C1TrueDBGrid), bUseUnicode, False) 'TH đúng font lưới nhưng phải convert font của Group
            Return False 'Chỉ convert 1 lần
        End If

        If TypeOf (ctrl) Is TextBox Then
            'Update 20/08/2010: CHuyển Font cho cả những textbox readonly
            ctrl.Font = FontUnicode(bUseUnicode, ctrl.Font.Style)
            Return True
        ElseIf TypeOf (ctrl) Is C1.Win.C1List.C1Combo Then
            Dim tdbc As C1.Win.C1List.C1Combo = CType(ctrl, C1.Win.C1List.C1Combo)
            tdbc.Font = FontUnicode(bUseUnicode)
            tdbc.EditorFont = FontUnicode(bUseUnicode)
            Return True
        ElseIf TypeOf (ctrl) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then 'Bổ sung 25/07/2012
            ctrl.Font = FontUnicode(bUseUnicode, ctrl.Font.Style)
            Return True
        ElseIf TypeOf (ctrl) Is C1.Win.C1TrueDBGrid.C1TrueDBGrid Then
            Dim C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid = CType(ctrl, C1.Win.C1TrueDBGrid.C1TrueDBGrid)
            C1Grid.Font = FontUnicode(bUseUnicode, C1Grid.Font.Style)
            ConvertFontTDBGrid(C1Grid, bUseUnicode)
            'Modify 02/10/2010
            'For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In C1Grid.Columns'Bị lỗi khi có 2 cột có DataField =""
            Return True
        End If
        Return False
    End Function

    Private Sub ConvertFontTDBGrid(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal bUseUnicode As Boolean, Optional ByVal fontGrid As Boolean = True)
        Dim bGroupBy As Boolean = C1Grid.DataView = C1.Win.C1TrueDBGrid.DataViewEnum.GroupBy
        For index As Integer = 0 To C1Grid.Columns.Count - 1
            For iSplit As Integer = 0 To C1Grid.Splits.ColCount - 1
                'Xét các cột
                If fontGrid Then
                    'Modify 02/08/2012
                    C1Grid.Splits(iSplit).DisplayColumns(index).Style.Font = FontUnicode(bUseUnicode, C1Grid.Splits(iSplit).DisplayColumns(index).Style.Font.Style)
                    C1Grid.Splits(iSplit).DisplayColumns(index).EditorStyle.Font = FontUnicode(bUseUnicode, C1Grid.Splits(iSplit).DisplayColumns(index).EditorStyle.Font.Style)
                End If
                '**************
                If bGroupBy Then
                    'TH dịch resource là VNI'Chỉ cần xét split 0
                    If C1Grid.Splits(0).DisplayColumns(index).HeadingStyle.Font.Name.Contains("Lemon3") Then
                        If bUseUnicode Then C1Grid.Columns(index).Caption = ConvertVniToUnicode(C1Grid.Columns(index).Caption)
                    Else 'dịch resource là UNI
                        If bUseUnicode = False Then C1Grid.Columns(index).Caption = ConvertUnicodeToVni(C1Grid.Columns(index).Caption)
                    End If
                    'TH dịch resource là VNI
                    If C1Grid.GroupStyle.Font.Name.Contains("Lemon3") Then
                        If bUseUnicode Then C1Grid.GroupByCaption = ConvertVniToUnicode(C1Grid.GroupByCaption)
                    Else 'dịch resource là UNI
                        If bUseUnicode = False Then C1Grid.GroupByCaption = ConvertUnicodeToVni(C1Grid.GroupByCaption)
                    End If
                    C1Grid.GroupStyle.Font = FontUnicode(bUseUnicode)
                    C1Grid.Splits(iSplit).DisplayColumns(index).HeadingStyle.Font = FontUnicode(bUseUnicode)
                End If
                '  Exit For
                ' End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Chuyển các control sang nhập liệu Uniocde
    ''' </summary>
    ''' <param name="control"></param>
    ''' <param name="bUseUnicode"></param>
    ''' <param name="bGroupBy">lưới có dùng GroupBy không?</param>
    ''' <remarks>bGroupBy không có tác dụng</remarks>
    Public Sub InputbyUnicode(ByVal control As Control, ByVal bUseUnicode As Boolean, Optional ByVal bGroupBy As Boolean = False)

        'If Not bUseUnicode Then Exit Sub'Update 25/06/2013 bỏ đoạn code này vì khi design có thể là font Microsoft Sans Serif

        For Each ctrl As Control In control.Controls
            'If TypeOf (ctrl) Is TextBox Then
            '    'Update 20/08/2010: CHuyển Font cho cả những textbox readonly
            '    ctrl.Font = FontUnicode(, ctrl.Font.Style)
            '    Continue For
            'ElseIf TypeOf (ctrl) Is C1.Win.C1List.C1Combo Then
            '    Dim tdbc As C1.Win.C1List.C1Combo = CType(ctrl, C1.Win.C1List.C1Combo)
            '    tdbc.Font = FontUnicode()
            '    tdbc.EditorFont = FontUnicode()
            '    Continue For
            'ElseIf TypeOf (ctrl) Is C1.Win.C1TrueDBGrid.C1TrueDBDropdown Then 'Bổ sung 25/07/2012
            '    ctrl.Font = FontUnicode(bUseUnicode, ctrl.Font.Style)
            '    Continue For
            'ElseIf TypeOf (ctrl) Is C1.Win.C1TrueDBGrid.C1TrueDBGrid Then
            '    Dim C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid = CType(ctrl, C1.Win.C1TrueDBGrid.C1TrueDBGrid)
            '    C1Grid.Font = FontUnicode(bUseUnicode, C1Grid.Font.Style)

            '    For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In C1Grid.Columns
            '        For i As Integer = 0 To C1Grid.Splits.ColCount - 1
            '            C1Grid.Splits(i).DisplayColumns(dc.DataField).Style.Font = C1Grid.Font
            '            C1Grid.Splits(i).DisplayColumns(dc.DataField).EditorStyle.Font = C1Grid.Font 'Bổ sung 27/05/2011

            '            If bGroupBy Then
            '                If C1Grid.Splits(i).DisplayColumns(dc).HeadingStyle.Font.Name.Contains("Lemon3") Then C1Grid.Columns(dc.DataField).Caption = ConvertVniToUnicode(C1Grid.Columns(dc.DataField).Caption)

            '                C1Grid.Splits(i).DisplayColumns(dc).HeadingStyle.Font = FontUnicode(gbUnicode)
            '            End If
            '        Next
            '    Next
            If UnicodeConvertFont(ctrl, bUseUnicode) Then
                Continue For
            ElseIf TypeOf (ctrl) Is TabControl Or TypeOf (ctrl) Is TabPage Or TypeOf (ctrl) Is GroupBox Or TypeOf (ctrl) Is Panel Then
                ' AdjustFontChildControl(ctrl)
                InputbyUnicode(ctrl, bUseUnicode, bGroupBy)
            End If
        Next
    End Sub

    'Private Sub AdjustFontChildControl(ByVal ctrl As Control)
    '    If TypeOf (ctrl) Is TextBox Then
    '        'Update 20/08/2010: CHuyển Font cho cả những textbox readonly
    '        'If CType(ctrl, TextBox).ReadOnly = False Then ctrl.Font = FontUnicode()
    '        ctrl.Font = FontUnicode(, ctrl.Font.Style)
    '        Exit Sub
    '    ElseIf TypeOf (ctrl) Is TabControl Or TypeOf (ctrl) Is TabPage Or TypeOf (ctrl) Is GroupBox Or TypeOf (ctrl) Is Panel Then
    '        For Each childControl As Control In ctrl.Controls
    '            AdjustFontChildControl(childControl)
    '        Next
    '    Else
    '        Exit Sub
    '    End If
    'End Sub



#End Region

End Module
