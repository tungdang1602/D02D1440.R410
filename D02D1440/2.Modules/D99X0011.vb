'#######################################################################################
'#                                     CHÚ Ý (Các hàm chung (viết tiếp cho D99X0000))
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày tạo: 29/11/2010 
'# Ngày cập nhật cuối cùng:  10/01/2014
'# Người cập nhật cuối cùng: Minh Hòa
'# Diễn giải: 
'# Đưa hàm CheckConnection vào đây
'# Sửa hàm ClearText
'# Bỏ sự kiện c1dateDate_LostFocus
'# Bổ sung WITH (NOLOCK) vào table, trong bảng D91T0000 23/9/2013
'#Bổ sung GetTextCreateBy_new và LockColums, UnLockColums 4/11/2013
'# Bổ sung Tìm dòng trên lưới :findrowInGrid() 12/11/2013
'# Sửa lại kiểm tra Ngày hóa đơn lớn hơn Ngày phiếu : CheckInvoiceDateWithVoucherDate  20/11/2013
'# Bổ sung các hàm thông báo chung AskDelete: 10/01/2014
'#######################################################################################
Imports system.IO
Imports System.Security.Cryptography

Module D99X0011

#Region "Khai báo biến toàn cục"
    ''' <summary>
    ''' Màu nền của Control bắt buộc nhập
    ''' </summary>
    ''' <remarks></remarks>
    Public COLOR_BACKCOLORWARNING As System.Drawing.Color = Color.Beige
    Public gbyD91RefDate As Integer = -1

#End Region

#Region "Lấy màu"
    ''' <summary>
    ''' Lấy màu bắt buộc nhập cho các control
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub GetBackColorObligatory()
        COLOR_BACKCOLOROBLIGATORY = ColorTranslator.FromHtml("#" & D99C0007.GetModulesSetting("D00", ModuleOption.lmOptions, "ColorObligatory", "F5F5DC"))
        COLOR_BACKCOLORWARNING = ColorTranslator.FromHtml("#" & D99C0007.GetModulesSetting("D00", ModuleOption.lmOptions, "ColorWarning", "FF0000"))
    End Sub
#End Region

#Region "Nhập ngày trên lưới"
    'Dùng Focus đúng các cột động dạng Ngày
    Private Sub c1dateDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim c1date As C1.Win.C1Input.C1DateEdit = CType(sender, C1.Win.C1Input.C1DateEdit)
        'If c1date.Parent.GetType.Name = "C1TrueDBGrid" Then c1date.Parent.Focus() 'Khi nhập liệu trên lưới Enter không qua cột kế tiếp
        'CType(sender, C1.Win.C1Input.C1DateEdit).Parent.Focus()'bị lỗi nhập liệu không có button chọn ngày
    End Sub

    Public Sub InputDateInTrueDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Date() As Integer)
        For i As Integer = 0 To COL_Date.Length - 1
            tdbg.Columns(COL_Date(i)).EditMask = ""
            tdbg.Columns(COL_Date(i)).EditMaskUpdate = False
            tdbg.Columns(COL_Date(i)).EnableDateTimeEditor = True

            tdbg.Columns(COL_Date(i)).NumberFormat = "Short Date"

            Dim c1dateTemp As New C1.Win.C1Input.C1DateEdit
            c1dateTemp.Visible = False
            c1dateTemp.CustomFormat = "dd/MM/yyyy"
            c1dateTemp.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
            c1dateTemp.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
            c1dateTemp.ValueIsDbNull = True
            c1dateTemp.EmptyAsNull = True
            c1dateTemp.TabStop = False
            AddHandler c1dateTemp.LostFocus, AddressOf c1dateDate_LostFocus

            tdbg.Columns(COL_Date(i)).Editor = c1dateTemp
        Next
    End Sub

    Public Sub InputDateInTrueDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ParamArray COL_Date() As String)
        Dim arrInteger(COL_Date.Length - 1) As Integer
        For i As Integer = 0 To COL_Date.Length - 1
            arrInteger.SetValue(IndexOfColumn(tdbg, COL_Date(i)), i)
        Next
        InputDateInTrueDBGrid(tdbg, arrInteger)
    End Sub
#End Region

#Region "Kiểm tra ngày hóa đơn lớn hơn ngày phiếu"
    ''' <summary>
    ''' Kiểm tra ngày chứng từ và Ngày hóa đơn
    ''' </summary>
    ''' <param name="c1VoucherDate">Required. Date voucher control.</param>
    ''' <param name="c1RefDate">Required. Date invoice control.</param>
    ''' <remarks>Return True : Valid</remarks>
    Public Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal c1VoucherDate As C1.Win.C1Input.C1DateEdit, ByVal c1RefDate As C1.Win.C1Input.C1DateEdit, ByVal sRefNo As String, ByVal sSerialNo As String, Optional ByVal tabSelection As System.Windows.Forms.TabControl = Nothing, Optional ByVal Index As Integer = -1) As Boolean
        If gbyD91RefDate = -1 Then
            GetD91RefDateCheck(sModuleID) '23/10/2013, Văn Tâm-Ngọc Thoại: id 79723- 	Kiểm tra ngày hóa đơn lớn hơn ngày phiếu
        End If
        If gbyD91RefDate = 0 Then Return True ' Không kiểm tra
        'Ngày phiếu hay Ngày hóa đơn bằng rỗng thì không kiểm tra
        If c1RefDate.Text = "" OrElse c1VoucherDate.Text = "" Then Return True
        'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
        If sRefNo = "" And sSerialNo = "" Then Return True

        'Sửa lỗi cho TH nhập tiếp gán   thì ngày phiếu
        '  If Convert.ToDateTime(c1InvoiceDate.Value) <= Convert.ToDateTime(c1VoucherDate.Value) Then Return True
        If CDate(SQLDateShow(c1RefDate.Value.ToString)) <= CDate(SQLDateShow(c1VoucherDate.Value.ToString)) Then Return True
        Select Case gbyD91RefDate
            Case 1 ' Kiểm tra thông báo
                If D99C0008.MsgAsk(r("MSG000036") & vbCrLf & r("MSG000021")) = Windows.Forms.DialogResult.No Then
                    If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                    c1RefDate.Focus()
                    Return False
                End If
            Case 2 ' Kiểm tra không cho lưu
                D99C0008.MsgL3(r("MSG000036") & vbCrLf & r("MSG000053"), L3MessageBoxIcon.Exclamation)
                If tabSelection IsNot Nothing AndAlso Index <> -1 Then tabSelection.SelectedIndex = Index
                c1RefDate.Focus()
                Return False
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Kiểm tra ngày chứng từ và Ngày hóa đơn trên lưới (dataField column)
    ''' </summary>
    ''' <param name="c1VoucherDate">Required. Date control.</param>
    ''' <param name="tdbg">lưới</param>
    ''' <param name="COL_RefDate">DataField cột Ngày hóa đơn trên lưới. String</param>
    ''' <param name="iSplit">split chứa cột Ngày hóa đơn. Integer</param>
    ''' <remarks>Return True : Valid</remarks>
    Private Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal c1VoucherDate As C1.Win.C1Input.C1DateEdit, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_RefDate As String, ByVal COL_RefNo As String, ByVal COL_SerialNo As String, Optional ByVal iSplit As Integer = 0) As Boolean ', ByRef bCheckOkDate As Boolean, ByRef iIndexError As Integer,
        If gbyD91RefDate = -1 Then
            GetD91RefDateCheck(sModuleID) '23/10/2013, Văn Tâm-Ngọc Thoại: id 79723- 	Kiểm tra ngày hóa đơn lớn hơn ngày phiếu
        End If
        If gbyD91RefDate = 0 Then Return True ' Không kiểm tra

        'Ngày phiếu bằng rỗng thì không kiểm tra
        If c1VoucherDate.Text = "" Then Return True


        Dim dtGrid As DataTable = CType(tdbg.DataSource, DataTable)
        'Xóa những dòng vi phạm trước đó
        For Each row As DataRow In dtGrid.GetErrors()
            row.ClearErrors()
        Next

        '*************
        Dim sWhere As String = ""
        sWhere = COL_RefDate & ">= #" & DateSave(CDate(c1VoucherDate.Value).AddDays(1)) & "#"
        'Số số hóa đơn và số sêrial bằng rỗng thì không kiểm tra
        sWhere &= " And (" & Col_RefNo & " <>'' Or " & COL_SerialNo & " <>'' )"

        'Dim dr() As DataRow = dtGrid.Select(COL_InvoiceDate & ">= #" & DateSave(CDate(c1VoucherDate.Value).AddDays(1)) & "#")
        Dim dr() As DataRow = dtGrid.Select(sWhere)
        If dr.Length = 0 Then Return True

        Select Case gbyD91RefDate
            Case 1 ' Kiểm tra thông báo
                If D99C0008.MsgAsk(r("MSG000036") & vbCrLf & r("MSG000021")) = Windows.Forms.DialogResult.No Then
                    For i As Integer = 0 To dr.Length - 1
                        dr(i).RowError = "ErrorDate"
                    Next
                    tdbg.Focus()
                    tdbg.SplitIndex = iSplit
                    tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                    tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                    Return False
                End If

            Case 2 ' Kiểm tra không cho lưu
                D99C0008.MsgL3(r("MSG000036") & vbCrLf & r("MSG000053"), L3MessageBoxIcon.Exclamation)
                For i As Integer = 0 To dr.Length - 1
                    dr(i).RowError = "ErrorDate"
                Next
                tdbg.Focus()
                tdbg.SplitIndex = iSplit
                tdbg.Col = tdbg.Columns.IndexOf(tdbg.Columns(COL_RefDate))
                tdbg.Row = dtGrid.Rows.IndexOf(dr(0))
                Return False
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Kiểm tra ngày chứng từ và Ngày hóa đơn trên lưới (index column)
    ''' </summary>
    ''' <param name="c1VoucherDate">Required. Date control.</param>
    ''' <param name="tdbg">lưới</param>
    ''' <param name="COL_InvoiceDate">index cột Ngày hóa đơn trên lưới. Integer</param>
    ''' <param name="iSplit">split chứa cột Ngày hóa đơn. Integer</param>
    ''' <remarks>Return ></remarks>
    Public Function CheckInvoiceDateWithVoucherDate(ByVal sModuleID As String, ByVal c1VoucherDate As C1.Win.C1Input.C1DateEdit, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_InvoiceDate As Integer, ByVal COL_RefNo As Integer, ByVal COL_SerialNo As Integer, Optional ByVal iSplit As Integer = 0) As Boolean
        Return CheckInvoiceDateWithVoucherDate(sModuleID, c1VoucherDate, tdbg, tdbg.Columns(COL_InvoiceDate).DataField, tdbg.Columns(COL_RefNo).DataField, tdbg.Columns(COL_SerialNo).DataField, iSplit)
    End Function

    'Lấy giá trị thiết lập từ D91
    Public Sub GetD91RefDateCheck(ByVal sModule As String)
        'Return CByte(ReturnScalar("SELECT Top 1 RefDateCheck FROM D91T9102 WITH(NOLOCK) "))
        If sModule.Length = 2 Then sModule = "D" & sModule
        gbyD91RefDate = CByte(ReturnScalar("Exec D91P9301 " & SQLString(sModule)))
    End Sub

#End Region

#Region "Kiểm tra số sêri trùng số hóa đơn"
    Public Function CheckSerialAndRefNo(ByVal sTableName As String, _
                        ByVal sBatchID As String, _
                        ByVal sRefNo As String, _
                        ByVal sSerialNo As String, _
                        ByVal sVoucherID As String, ByVal sModuleID As String) As Boolean

        sModuleID = L3Right(sModuleID, 2)
        Dim sSQL As String
        Dim dt As New DataTable
        sSQL = ""
        sSQL = "Exec D91P9103 " & SQLString(gsDivisionID) & COMMA & SQLString(sModuleID) & COMMA & ""
        sSQL = sSQL & SQLString(sTableName) & COMMA & SQLString(sBatchID) & COMMA & ""
        sSQL = sSQL & SQLString(sRefNo) & COMMA & SQLString(sSerialNo) & COMMA & SQLString(gsLanguage) & COMMA
        sSQL = sSQL & SQLString(sVoucherID)

        dt = ReturnDataTable(sSQL)

        If dt.Rows.Count > 0 Then
            Select Case CInt(dt.Rows(0).Item("Status"))
                Case 0
                Case 1
                    'If MessageBox.Show(dt.Rows(0).Item("Message").ToString, MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                    If D99C0008.Msg(ConvertVietwareFToUnicode(dt.Rows(0).Item("Message").ToString), MsgAnnouncement, L3MessageBoxButtons.YesNo, L3MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.No Then
                        dt = Nothing
                        Return True
                    End If
                Case 2
                    'MessageBox.Show(dt.Rows(0).Item("Message").ToString, MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    D99C0008.MsgL3(ConvertVietwareFToUnicode(dt.Rows(0).Item("Message").ToString), L3MessageBoxIcon.Exclamation)
                    dt = Nothing
                    Return True
            End Select
        End If
        dt = Nothing
        Return False

        'Return Not CheckStoreNew(sSQL)

    End Function
#End Region


#Region "Đọc tiền bằng chữ"

    Public Function ConvertNumberToString(ByVal sNumber As Object, Optional ByVal CurrencyID As String = "VND", Optional ByVal bEnglish As Boolean = False) As String
        If bEnglish Then ' Đọc tiền bằng tiếng Anh
            Return ConvertNumberToString_English(sNumber, CurrencyID)
        End If

        'Đọc tiền bằng tiếng Việt (dạng Unicode)
        Try
            Dim nLen As Integer
            Dim i As Integer
            Dim j As Integer
            Dim nDigit As Integer
            Dim sTemp As String
            Dim sNumText() As String
            Dim sSeparatorNumber() As String 'Lưu 2 giá trị số phần nguyên và phần lẻ
            Dim sSeparatorValue(1) As String 'Lưu 2 giá trị chuỗi phần nguyên và phần lẻ

            Dim sUnit As String
            Dim sSubUnit As String
            Dim sSeparator As String
            Dim sStartSymbol As String
            Dim sEndSymbol As String
            Dim sResult As String = ""

            If Not IsNumeric(sNumber) Then Return "Số không hợp lệ !"
            If Math.Abs(CDbl(sNumber)) >= 1.0E+15 Then Return "Số quá lớn !"

            Dim NumString As String = sNumber.ToString  '= Format(Math.Abs(WhatNumber), "##############0.00")

            sUnit = ""
            sSubUnit = ""
            sSeparator = "và"
            sEndSymbol = "chẵn"

            'Lấy tên đơn vị tiền và đơn vị tiền lẻ
            GetCurrencyUnitName(CurrencyID, sUnit, sSubUnit)

            Dim bHundred As Boolean = False 'Có hàng trăm hay không
            Dim bUnit As Boolean = False 'Có hàng chục và đơn vị hay không
            Dim iCount As Integer 'Đếm số lần tăng hàng chục và đơn vị

            If Val(NumString) >= 0 Then
                sStartSymbol = ""
            Else
                sStartSymbol = "Âm"
            End If
            sNumText = Split("không;một;hai;ba;bốn;năm;sáu;bảy;tám;chín", ";")

            NumString = Format(Math.Abs(CDbl(sNumber)), "##############0.00")

            'Nếu tiền là VND thì bỏ số lẻ, không đọc phần xu
            If CurrencyID = "VND" Then
                NumString = Math.Round(CDbl(NumString)).ToString
            End If

            sSeparatorNumber = Split(NumString, ".") 'Tách phần nguyên và phần thập phân

            For j = 0 To sSeparatorNumber.Length - 1

                sTemp = ""

                NumString = sSeparatorNumber(j)

                nLen = Len(NumString)

                For i = 1 To nLen

                    nDigit = CInt(Mid(NumString, i, 1))

                    If bHundred Then
                        iCount += 1
                        If nDigit <> 0 Then
                            bUnit = True
                        End If
                    End If
                    sTemp &= " " & sNumText(nDigit)
                    If (nLen = i) Then Exit For

                    Select Case (nLen - i) Mod 9
                        Case 0
                            'sTemp &= " tỉ,"
                            sTemp &= " tỉ"
                            If Mid(NumString, i + 1, 3) = "000" Then
                                If nLen - i > 9 Then
                                    sTemp = Mid(sTemp, 1, sTemp.Length - 1) 'Cắt dấu phầy cuối
                                    'sTemp &= " tỉ,"
                                    sTemp &= " tỉ"
                                End If
                                i = i + 3
                            End If
                            If Mid(NumString, i + 1, 3) = "000" Then i = i + 3
                            If Mid(NumString, i + 1, 3) = "000" Then i = i + 3
                        Case 6
                            'sTemp &= " triệu,"
                            sTemp &= " triệu"
                            If Mid(NumString, i + 1, 3) = "000" Then
                                If nLen - i > 9 Then
                                    sTemp = Mid(sTemp, 1, sTemp.Length - 1) 'Cắt dấu phầy cuối
                                    'sTemp &= " tỉ,"
                                    sTemp &= " tỉ"
                                End If
                                i = i + 3
                            End If
                            If Mid(NumString, i + 1, 3) = "000" Then i = i + 3
                        Case 3
                            'sTemp &= " ngàn,"
                            sTemp &= " ngàn"
                            If Mid(NumString, i + 1, 3) = "000" Then
                                If nLen - i > 9 Then
                                    sTemp = Mid(sTemp, 1, sTemp.Length - 1) 'Cắt dấu phầy cuối
                                    'sTemp &= " tỉ,"
                                    sTemp &= " tỉ"
                                End If
                                i = i + 3
                            End If
                        Case Else
                            Select Case (nLen - i) Mod 3
                                Case 2
                                    bHundred = True
                                    sTemp &= " trăm"
                                Case 1
                                    sTemp &= " mươi"
                            End Select
                    End Select
                    If iCount = 2 And bUnit = False Then
                        iCount = 0
                        bHundred = False
                        bUnit = False
                        Dim sValues() As String = Split(sTemp, ",")
                        If sValues.Length > 1 AndAlso sValues(1).Contains("tỉ") Then
                            'sTemp = sValues(0) & " tỉ,"
                            sTemp = sValues(0) & " tỉ"
                        Else
                            'sTemp = sValues(0) & ","
                            sTemp = sValues(0) '& ","
                        End If
                    End If
                Next i

                sTemp = Replace(sTemp, "không mươi không ", "")

                sTemp = Replace(sTemp, "không mươi ", "lẻ ")

                sTemp = Replace(sTemp, "mươi không ", "mươi")

                sTemp = Replace(sTemp, "một mươi", "mười")

                'sTemp = Replace(sTemp, "mươi bốn", "mươi tư")

                'sTemp = Replace(sTemp, "lẻ bốn", "lẻ tư")

                sTemp = Replace(sTemp, "mươi năm", "mươi lăm")

                sTemp = Replace(sTemp, "mươi một", "mươi mốt")

                sTemp = Replace(sTemp, "mười năm", "mười lăm")

                sTemp = Trim(sTemp)

                If sTemp <> "" Then If Mid(sTemp, sTemp.Length, 1) = "," Then sTemp = Mid(sTemp, 1, sTemp.Length - 1)

                sSeparatorValue(j) = sTemp

            Next j
            Select Case sSeparatorValue(1)
                Case Nothing, "không", "không trăm", "không ngàn", "không triệu", "không tỉ"
                    sSeparatorValue(1) = ""
            End Select
            If sSeparatorValue(1) Is Nothing Then sSeparatorValue(1) = ""

            If sSeparatorNumber.Length - 1 > 0 And sSeparatorValue(1).Length > 0 Then
                If sStartSymbol = "" Then
                    sResult = UCase(Mid(sSeparatorValue(0), 1, 1)) & Mid(sSeparatorValue(0), 2) & " " & sUnit & " " & sSeparator & " " & Mid(sSeparatorValue(1), 1, 1) & Mid(sSeparatorValue(1), 2) & " " & sSubUnit
                Else
                    sResult = sStartSymbol & " " & Mid(sSeparatorValue(0), 1, 1) & Mid(sSeparatorValue(0), 2) & " " & sUnit & " " & sSeparator & " " & Mid(sSeparatorValue(1), 1, 1) & Mid(sSeparatorValue(1), 2) & " " & sSubUnit
                End If
            Else
                If sStartSymbol = "" Then
                    sResult = UCase(Mid(sSeparatorValue(0), 1, 1)) & Mid(sSeparatorValue(0), 2) & " " & sUnit & " " & sEndSymbol
                Else
                    sResult = sStartSymbol & " " & Mid(sSeparatorValue(0), 1, 1) & Mid(sSeparatorValue(0), 2) & " " & sUnit & " " & sEndSymbol
                End If
            End If

            Return sResult
        Catch ex As Exception
            D99C0008.MsgL3("Lỗi đọc tiền", L3MessageBoxIcon.Err)
            Return ""
        End Try

    End Function

    Private Function ConvertNumberToString_English(ByVal sNumber As Object, Optional ByVal CurrencyID As String = "VND") As String
        ' Đọc tiền bằng tiếng Anh
        Dim ToRead, NumString, Group, Word As String
        Dim WhatNumber As Double = CDbl(sNumber)

        If Not IsNumeric(sNumber) Then Return "Number is invalid !"
        If Math.Abs(CDbl(sNumber)) >= 1.0E+15 Then Return "Too long number !"

        If WhatNumber = 0 Then
            ToRead = "None"
        Else
            Dim i As Byte, W, X, Y, Z As Integer
            Dim FristColum() As String = {"None", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", _
                    "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eightteen", "nineteen"}
            Dim SecondColum() As String = {"None", "None", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"}

            Dim sUnit As String = ""
            Dim sSubUnit As String = ""
            'Lấy tên đơn vị tiền và đơn vị tiền lẻ
            GetCurrencyUnitName(CurrencyID, sUnit, sSubUnit)

            Dim ReadMetho() As String = {"None", "trillion", "billion", "million", "thousand", sUnit, sSubUnit}

            If WhatNumber < 0 Then
                ToRead = "Minus" & Space(1)
            Else
                ToRead = Space(0)
            End If
            'Nếu tiền là VND thì bỏ số lẻ, không đọc phần xu
            If CurrencyID = "VND" Then
                NumString = Format(Math.Round(Math.Abs(WhatNumber)), "##############0.00")
            Else
                NumString = Format(Math.Abs(WhatNumber), "##############0.00")
            End If

            NumString = Right(Space(15) & NumString, 18)
            For i = 1 To 6
                Group = Mid(NumString, i * 3 - 2, 3)
                If Group <> Space(3) Then
                    Select Case Group
                        Case "000"
                            If i = 5 And Math.Abs(WhatNumber) > 1 Then
                                'Word = "Vietnamese dong" & Space(1)
                                Word = ReadMetho(i) & Space(1)
                            Else
                                Word = Space(0)
                            End If
                        Case ".00"
                            Word = "only"
                        Case Else
                            X = L3Int(Left(Group, 1))
                            Y = L3Int(Mid(Group, 2, 1))
                            Z = L3Int(Right(Group, 1))
                            W = L3Int(Right(Group, 2))
                            If X = 0 Then
                                Word = Space(0)
                            Else
                                Word = FristColum(X) & Space(1) & "hundred" & Space(1)
                                If W > 0 And W < 21 Then
                                    Word = Word & "and" & Space(1)
                                End If
                            End If
                            If i = 6 And Math.Abs(WhatNumber) > 1 Then
                                Word = "and" & Space(1) & Word
                            End If
                            If W < 20 And W > 0 Then
                                Word = Word & FristColum(W) & Space(1)
                            Else
                                If W >= 20 Then
                                    Word = Word & SecondColum(Y) & Space(1)
                                    If Z > 0 Then
                                        Word = Word & FristColum(Z) & Space(1)
                                    End If
                                End If
                            End If
                            Word = Word & ReadMetho(i) & Space(1)
                    End Select
                    ToRead = ToRead & Word
                End If
            Next i
        End If

        Dim sResult As String = ""
        sResult = UCase(Left(ToRead, 1)) & Mid(ToRead, 2)

        Return sResult
    End Function

    Private Sub GetCurrencyUnitName(ByVal CurrencyID As String, ByRef UnitName As String, ByRef SubUnitName As String)
        ' Lấy Tên tiền và đơn vị tiền
        Dim sSQL As String
        sSQL = "Select UnitNameU, SubUnitNameU From D91T0010 WITH(NOLOCK)  Where (CurrencyID = " & SQLString(CurrencyID) & ")  And (Disabled = 0)"
        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            UnitName = dt.Rows(0).Item("UnitNameU").ToString
            SubUnitName = dt.Rows(0).Item("SubUnitNameU").ToString
        End If
        dt.Dispose()

    End Sub

#End Region

#Region "Xóa text các control khi nhấn Nhập tiếp"

    '''' <summary>
    '''' Xóa tất cả các control
    '''' </summary>
    '''' <param name="ctrl">truyền vào Me</param>
    '''' <remarks>Cách gọi: tại nút Nhập tiếp, gọi ClearText (Me)</remarks>
    '<DebuggerStepThrough()> _
    Public Sub ClearText(ByVal ctrl As Control)
        If TypeOf (ctrl) Is C1.Win.C1Input.C1DateEdit Then
            CType(ctrl, C1.Win.C1Input.C1DateEdit).Value = ""
        ElseIf TypeOf (ctrl) Is C1.Win.C1Input.C1NumericEdit Then
            CType(ctrl, C1.Win.C1Input.C1NumericEdit).Value = ""
        ElseIf (TypeOf (ctrl) Is CheckBox) Then
            CType(ctrl, CheckBox).Checked = False
        ElseIf (TypeOf (ctrl) Is TextBox Or TypeOf (ctrl) Is C1.Win.C1List.C1Combo) Then
            'Update 15/08/2013: Filter Bar trên lưới có TypeOf là TextBox nên chạy vào điều kiện này
            'ctrl.Text = ""
            If ctrl.Name <> "" Then ctrl.Text = ""
        End If

        For Each childControl As Control In ctrl.Controls
            ClearText(childControl)
        Next
    End Sub


    ''' <summary>
    ''' Xóa tất cả các control trừ các control trong tập truyền vào ctrlExclude
    ''' </summary>
    ''' <param name="ctrl">truyền vào Me</param>
    ''' <param name="ctrlExclude">truyền vào tập các control (txtName, tdbcCurrencyID, c1dateVoucherDate) không gán text =""</param>
    ''' <remarks>Cách gọi: tại nút Nhập tiếp, gọi ClearText (Me, txtName, tdbcCurrencyID, c1dateVoucherDate)</remarks>
    Public Sub ClearText(ByVal ctrl As Control, ByVal ParamArray ctrlExclude() As Control)
        If TypeOf (ctrl) Is C1.Win.C1Input.C1DateEdit Then
            'If ctrl.Visible = True Then CType(ctrl, C1.Win.C1Input.C1DateEdit).Value = ""
            CType(ctrl, C1.Win.C1Input.C1DateEdit).Value = ""
        ElseIf TypeOf (ctrl) Is C1.Win.C1Input.C1NumericEdit Then
            CType(ctrl, C1.Win.C1Input.C1NumericEdit).Value = ""
        ElseIf (TypeOf (ctrl) Is CheckBox) Then
            CType(ctrl, CheckBox).Checked = False
        ElseIf (TypeOf (ctrl) Is TextBox Or TypeOf (ctrl) Is C1.Win.C1List.C1Combo) Then
            ctrl.Text = ""
        End If

        For Each childControl As Control In ctrl.Controls
            If (TypeOf (childControl) Is TextBox) OrElse (TypeOf (childControl) Is C1.Win.C1List.C1Combo) _
            OrElse (TypeOf (childControl) Is C1.Win.C1Input.C1DateEdit) OrElse (TypeOf (childControl) Is CheckBox) Then
                If FindControl(childControl, ctrlExclude) = False Then
                    ClearText(childControl, ctrlExclude)
                End If
            Else
                ClearText(childControl, ctrlExclude)
            End If
        Next
    End Sub


    Dim sValue As String = ""
    Private Function ContainsValue(ByVal s As Control) As Boolean
        Return s.Name.Equals(sValue)
    End Function

    Private Function FindControl(ByVal ctrl As Control, ByVal ParamArray ArrString() As Control) As Boolean
        sValue = ctrl.Name
        If Array.Exists(ArrString, AddressOf ContainsValue) Then
            Return True
        End If
        Return False
    End Function
#End Region

#Region "Enalbed TabPages"
    ''' <summary>
    ''' Set thuộc tính Enable = False của Tab page
    ''' </summary>
    ''' <param name="tabPage">tập các tab cần set Enabled = False</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub EnabledTabPage(ByVal tabPage() As TabPage, Optional ByVal bEnabled As Boolean = False)
        Dim tabMain As TabControl = CType(tabPage(0).Parent, TabControl)
        tabMain.DrawMode = TabDrawMode.OwnerDrawFixed
        For i As Integer = 0 To tabPage.Length - 1
            tabPage(i).Enabled = bEnabled

        Next
        For i As Integer = 0 To tabMain.TabPages.Count - 1
            If tabMain.TabPages(i).Enabled Then tabMain.SelectedTab = tabMain.TabPages(i) : Exit For
        Next
        AddHandler tabMain.Selecting, AddressOf tabMain_Selecting
        AddHandler tabMain.DrawItem, AddressOf OnDrawItem
        tabMain.TopLevelControl.Refresh()
    End Sub

    Private Sub tabMain_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs)
        If e.TabPage.Enabled = False Then
            e.Cancel = True
        Else
            e.Cancel = False
        End If
    End Sub

    Private Sub OnDrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        Dim tabMain As TabControl = CType(sender, TabControl)
        ' Create pen.
        Dim blackPen As New Pen(tabMain.TabPages(0).BackColor, 3)
        'Get Location tabpage
        Dim myTabRect As Rectangle
        myTabRect = tabMain.GetTabRect(tabMain.SelectedIndex)
        ' Create coordinates of points that define line.
        Dim x1 As Integer = myTabRect.X
        Dim y1 As Integer = myTabRect.Bottom
        Dim x2 As Integer = myTabRect.X + myTabRect.Width
        ' Draw line to screen.
        e.Graphics.DrawLine(blackPen, x1, y1, x2, y1)
        '**************
        ' Set format of string.
        Dim drawFormat As New StringFormat
        drawFormat.LineAlignment = StringAlignment.Center

        Dim eBounds As System.Drawing.Rectangle
        eBounds.X = e.Bounds.X
        eBounds.Y = e.Bounds.Y
        eBounds.Width = 300
        eBounds.Height = 22
        'If e.Bounds.Width > eBounds.Width Then
        '    eBounds.Width = 300 'e.Bounds.Width 
        'End If
        'If e.Bounds.Height > eBounds.Height Then
        '    eBounds.Height = 22 'e.Bounds.Height
        'End If

        Dim page As TabPage = tabMain.TabPages(e.Index)
        If Not page.Enabled Then
            Dim brush As New SolidBrush(SystemColors.GrayText)
            e.Graphics.DrawString(page.Text, page.Font, brush, eBounds, drawFormat)
        Else
            Dim brush As New SolidBrush(page.ForeColor)
            e.Graphics.DrawString(page.Text, page.Font, brush, eBounds, drawFormat)
        End If

    End Sub

#End Region

#Region "Kiểm tra chuỗi kết nối"

    Public Function CheckConnection() As Boolean
        'Update 16/03/2013: Đưa hàm CheckConnection về Public 
        'Tạo chuỗi kết nối 60 giây 
        gsConnectionString = "Data Source=" & gsServer & ";Initial Catalog=" & gsCompanyID & ";User ID=" & gsConnectionUser & ";Password=" & gsPassword & ";" & gsConnectionTimeout60 'Tạo chuỗi kết nối dùng cho toàn bộ project
        'Kiểm tra chuỗi kết nối
        Try
            gConn = New SqlConnection(gsConnectionString)
            gConn.Open()
            gConn.Close()
            'Update 20/01/2013: trả lại Kết nối không giới hạn thời gian
            gsConnectionString = gsConnectionString.Replace(gsConnectionTimeout60, gsConnectionTimeout)
            Return True
        Catch
            gConn.Close()
            D99C0008.MsgInvalidConnection()
            Return False
        End Try
    End Function
#End Region

#Region "Đồng bộ exe và fix"

    ''' <summary>
    ''' Kiểm tra đồng bộ giữa exe và fix 
    ''' </summary>
    ''' <param name="sExeName">DxxExxxx</param>
    ''' <returns>True: tiếp tục chạy Lemon3; False: kết thúc chương trình</returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function CheckExeFixSynchronous(ByVal sExeName As String) As Boolean
        'Update 05/08/2011: Lấy màu bắt buộc nhập được thiết lập từ Tùy chọn của D00
        GetBackColorObligatory() 'hàm này không liên quan đến Đồng bộ exe, nhưng do exe nào cũng được gọi hàm này nên để chung
        'Update 14/01/2013: Kiểm tra kết nối Link server
        CheckConnectionLinkServer()

        'Đang chạy exe nào thì kiểm tra exe đó

        'Kiểm tra trước khi kiểm tra đồng bộ
        If Not AllowCheckSynchronous() Then Return True

        'Kiểm tra đồng bộ
        If Not CheckVersionUpdate(sExeName) Then
            Return CallExe_Lemon3ServiceUpdate()
        End If

        Return True
    End Function

    'Tạm thời bỏ 2 hàm này: vì fix chưa kiểm tra được.
    '''' <summary>
    '''' Kiểm tra đồng bộ giữa exe và fix của exe cha
    '''' </summary>
    '''' <param name="sArrExeName">tập các exe, VD: Dim arrExe As String() = {"D09E0040", "D09E0140", "D09E0240", "D09E0340"}</param>
    '''' <returns>True: tiếp tục chạy Lemon3; False: kết thúc chương trình</returns>
    '''' <remarks></remarks>
    '<DebuggerStepThrough()> _
    'Public Function CheckExeFixSynchronous(ByVal sArrExeName As String()) As Boolean
    '    'Kiểm tra trước khi kiểm tra đồng bộ
    '    If Not AllowCheckSynchronous() Then Return True

    '    For i As Integer = 0 To sArrExeName.Length - 1
    '        'Kiểm tra đồng bộ cho từng exe, nếu một exe không đồng bộ thì cập nhật lại tất cả các exe cùng module
    '        If Not CheckVersionUpdate(sArrExeName(i)) Then
    '            Return CallExe_Lemon3ServiceUpdate()
    '        End If
    '    Next
    '    Return True
    'End Function

    '''' <summary>
    '''' Kiểm tra đồng bộ giữa exe và fix của exe con
    '''' </summary>
    '''' <param name="sExeName">DxxExxxx</param>
    '''' <param name="sModuleID">Dxx</param>
    '''' <returns>True: tiếp tục chạy Lemon3; False: kết thúc chương trình</returns>
    '''' <remarks></remarks>
    '<DebuggerStepThrough()> _
    'Public Function CheckExeFixSynchronous(ByVal sExeName As String, ByVal sModuleID As String) As Boolean
    '    'Nếu exe con thuộc module của mình thì không gọi kiểm tra đồng bộ
    '    'VD: D09E0040 gọi D09E0140 thì không kiểm tra đồng bộ
    '    'Còn D09E0040 gọi D91E0240 thì kiểm tra đồng bộ

    '    'Kiểm tra trước khi kiểm tra đồng bộ
    '    If Not AllowCheckSynchronous() Then Return True

    '    If sModuleID = "" Then
    '        D99C0008.MsgL3(r("Khong_ton_tai_ma_module"))
    '        Return False
    '    End If

    '    If sModuleID.Length <> 3 Then
    '        D99C0008.MsgL3("Mã module không hợp lệ.")
    '        Return False
    '    End If

    '    If sExeName.Trim.Substring(0, 3) = sModuleID Then Return True

    '    'Kiểm tra đồng bộ
    '    If Not CheckVersionUpdate(sExeName) Then
    '        Return CallExe_Lemon3ServiceUpdate()
    '    End If

    '    Return True
    'End Function

    Private Function AllowCheckSynchronous() As Boolean
        'Quy tắc kiểm tra trước kiểm tra đồng bộ: 
        'Nếu tên máy không chứa "DRD" thì luôn luôn kiểm tra đồng bộ
        'Nếu tên máy có chứa "DRD" thì tiếp tục kiểm tra file D99E0140.exp 
        'Nếu file D99E0140.exp có giá trị = 0 thì không kiểm tra đồng bộ
        'Nếu file D99E0140.exp có giá trị <> 0 thì kiểm tra đồng bộ

        Dim sHostName As String = My.Computer.Name

        'Update 5/1/2011: sửa lỗi cho máy <3 ký tự
        If sHostName.Length >= 3 Then
            If sHostName.Substring(0, 3).ToUpper <> "DRD" Then Return True
        Else
            Return True
        End If

        Dim bCheck As Boolean = False
        Dim sStringFile As String = ReadFile()
        If sStringFile <> "" Then
            sStringFile = sStringFile.Substring(sStringFile.LastIndexOf("=") + 1, 1)
            If sStringFile = "0" Then
                Return False
            Else
                Return True
            End If
        Else
            Return True
        End If

    End Function

    Private Function ReadFile() As String
        Dim fileName As String = "D99E0140.exp"
        Dim fileReader As String
        Dim sPathFile As String = My.Application.Info.DirectoryPath & "\" & fileName

        'Đọc file mặc định của Font hiện tại
        Dim encoding As System.Text.Encoding = System.Text.Encoding.Default
        If Not File.Exists(sPathFile) Then Return ""
        fileReader = My.Computer.FileSystem.ReadAllText(sPathFile, encoding)
        Return fileReader
    End Function

    ''' <summary>
    ''' Kiểm tra đồng bộ giữa exe và fix
    ''' </summary>
    ''' <param name="sExeName">DxxExxxx</param>
    ''' <returns>True: Đã đồng bộ; False: Chưa đồng bộ</returns>
    ''' <remarks></remarks>
    Private Function CheckVersionUpdate(ByVal sExeName As String) As Boolean
        Dim dtExeInfo As New DataTable
        Dim sModule As String = Left(sExeName, 3)

        If ExistRecord("SELECT * FROM DBO.SYSOBJECTS WHERE ID = OBJECT_ID(N'[DBO].[D91T9155]') AND OBJECTPROPERTY(ID, N'ISUSERTABLE') = 1") Then
            dtExeInfo = ReturnDataTable("SELECT ISNULL(MAX(DBUpgradeStatus), 0) AS DBUpgradeStatus FROM D91T9155 WITH (NOLOCK) WHERE Module = '" & sModule & "'")
            If dtExeInfo.Rows.Count > 0 And (dtExeInfo.Rows(0).Item("DBUpgradeStatus").ToString = "1") Then
                D99C0008.MsgL3("Module " + sModule + Space(1) + r("dang_trong_qua_trinh_nang_cap"))
                Return False
            End If
            dtExeInfo = ReturnDataTable("SELECT * FROM D91T9155 WITH(NOLOCK) WHERE EXEName = '" & sExeName & "'")
            If dtExeInfo.Rows.Count <= 0 Then
                dtExeInfo.Dispose()
                Return True
            Else
                Dim aArray1() As String = Split(dtExeInfo.Rows(0).Item("EXERequireLastModifyFixDate").ToString, ";")
                Dim dDateFile As DateTime
                For i As Integer = 0 To aArray1.Length - 1
                    Dim aArray2() As String = Split(aArray1(i), "=")
                    dDateFile = Convert.ToDateTime(ReturnScalar("SELECT isnull (Max(FileDate), '1900/01/01') As MaxFileDate FROM D91T9150 WITH(NOLOCK) WHERE Module = " & SQLString(aArray2(0)) & " And SPBuildDate  = " & SQLDateTimeSave(dtExeInfo.Rows(0).Item("SPBuildDate"))))
                    If aArray2(1) <> Format(dDateFile, "yyyyMMddHHmm") Then
                        dtExeInfo.Dispose()
                        Return False
                    End If
                Next
            End If
        Else
            'Chưa có fix kiểm tra đồng bộ
            dtExeInfo.Dispose()
            Return True
        End If

        Dim FileName As New FileInfo(Application.StartupPath & "\" & sExeName & ".exe")
        If CheckSum(FileName) <> dtExeInfo.Rows(0).Item("EXECheckSum").ToString Then
            dtExeInfo.Dispose()
            Return False
        End If

        dtExeInfo.Dispose()
        Return True
    End Function

    Private Function CheckSum(ByVal FileName As FileInfo) As String
        Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        Dim hash As Byte() = md5.ComputeHash(File.ReadAllBytes(FileName.FullName))
        Return BitConverter.ToString(hash)
    End Function

    Private Function CallExe_Lemon3ServiceUpdate() As Boolean
        If D99C0008.MsgAsk(r("Exe_va_fix_khong_dong_bo_Ban_co_muon_cap_nhat_khong")) = DialogResult.Yes Then
            Dim sFile As String = My.Application.Info.DirectoryPath & "\Lemon3ServiceUpdate.exe"
            If Not File.Exists(sFile) Then
                If geLanguage = EnumLanguage.Vietnamese Then
                    D99C0008.MsgL3("Không tồn tại file " & "Lemon3ServiceUpdate.exe")
                Else
                    D99C0008.MsgL3("Not exist file " & "Lemon3ServiceUpdate.exe")
                End If
                Return True
            End If
            Shell(sFile, AppWinStyle.NormalFocus)
            Return False
        Else
            Return True
        End If
    End Function
#End Region

#Region "Tạo chuỗi kết nối cho Link server"


    Private Function CheckConnectionLinkServer() As Boolean
        'Kiểm tra các kết nối được tạo từ store D00P0020
        Dim sConnectionStringLEMONSYS As String = "Data Source=" & gsServer & ";Initial Catalog=LEMONSYS;User ID=" & gsConnectionUser & ";Password=" & gsPassword & ";" & gsConnectionTimeout60
        Dim sSQL As String = ""
        sSQL &= "EXEC D00P0020 " & SQLString(gsCompanyID) & "," & SQLString(gsUserID)

        Dim conn As SqlConnection = New SqlConnection(sConnectionStringLEMONSYS)
        Dim cmd As SqlCommand = New SqlCommand(sSQL, conn)
        Dim da As SqlDataAdapter = New SqlDataAdapter(cmd)
        Dim ds As DataSet = New DataSet()
        Try
            conn.Open()
            cmd.CommandTimeout = 0
            da.Fill(ds)
            conn.Close()
            Dim dtCon As DataTable = ds.Tables(0)

            For Each dr1 As DataRow In dtCon.Rows
                If dr1("ServerType").ToString = "A" Then
                    gsConnectionStringApp = "Data Source=" & dr1("ServerName").ToString & ";Initial Catalog=" & dr1("CompanyID").ToString & ";User ID=" & dr1("LoginUser").ToString & ";Password=" & dr1("LoginPassword").ToString & ";" & gsConnectionTimeout60 'Tạo chuỗi kết nối dùng cho toàn bộ project
                    gsServerApp = dr1("ServerName").ToString
                    gsCompanyIDApp = dr1("CompanyID").ToString
                ElseIf dr1("ServerType").ToString = "R" Then
                    gsConnectionStringReport = "Data Source=" & dr1("ServerName").ToString & ";Initial Catalog=" & dr1("CompanyID").ToString & ";User ID=" & dr1("LoginUser").ToString & ";Password=" & dr1("LoginPassword").ToString & ";" & gsConnectionTimeout60 'Tạo chuỗi kết nối dùng cho toàn bộ project
                    gsServerReport = dr1("ServerName").ToString
                    gsCompanyIDReport = dr1("CompanyID").ToString
                End If
            Next
            dtCon.Dispose()

            'Nếu có Link kết nối thì kiểm tra kết nối, ngược lại thì gán bằng chuỗi kết nối chuẩn
            If gsConnectionStringApp <> "" Then
                If Not CheckConnect(gsConnectionStringApp) Then
                    gsConnectionStringApp = gsConnectionString
                    gsServerApp = gsServer
                    gsCompanyIDApp = gsCompanyID
                End If
            Else
                gsConnectionStringApp = gsConnectionString
                gsServerApp = gsServer
                gsCompanyIDApp = gsCompanyID
            End If

            If gsConnectionStringReport <> "" Then
                If Not CheckConnect(gsConnectionStringReport) Then
                    gsConnectionStringReport = gsConnectionString
                    gsServerReport = gsServer
                    gsCompanyIDReport = gsCompanyID
                End If

            Else
                gsConnectionStringReport = gsConnectionString
                gsServerReport = gsServer
                gsCompanyIDReport = gsCompanyID
            End If


        Catch
            conn.Close()
            Clipboard.Clear()
            Clipboard.SetText(sSQL)
            MsgErr("Error when excute SQL in function ReturnDataSet(). Paste your SQL code from Clipboard")
            Return Nothing
        End Try

    End Function

    Private Function CheckConnect(ByRef sConnectionString As String) As Boolean
        Try
            gConn = New SqlConnection(sConnectionString)
            gConn.Open()
            gConn.Close()
            sConnectionString = sConnectionString.Replace(gsConnectionTimeout60, gsConnectionTimeout)
            Return True
        Catch
            gConn.Close()
            '            D99C0008.MsgInvalidConnection()
            Return False
        End Try
    End Function
#End Region

#Region "Đánh số TT cho menu"

    Private Sub SetAtoZ(ByVal mnu As C1.Win.C1Command.C1CommandMenu)
        Dim chrcode As Integer = 65 '65->90: A -> Z
        For i As Integer = 0 To mnu.CommandLinks.Count - 1
            If mnu.CommandLinks(i).Visible = False OrElse mnu.CommandLinks(i).Command.Name = "mnuSystemQuit" Then Continue For

            Select Case ChrW(chrcode).ToString
                Case "I"
                    mnu.CommandLinks(i).Command.Text = "&" & ChrW(chrcode) & Space(3) & mnu.CommandLinks(i).Command.Text
                Case Else
                    mnu.CommandLinks(i).Command.Text = "&" & ChrW(chrcode) & Space(2) & mnu.CommandLinks(i).Command.Text
            End Select

            If TypeOf (mnu.CommandLinks(i).Command) Is C1.Win.C1Command.C1CommandMenu Then 'Là menu
                Dim temp As C1.Win.C1Command.C1CommandMenu = CType(mnu.CommandLinks(i).Command, C1.Win.C1Command.C1CommandMenu)
                SetNumber(temp)
            End If
            chrcode += 1
        Next
    End Sub

    Private Sub SetNumber(ByVal mnu As C1.Win.C1Command.C1CommandMenu)
        Dim chrcode As Integer = 1 '1 ->
        For i As Integer = 0 To mnu.CommandLinks.Count - 1
            If mnu.CommandLinks(i).Visible = False Then Continue For
            mnu.CommandLinks(i).Command.Text = "&" & chrcode.ToString & Space(2) & mnu.CommandLinks(i).Command.Text
            mnu.CommandLinks(i).Text = mnu.CommandLinks(i).Command.Text
            chrcode += 1
        Next
    End Sub

    Public Sub SetTextMenu(ByVal menu As C1.Win.C1Command.C1MainMenu)
        For i As Integer = 0 To menu.CommandLinks.Count - 1
            If TypeOf (menu.CommandLinks(i).Command) Is C1.Win.C1Command.C1CommandMenu Then 'Lấy 5 menu chính
                Dim mnu As C1.Win.C1Command.C1CommandMenu = CType(menu.CommandLinks(i).Command, C1.Win.C1Command.C1CommandMenu)
                SetAtoZ(mnu) 'Set menu con trong từng menu chính
            End If
        Next
    End Sub
#End Region

#Region "Phân quyền Menu In"
    'bước 3: phân quyền menu in
    Public Sub GetPrintNumber(ByVal sSQL As String)
        Dim dt1 As DataTable

        giNumberOfPrint = 0
        giPrintNumber = 0
        dt1 = ReturnDataTable(sSQL)

        If dt1.Rows.Count > 0 Then giPrintNumber = L3Int(dt1.Rows(0).Item(0))
    End Sub
#End Region

#Region "Các hàm kiểm tra theo thiết lập của Tài khoản"
    Private Function ReturnTableKcode(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal drTransTypeID As DataRow) As DataTable
        ' Lấy về danh sách Kcode theo thiết lập kiểm tra tại D90
        Dim dtCheckAccount As DataTable = ReturnDataTable("Exec D91P9310")
        For Each dc As DataColumn In dtCheckAccount.Columns
            Try
                If dc.ColumnName = "AccountID" Then Continue For
                If dc.ColumnName = "PeriodID" Then dc.Caption = "Visible" & dc.ColumnName : Continue For
                If drTransTypeID Is Nothing Then
                    If L3Bool(tdbg.Columns(dc.ColumnName).Tag) Then dc.Caption = "Visible" & dc.ColumnName
                Else
                    If L3Bool(tdbg.Columns(dc.ColumnName).Tag) And L3Bool(drTransTypeID.Item("Use" & dc.ColumnName)) Then dc.Caption = "Visible" & dc.ColumnName
                End If

            Catch ex As Exception
                D99C0008.MsgL3("ReturnTableKcode(): " & ex.Message)
            End Try
        Next
        Return dtCheckAccount
    End Function

    Private Function CaptionAna(ByVal sAna As String) As String
        If gbUnicode Then
            Return sAna
        Else
            Return ConvertVniToUnicode(sAna)
        End If
    End Function

    Private Function CheckKCode(ByVal dtCheckAccount As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sAccountID As String, ByVal COL_KCode As String, ByVal iRow As Integer) As Boolean
        ' Chỉ kiểm tra đối với những cột Kcode có hiển thị trên lưới
        If dtCheckAccount.Columns(COL_KCode).Caption = "Visible" & dtCheckAccount.Columns.Item(COL_KCode).ColumnName Then
            ' Ứng với tài khoản sAccountID, tại cột COL_Kcode có được thiết lập kiểm tra hay không 
            Dim dr() As DataRow = dtCheckAccount.Select("AccountID = '" & sAccountID & "'")
            If dr Is Nothing OrElse dr.Length = 0 Then Return True
            Select Case dr(0)(COL_KCode).ToString
                Case "0" ' Không kiểm tra
                    Return True
                Case "1" ' Có thông báo
                    dr(0)(COL_KCode) = "-1" 'Ghi lại đã kiểm tra qua 1 lần khi nhấn lưu
                    dtCheckAccount.AcceptChanges()
                    If D99C0008.MsgAsk(r("Ban_chua_nhap") & Space(1) & CaptionAna(tdbg.Columns(COL_KCode).Caption) & Space(1) & r("cho_tai_khoan") & Space(1) & sAccountID & vbCrLf & r("Ban_co_muon_nhap_khong")) = Windows.Forms.DialogResult.Yes Then
                        Return False
                    Else
                        Return True
                    End If
                Case "2" ' Bắt buộc nhập
                    D99C0008.MsgL3(r("Ban_phai_nhap") & Space(1) & CaptionAna(tdbg.Columns(COL_KCode).Caption) & Space(1) & r("cho_tai_khoan") & Space(1) & sAccountID, L3MessageBoxIcon.Exclamation)
                    Return False
            End Select
        End If

        Return True
    End Function

    Public Function AllowKCode(ByRef dtCheckAccount As DataTable, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iRow As Integer, ByVal drTransTypeID As DataRow, ByRef COL_KCode As String, ByVal ParamArray sAccountID() As String) As Boolean
        If dtCheckAccount Is Nothing Then dtCheckAccount = ReturnTableKcode(tdbg, drTransTypeID)
        For Each dc As DataColumn In dtCheckAccount.Columns
            Try
                If dc.ColumnName = "AccountID" OrElse (IsDBNull(tdbg(iRow, dc.ColumnName)) = False And tdbg(iRow, dc.ColumnName).ToString.Trim <> "") Then Continue For
                COL_KCode = dc.ColumnName
                For i As Integer = 0 To sAccountID.Length - 1
                    ' Chỉ kiểm tại cell trên lưới chưa có giá trị
                    If Not CheckKCode(dtCheckAccount, tdbg, sAccountID(i), dc.ColumnName, iRow) Then Return False
                Next
            Catch ex As Exception

            End Try
        Next
        Return True
    End Function
#End Region

#Region "Kiểm tra xuyên kỳ"
    Public Function CheckThroughPeriod(ByVal sTranMonth As String, ByVal sTranYear As String) As Boolean

        'If tdbg.Columns(COL_Period).Text <> Format(giTranMonth, "00") & "/" & Format(giTranYear, "0000") Then '
        If giTranMonth <> Number(sTranMonth) Or giTranYear <> Number(sTranYear) Then
            D99C0008.MsgL3(r("MSG000001"))
            Return False
        End If
        Return True
    End Function

    Public Function CheckThroughPeriod(ByVal sPeriod As String) As Boolean
        Return CheckThroughPeriod(Strings.Left(sPeriod, 2), Strings.Right(sPeriod, 4))
    End Function
#End Region

#Region "Grid"

    Public Sub LockColums(ByVal bLock As Boolean, ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal isplit As Integer, ByVal ParamArray arrcol() As String)
        If bLock Then
            LockColums(tdbg, isplit, arrcol)
        Else
            UnLockColums(tdbg, isplit, arrcol)
        End If
    End Sub

    Public Sub LockColums(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal isplit As Integer, ByVal ParamArray arrcol() As String)
        For Each col As String In arrcol
            tdbg.Splits(isplit).DisplayColumns(col).Button = False
            tdbg.Splits(isplit).DisplayColumns(col).Locked = True
            tdbg.Splits(isplit).DisplayColumns(col).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        Next
    End Sub

    Public Sub UnLockColums(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal isplit As Integer, ByVal ParamArray arrcol() As String)
        For Each col As String In arrcol
            If tdbg.Columns(col).DropDown IsNot Nothing Then tdbg.Splits(isplit).DisplayColumns(col).Button = True
            tdbg.Splits(isplit).DisplayColumns(col).Locked = False
            tdbg.Splits(isplit).DisplayColumns(col).AllowFocus = True
            tdbg.Splits(isplit).DisplayColumns(col).Style.ResetBackColor()
        Next
    End Sub

    Public Function findrowInGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sValueFind As Object, ByVal sColName As String) As Integer
        ' get the currency manager that the grid is bound to
        Dim cm As CurrencyManager = CType(CType(tdbg.TopLevelControl, Form).BindingContext(tdbg.DataSource, tdbg.DataMember), CurrencyManager)
        ' get the property descriptor for the "integer" column
        Dim prop As System.ComponentModel.PropertyDescriptor = cm.GetItemProperties()(sColName)

        ' get the binding list
        Dim blist As System.ComponentModel.IBindingList = CType(cm.List, System.ComponentModel.IBindingList)

        ' find the newly added record
        Return blist.Find(prop, sValueFind)
    End Function '_findrow
#End Region

#Region "Load default EmployeeID"
    Public Sub GetTextCreateByNew(ByVal tdbc As C1.Win.C1List.C1Combo, Optional ByVal bDefault As Boolean = True)
        If bDefault Then tdbc.SelectedValue = gsCreateBy
        'Update 31/07/2012: Kiểm tra Khóa người dùng Lemon3
        'Nếu D91 thiết lập "Khóa người dùng Lemon3" (gbLockL3UserID = True) và combo Người lập có giá trị thì Lock Combo lại.
        If gbLockL3UserID And tdbc.SelectedValue IsNot Nothing Then tdbc.ReadOnly = (tdbc.Text <> "")
    End Sub

    Public Sub GetTextCreateByNew(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal COL_CreatorID As Integer, ByVal iSplit As Integer)
        tdbg.Columns(COL_CreatorID).DefaultValue = gsCreateBy
        If gbLockL3UserID And ReturnValueC1DropDown(tdbg.Columns(COL_CreatorID).DropDown, "EmployeeID", "EmployeeID=" & SQLString(gsCreateBy)) <> "" Then
            tdbg.Splits(iSplit).DisplayColumns(COL_CreatorID).Locked = True
            tdbg.Splits(iSplit).DisplayColumns(COL_CreatorID).Button = False
            tdbg.Splits(iSplit).DisplayColumns(COL_CreatorID).AutoDropDown = False
            tdbg.Splits(iSplit).DisplayColumns(COL_CreatorID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        End If
    End Sub
#End Region


#Region "Thông báo chung"
    'Update 10/1/2014: incident 62703 Bỏ Tùy chọn của mỗi module

    ''' <summary>
    ''' Thông báo trước khi xóa: update 10/1/2014: incident 62703 Bỏ Tùy chọn của mỗi module
    ''' </summary>    
    Public Function AskDelete() As DialogResult
        '    If D08Options.MessageAskBeforeSave Then
        '        Return D99C0008.MsgAskDelete
        '    Else
        '        Return DialogResult.Yes
        '    End If
        Return D99C0008.MsgAskDelete
    End Function


    ''' <summary>
    ''' Thông báo sau khi xóa thành công
    ''' </summary>
    Public Sub DeleteOK()
        'If D08Options.MessageWhenSaveOK Then D99C0008.MsgL3(r("MSG000008"))
        D99C0008.MsgL3(r("MSG000008"))
    End Sub

    ''' <summary>
    ''' Thông báo trước khi khóa phiếu
    ''' </summary>    
    Public Function AskLocked() As DialogResult
        Return D99C0008.MsgAsk(r("MSG000002"), MessageBoxDefaultButton.Button2)

    End Function


    ''' <summary>
    ''' Thông báo sau khi khóa phiếu thành công
    ''' </summary>
    Public Sub LockedOK()
        D99C0008.MsgSaveOK() 'MsgL3(r("Khoa_phieu_thanh_cong"))
    End Sub


    ''' <summary>
    ''' Thông báo không lưu được dữ liệu
    ''' </summary>
    Public Sub SaveNotOK()
        D99C0008.MsgSaveNotOK()
    End Sub


    Public Sub MyMsg(ByVal strMsg As String)
        D99C0008.MsgL3(strMsg, L3MessageBoxIcon.Information)
    End Sub
#End Region


End Module
