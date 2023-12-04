Imports System
'#######################################################################################
'#                                     CHÚ Ý (New)
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 13/11/2013
'# Người cập nhật cuối cùng: MINHHOA
'# Diễn giải: các vấn đề về form Xuất Excel (D99F2222) và Nút Hiển thị(F12) trên lưới
'# Bổ sung phần DataType
'# Bổ sung hàm ResetTableForExcel cho trường hợp định dạng trên cột "Nx"
'# Bổ sung hàm AddRowStyleValue và AddCellStyleValue cho trường hợp Xuất Excel theo định dạng lưới
'# Bổ sung hàm CreateColStyeleExcel để thêm cột StyleExcel để xuất Excel theo định dạng lưới
'# Bổ sung set màu của Phương viết
'#######################################################################################

Module D99X2222
    Public Structure VisibleColumn
        Public ColFieldName As String
        Public ColVisible As Boolean
    End Structure
    Public vcNew(-1, -1) As VisibleColumn
    Public matrix(9, 1) As Integer

    Public giMaxLengthColumnCaption As Integer = 0 'Ghi nhận Max Length của Caption cột
    Public gdtCaptionExcel As DataTable
    Public gsGroupColumns As String() 'Mảng nhóm cột được chọn
    Public giRefreshUserControl As Integer = -1 'Trạng thái -1: chưa thay đổi, 0: có thay đổi, 1: đã refresh
    Public Const COL_StyleExcel As String = "StyleExcel"

    Public Structure ShowColumn
        Public FieldName As String
        Public NumberFormat As String
        Public DataType As String
        Public Caption As String
        Public Obligatory As Byte ' 1: Những cột bắt buộc nhập
        Public Grouped As Byte
        Public IsSum As Byte ' 1: Những cột Số có Sum
        Public IsDateTime As Byte ' 1: Ngày định dạng theo dd/MM/yyyy hh:mm:ss
        Public DataWidth As Integer ' Chiều dài dữ liệu bắt buộc nhập

    End Structure

    ''''' <summary>
    ''''' gán lại các thuộc tính cho table caption Xuất excel sau khi đã đổ nguồn vào lưới
    ''''' </summary>
    ''''' <param name="c1Grid"></param>
    ''''' <param name="dtCaption"></param>
    ''''' <param name="arrColSum">mảng các cột có tính tổng</param>
    ''''' <param name="arrColDateLong">mảng các cột có ngày định dạng dd/MM/yyyy hh:mm:ss</param>
    ''''' <remarks></remarks>
    '<DebuggerStepThrough()> _
    Public Sub ResetTableForExcel(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dtCaption As DataTable, Optional ByVal arrColSum() As Integer = Nothing, Optional ByVal arrColDateLong() As Integer = Nothing)
        Dim arrGroupColumns(c1Grid.GroupedColumns.Count - 1) As String
        'Set lại DataType và NumberFormat cho Table Load Caption cho xuất Excel
        For i As Integer = 0 To dtCaption.Rows.Count - 1
            Dim dr As DataRow = dtCaption.Rows(i)
            For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In c1Grid.Columns

                'dc.ValueItems.Presentation=C1.Win.C1TrueDBGrid.PresentationEnum.Normal
                If dc.DataField = dr("FieldName").ToString Then
                    'Cần phân biệt kiểu số để đưa vào tìm kiếm: N1(TinyInt); N2 (Int); N (Các kiểu số còn lại)
                    'Update 13/07/2010: Byte có thể không là Checkbox, VD: cột TranMonth trên lưới
                    'If dc.DataType.Name = "Boolean" OrElse dc.DataType.Name = "Byte" Then 'Kiểu số dạng Checkbox (chỉ gõ số, không dấu - dấu .)
                    If dc.ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then
                        dr("DataType") = "N1"
                    ElseIf dc.DataType.Name.Contains("Int") OrElse dc.DataType.Name = "Byte" Then 'Kiểu số không format (có dấu - không dấu .)
                        dr("DataType") = "N2"
                        'Kiểm tra cột nào có tính tổng
                        If arrColSum IsNot Nothing AndAlso arrColSum.Length > 0 Then
                            If L3FindInteger(arrColSum, c1Grid.Columns.IndexOf(dc)) Then
                                dr("IsSum") = "1"
                            Else
                                dr("IsSum") = "0"
                            End If
                        Else
                            dr("IsSum") = IIf(IsNumeric(dc.FooterText), "1", "0")
                        End If
                    ElseIf dc.DataType.Name = "Decimal" OrElse dc.DataType.Name = "Double" OrElse dc.DataType.Name = "Single" Then 'Kiểu số có format (có dấu - có dấu .)
                        dr("DataType") = "N"
                        If dc.DataType.Name.Contains("Decimal") Then 'Kiểu số->lấy Format
                            Dim sFormat As String = dc.NumberFormat
                            If sFormat IsNot Nothing AndAlso sFormat <> "" Then
                                If sFormat = "Percent" Or sFormat.Contains("%") Then  'Thêm trường hợp format Percent - Update 06/11/2012
                                    dr("DataType") = "Percent"
                                ElseIf sFormat.Contains("Event") Then
                                    dr("NumberFormat") = ReturnNumDigits(dc.Tag)
                                ElseIf sFormat.Contains(".") Then
                                    Dim arr() As String = sFormat.Split(CType(".", Char))
                                    dr("NumberFormat") = arr(arr.Length - 1).Length
                                ElseIf sFormat.Contains("N") Then 'Bổ sung 07/10/2011 TH định dạng trên cột "Nx"
                                    dr("NumberFormat") = sFormat.Substring(1, sFormat.Length - 1)
                                ElseIf IsNumeric(sFormat) Then
                                    dr("NumberFormat") = sFormat
                                End If
                            End If
                        End If

                        'Kiểm tra cột nào có tính tổng
                        If arrColSum IsNot Nothing AndAlso arrColSum.Length > 0 Then
                            If L3FindInteger(arrColSum, c1Grid.Columns.IndexOf(dc)) Then
                                dr("IsSum") = "1"
                            Else
                                dr("IsSum") = "0"
                            End If
                        Else
                            dr("IsSum") = IIf(IsNumeric(dc.FooterText), "1", "0")
                        End If

                    ElseIf dc.DataType.Name = "DateTime" Then 'Kiểu ngày
                        dr("DataType") = "D"
                        'Kiểm tra cột nào có dạng ngày
                        If arrColDateLong IsNot Nothing AndAlso arrColDateLong.Length > 0 Then
                            If L3FindInteger(arrColDateLong, c1Grid.Columns.IndexOf(dc)) Then
                                dr("IsDateTime") = "1"
                            Else
                                dr("IsDateTime") = "0"
                            End If
                        Else
                            dr("IsDateTime") = "0"
                        End If

                    Else 'Kiểu chuỗi
                        dr("DataType") = "S"
                    End If

                    'Set Group 
                    Dim bGroup As Boolean = False
                    For j As Integer = 0 To c1Grid.GroupedColumns.Count - 1
                        If c1Grid.GroupedColumns(j).Caption = dr("Description").ToString OrElse c1Grid.GroupedColumns(j).DataField = dr("FieldName").ToString Then
                            'Ghi lại mảng các cột được Group để đưa qua Excel
                            arrGroupColumns(j) = dr("FieldName").ToString
                            dr("Grouped") = 1
                            bGroup = True
                            Exit For
                        End If
                    Next

                    If Not bGroup Then dr("Grouped") = 0
                    Exit For
                End If
            Next
        Next

        gsGroupColumns = arrGroupColumns
    End Sub

    '''' <summary>
    '''' gán lại các thuộc tính cho table caption Xuất excel sau khi đã đổ nguồn vào lưới (Cột trên lưới định nghĩa dạng String)
    '''' </summary>
    '''' <param name="c1Grid"></param>
    '''' <param name="dtCaption"></param>
    '''' <param name="arrColSum">mảng các cột có tính tổng</param>
    '''' <param name="arrColDateLong">mảng các cột có ngày định dạng dd/MM/yyyy hh:mm:ss</param>
    '''' <remarks>Hàm mới</remarks>
    <DebuggerStepThrough()> _
    Public Sub ResetTableForExcel_String(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dtCaption As DataTable, Optional ByVal arrColSum() As String = Nothing, Optional ByVal arrColDateLong() As String = Nothing)
        Dim arrGroupColumns(c1Grid.GroupedColumns.Count - 1) As String
        'Set lại DataType và NumberFormat cho Table Load Caption cho xuất Excel
        For i As Integer = 0 To dtCaption.Rows.Count - 1
            Dim dr As DataRow = dtCaption.Rows(i)
            For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In c1Grid.Columns
                If dc.DataField = dr("FieldName").ToString Then
                    'Cần phân biệt kiểu số để đưa vào tìm kiếm: N1(TinyInt); N2 (Int); N (Các kiểu số còn lại)
                    'Update 13/07/2010: Byte có thể không là Checkbox, VD: cột TranMonth trên lưới
                    'If dc.DataType.Name = "Boolean" OrElse dc.DataType.Name = "Byte" Then 'Kiểu số dạng Checkbox (chỉ gõ số, không dấu - dấu .)
                    If dc.ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then 'Kiểu số dạng Checkbox
                        dr("DataType") = "N1"
                    ElseIf dc.DataType.Name.Contains("Int") OrElse dc.DataType.Name = "Byte" Then 'Kiểu số không format (có dấu - không dấu .)
                        dr("DataType") = "N2"
                        'Kiểm tra cột nào có tính tổng
                        If arrColSum IsNot Nothing AndAlso arrColSum.Length > 0 Then
                            If L3FindArrString(arrColSum, dc.DataField) Then
                                dr("IsSum") = "1"
                            Else
                                dr("IsSum") = "0"
                            End If
                        Else
                            dr("IsSum") = IIf(IsNumeric(dc.FooterText), "1", "0")
                        End If
                    ElseIf dc.DataType.Name = "Decimal" OrElse dc.DataType.Name = "Double" OrElse dc.DataType.Name = "Single" Then 'Kiểu số có format (có dấu - có dấu .)
                        dr("DataType") = "N"
                        If dc.DataType.Name.Contains("Decimal") Then 'Kiểu số->lấy Format
                            Dim sFormat As String = dc.NumberFormat

                            If sFormat IsNot Nothing AndAlso sFormat <> "" Then
                                If sFormat = "Percent" Or sFormat.Contains("%") Then  'Thêm trường hợp format Percent - Update 06/11/2012
                                    dr("DataType") = "Percent"
                                ElseIf sFormat.Contains("Event") Then
                                    dr("NumberFormat") = dc.Tag
                                ElseIf sFormat.Contains(".") Then
                                    Dim arr() As String = sFormat.Split(CType(".", Char))
                                    dr("NumberFormat") = arr(arr.Length - 1).Length
                                ElseIf sFormat.Contains("N") Then 'Bổ sung 07/10/2011 TH định dạng trên cột "Nx"
                                    dr("NumberFormat") = sFormat.Substring(1, sFormat.Length - 1)
                                End If
                            End If
                        End If


                        'Kiểm tra cột nào có tính tổng
                        If arrColSum IsNot Nothing AndAlso arrColSum.Length > 0 Then
                            If L3FindArrString(arrColSum, dc.DataField) Then
                                dr("IsSum") = "1"
                            Else
                                dr("IsSum") = "0"
                            End If
                        Else
                           dr("IsSum") = IIf(IsNumeric(dc.FooterText), "1", "0")
                        End If

                        ElseIf dc.DataType.Name = "DateTime" Then 'Kiểu ngày
                            dr("DataType") = "D"
                            'Kiểm tra cột có dạng ngày nào
                            If arrColDateLong IsNot Nothing AndAlso arrColDateLong.Length > 0 Then
                                If L3FindArrString(arrColDateLong, dc.DataField) Then
                                    dr("IsDateTime") = "1"
                                Else
                                    dr("IsDateTime") = "0"
                                End If
                            Else
                                dr("IsDateTime") = "0"
                            End If

                        Else 'Kiểu chuỗi
                            dr("DataType") = "S"
                        End If

                        'Set Group 
                        Dim bGroup As Boolean = False
                        'Chú ý: không dùng (c1Grid.GroupedColumns.IndexOf(dc) <> -1) để kiểm tra, những cột Group có thể cùng tên Caption nhưng khác FieldName
                        For j As Integer = 0 To c1Grid.GroupedColumns.Count - 1
                            'Những cột Group có thể cùng tên Caption nhưng khác FieldName (VD: Level01, ...) 
                            If c1Grid.GroupedColumns(j).Caption = dr("Description").ToString OrElse c1Grid.GroupedColumns(j).DataField = dr("FieldName").ToString Then
                                'Ghi lại mảng các cột được Group để đưa qua Excel, chú ý phải đúng thứ tự trên Group của lưới
                                arrGroupColumns(j) = dr("FieldName").ToString
                                dr("Grouped") = 1
                                bGroup = True
                                Exit For
                            End If
                        Next

                        If Not bGroup Then dr("Grouped") = 0

                        Exit For
                End If
            Next
        Next

        gsGroupColumns = arrGroupColumns
    End Sub


    ''' <summary>
    ''' Đưa vào mảng ar các cột Hiển thị trên lưới
    ''' </summary>
    ''' <param name="c1Grid"></param>
    ''' <param name="iSplit"></param>
    ''' <param name="ar"></param>
    ''' <param name="ArrColObl"></param>
    ''' <param name="bCheckGroup">chỉ bật True khi lưới không gọi usercontrol D09U1111</param>
    ''' <remarks></remarks>
    '<DebuggerStepThrough()> _
    Public Sub AddColVisible(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal iSplit As Integer, ByRef ar As ArrayList, _
    Optional ByVal ArrColObl() As Integer = Nothing, Optional ByVal bCheckGroup As Boolean = False, Optional ByVal bCheckExistColumn As Boolean = False, _
    Optional ByVal bUseUnicode As Boolean = False)
        'Biến bCheckGroup chưa được dùng
        Try
            For i As Integer = 0 To c1Grid.Columns.Count - 1
                Dim dc As C1.Win.C1TrueDBGrid.C1DataColumn = c1Grid.Columns(i)
                If c1Grid.Splits(iSplit).DisplayColumns(i).Visible Then
                    Dim e As New ShowColumn
                    e.FieldName = dc.DataField
                    'Kiểm tra Cột nào bắt buộc nhập
                    If ArrColObl IsNot Nothing AndAlso ArrColObl.Length > 0 Then
                        'Dạng số
                        If L3FindInteger(ArrColObl, i) Then
                            e.Obligatory = 1 ' Bắt buộc nhập
                        Else
                            e.Obligatory = 0 ' Không bắt buộc nhập
                        End If
                    Else
                        e.Obligatory = 0 ' Không bắt buộc nhập
                    End If

                    If bUseUnicode Then
                        If c1Grid.Splits(iSplit).DisplayColumns(i).HeadingStyle.Font.Name = "Microsoft Sans Serif" Then
                            e.Caption = dc.Caption
                        Else
                            e.Caption = ConvertVniToUnicode(dc.Caption)
                        End If
                    Else
                        If c1Grid.Splits(iSplit).DisplayColumns(i).HeadingStyle.Font.Name = "Microsoft Sans Serif" Then
                            e.Caption = ConvertUnicodeToVni(dc.Caption)
                        Else
                            e.Caption = dc.Caption
                        End If
                    End If

                    'Không có ý nghĩa khi chưa Load grid
                    If dc.DataType.Name = "Decimal" OrElse dc.DataType.Name = "Double" OrElse dc.DataType.Name = "Single" Then 'Kiểu số có format (có dấu - có dấu .)
                        If dc.Tag Is Nothing Then dc.Tag = 0 'Update 13/11/2013
                        e.NumberFormat = IIf(dc.NumberFormat.Contains("Event"), dc.Tag, dc.NumberFormat).ToString 'dc.NumberFormat
                        e.IsSum = CByte(IIf(IsNumeric(dc.FooterText), 1, 0))
                    Else
                        e.NumberFormat = ""
                        e.IsSum = 0
                    End If
                    e.Grouped = 0
                    e.IsDateTime = 0
                    e.DataType = dc.DataType.Name
                    e.DataWidth = dc.DataWidth
                    ar.Add(e)
                End If
            Next i

            If bCheckGroup Then ' add vào thêm những cột Group cho TH lưới không có nút Hiển thị
                For i As Integer = 0 To c1Grid.GroupedColumns.Count - 1
                    Dim dc As C1.Win.C1TrueDBGrid.C1DataColumn = c1Grid.Columns(c1Grid.GroupedColumns(i).DataField)

                    Dim e As New ShowColumn
                    e.FieldName = dc.DataField

                    'Kiểm tra Cột nào bắt buộc nhập
                    e.Obligatory = 1 ' Bắt buộc nhập

                    If bUseUnicode Then
                        If c1Grid.Splits(iSplit).DisplayColumns(i).HeadingStyle.Font.Name = "Microsoft Sans Serif" Then
                            e.Caption = dc.Caption
                        Else
                            e.Caption = ConvertVniToUnicode(dc.Caption)
                        End If
                    Else
                        If c1Grid.Splits(iSplit).DisplayColumns(i).HeadingStyle.Font.Name = "Microsoft Sans Serif" Then
                            e.Caption = ConvertUnicodeToVni(dc.Caption)
                        Else
                            e.Caption = dc.Caption
                        End If
                    End If

                    'Không có ý nghĩa khi chưa Load grid
                    If dc.DataType.ToString.Contains("Decimal") Then
                        e.NumberFormat = IIf(dc.NumberFormat.Contains("Event"), dc.Tag, dc.NumberFormat).ToString 'dc.NumberFormat
                    Else
                        e.NumberFormat = ""
                    End If
                    e.Grouped = 1
                    e.IsSum = 0
                    e.IsDateTime = 0
                    e.DataType = dc.DataType.Name
                    e.DataWidth = dc.DataWidth
                    ar.Add(e)

                Next
            End If
        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' Đưa vào mảng ar các cột Hiển thị trên lưới (Cột trên lưới định nghĩa dạng String)
    ''' </summary>
    ''' <param name="ArrTemplate"></param>
    ''' <param name="c1Grid"></param>
    ''' <param name="ar"></param>
    ''' <param name="iSplit"></param>
    ''' <param name="ArrColObl"></param>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Sub AddColVisible_String(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef ar As ArrayList, Optional ByVal iSplit As Integer = 0, _
    Optional ByVal ArrColObl() As String = Nothing, Optional ByVal ArrTemplate As ArrayList = Nothing, Optional ByVal bCheckExistColumn As Boolean = False, _
    Optional ByVal bUseUnicode As Boolean = False)
        Try
            'Kiểm tra có Mẫu template không
            If ArrTemplate IsNot Nothing AndAlso ArrTemplate.Count > 0 Then ' Có mẫu template
                'Add các cột vào mảng Ar theo mảng ArrTemplale có sẵn được thiết lập
                For i As Integer = 0 To ArrTemplate.Count - 1
                    For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In c1Grid.Columns
                        If ArrTemplate(i).ToString = dc.DataField.ToString Then
                            Dim e As New ShowColumn
                            e.FieldName = dc.DataField

                            If bUseUnicode Then
                                If c1Grid.Splits(iSplit).DisplayColumns(dc.DataField).HeadingStyle.Font.Name = "Microsoft Sans Serif" Then
                                    e.Caption = dc.Caption
                                Else
                                    e.Caption = ConvertVniToUnicode(dc.Caption)
                                End If
                            Else
                                If c1Grid.Splits(iSplit).DisplayColumns(dc.DataField).HeadingStyle.Font.Name = "Microsoft Sans Serif" Then
                                    e.Caption = ConvertUnicodeToVni(dc.Caption)
                                Else
                                    e.Caption = dc.Caption
                                End If
                            End If

                            'Kiểm tra có Group không
                            e.Grouped = CByte(IIf(CheckInGroup(c1Grid, dc), 1, 0))
                            'Kiểm tra Cột nào bắt buộc nhập
                            If e.Grouped = 1 Then ' Nếu là Group thì bắt buộc nhập
                                e.Obligatory = 1
                            Else
                                e.Obligatory = 0
                                If ArrColObl IsNot Nothing AndAlso ArrColObl.Length > 0 Then
                                    If L3FindArrString(ArrColObl, dc.DataField) Then e.Obligatory = 1 ' Bắt buộc nhập
                                End If

                            End If

                            'Không có ý nghĩa khi chưa Load grid
                            'Bỏ: không dùng đinh dạng
                            'Select Case c1Grid.Splits(iSplit).DisplayColumns(dc.DataField).Style.HorizontalAlignment
                            '    Case C1.Win.C1TrueDBGrid.AlignHorzEnum.Far, C1.Win.C1TrueDBGrid.AlignHorzEnum.Justify
                            '        e.NumberFormat = dc.NumberFormat
                            '    Case Else
                            '        e.NumberFormat = ""
                            'End Select
                            If dc.DataType.Name = "Decimal" OrElse dc.DataType.Name = "Double" OrElse dc.DataType.Name = "Single" Then 'Kiểu số có format (có dấu - có dấu .)
                                e.NumberFormat = IIf(dc.NumberFormat.Contains("Event"), dc.Tag, dc.NumberFormat).ToString ' dc.NumberFormat
                            Else
                                e.NumberFormat = ""
                            End If

                            e.IsSum = 0
                            e.IsDateTime = 0

                            e.DataType = dc.DataType.Name
                            ar.Add(e)

                            Exit For
                        End If
                    Next
                Next

            Else ' Không có mẫu template
                For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In c1Grid.Columns
                    If c1Grid.Splits(iSplit).DisplayColumns(dc.DataField).Visible Then
                        Dim e As New ShowColumn
                        e.FieldName = dc.DataField
                        e.Grouped = CByte(IIf(c1Grid.GroupedColumns.IndexOf(dc) <> -1, 1, 0))

                        'Kiểm tra Cột nào bắt buộc nhập
                        e.Obligatory = 0
                        If ArrColObl IsNot Nothing AndAlso ArrColObl.Length > 0 Then
                            If L3FindArrString(ArrColObl, dc.DataField) Then e.Obligatory = 1 ' Bắt buộc nhập
                        End If

                        If bUseUnicode Then
                            If c1Grid.Splits(iSplit).DisplayColumns(dc.DataField).HeadingStyle.Font.Name = "Microsoft Sans Serif" Then
                                e.Caption = dc.Caption
                            Else
                                e.Caption = ConvertVniToUnicode(dc.Caption)
                            End If
                        Else
                            If c1Grid.Splits(iSplit).DisplayColumns(dc.DataField).HeadingStyle.Font.Name = "Microsoft Sans Serif" Then
                                e.Caption = ConvertUnicodeToVni(dc.Caption)
                            Else
                                e.Caption = dc.Caption
                            End If
                        End If

                        'Không có ý nghĩa khi chưa Load grid
                        'Select Case c1Grid.Splits(iSplit).DisplayColumns(dc.DataField).Style.HorizontalAlignment
                        '    Case C1.Win.C1TrueDBGrid.AlignHorzEnum.Far, C1.Win.C1TrueDBGrid.AlignHorzEnum.Justify
                        '        e.NumberFormat = dc.NumberFormat
                        '    Case Else
                        '        e.NumberFormat = ""
                        'End Select
                        'Bỏ: không dùng đinh dạng
                        If dc.DataType.Name = "Decimal" OrElse dc.DataType.Name = "Double" OrElse dc.DataType.Name = "Single" Then 'Kiểu số có format (có dấu - có dấu .)
                            e.NumberFormat = IIf(dc.NumberFormat.Contains("Event"), dc.Tag, dc.NumberFormat).ToString 'dc.NumberFormat
                        Else
                            e.NumberFormat = ""
                        End If

                        e.IsSum = 0
                        e.IsDateTime = 0

                        e.DataType = dc.DataType.Name
                        ar.Add(e)
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub


    Private Function CheckInGroup(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dcRoot As C1.Win.C1TrueDBGrid.C1DataColumn) As Boolean
        'Thiên Huỳnh Edit 23/09/2010
        'Kiểm tra Caption hiện tại có phải là Group (vì Caption giống nhau nhưng FieldName khác nhau)
        For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In c1Grid.GroupedColumns
            If dc.Caption.Trim = dcRoot.Caption.Trim Or dc.DataField.ToString = dcRoot.DataField.ToString Then
                Return True
            End If
        Next
        Return False
    End Function

    'Private Function CheckInGroup(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal sCaptionName As String) As Boolean
    '    'Kiểm tra Caption hiện tại có phải là Group (vì Caption giống nhau nhưng FieldName khác nhau)
    '    For Each dc As C1.Win.C1TrueDBGrid.C1DataColumn In c1Grid.GroupedColumns
    '        If dc.Caption = sCaptionName Then
    '            Return True
    '        End If
    '    Next
    '    Return False

    'End Function


#Region "Create table load Grid for .NET"

    ''' <summary>
    ''' Tạo Table Caption cho Grid Export Excel 
    ''' </summary>
    ''' <param name="C1Grid">Lưới của form truyền vào</param>
    ''' <param name="arrColVisible">mảng các cột được hiển thị trên lưới</param>
    ''' <param name="arrColSum">mảng các cột có tính tổng</param>
    ''' <param name="arrColDateLong">mảng các cột có ngày định dạng dd/MM/yyyy hh:mm:ss</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ' <DebuggerStepThrough()> _
    Public Function CreateTableForExcelOnly(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal arrColVisible As ArrayList, Optional ByVal arrColSum() As Integer = Nothing, Optional ByVal arrColDateLong() As Integer = Nothing) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("FieldName", GetType(System.String))
        dt.Columns.Add("Description", GetType(System.String))
        dt.Columns.Add("OrderNum", GetType(System.Int32))
        dt.Columns.Add("OrderNo", GetType(System.Int32))
        dt.Columns.Add("DataType", GetType(System.String))
        dt.Columns.Add("IsUsed", GetType(System.Boolean))
        dt.Columns.Add("IsUnicode", GetType(System.Boolean))
        dt.Columns.Add("NumberFormat", GetType(System.Byte))
        dt.Columns.Add("Obligatory", GetType(System.Byte))
        dt.Columns.Add("Grouped", GetType(System.Byte))
        dt.Columns.Add("IsSum", GetType(System.Byte))
        dt.Columns.Add("IsDateTime", GetType(System.Byte))
        dt.Columns.Add("IsExport", GetType(System.Byte))
        dt.Columns.Add("DataWidth", GetType(System.Int32))

        Dim dr As DataRow
        Dim iCount As Integer = 0

        For i As Integer = 0 To arrColVisible.Count - 1
            iCount += 1
            dr = dt.NewRow
            Dim e As ShowColumn = CType(arrColVisible(i), ShowColumn)
            dr("FieldName") = e.FieldName
            dr("Description") = e.Caption
            dr("Obligatory") = e.Obligatory
            dr("Grouped") = e.Grouped
            dr("IsSum") = e.IsSum
            dr("IsDateTime") = e.IsDateTime
            dr("IsExport") = 0
            dr("IsUsed") = 1
            dr("IsUnicode") = 0
            dr("DataWidth") = e.DataWidth

            'Lay do dai max
            If dr("Description").ToString.Length > giMaxLengthColumnCaption Then giMaxLengthColumnCaption = dr("Description").ToString.Length

            dr("OrderNum") = iCount
            dr("OrderNo") = 0

            'Có ý nghĩa khi grid đã được load nguồn
            dr("NumberFormat") = 0
            'Update 26/03/2010
            'Cần phân biệt kiểu số để đưa vào tìm kiếm: N1(TinyInt); N2 (Int); N (Các kiểu số còn lại)
            'Update 13/07/2010: Byte có thể không là Checkbox, VD: cột TranMonth trên lưới
            'If e.DataType = "Boolean" OrElse e.DataType = "Byte" Then 'Kiểu số dạng Checkbox (chỉ gõ số, không dấu - dấu .)
            If C1Grid.Columns(e.FieldName).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then 'Kiểu số dạng Checkbox
                dr("DataType") = "N1"
            ElseIf e.DataType.Contains("Int") OrElse e.DataType = "Byte" Then 'Kiểu số không format (có dấu - không dấu .)
                dr("DataType") = "N2"
                'Kiểm tra cột nào có tính tổng
                If arrColSum IsNot Nothing AndAlso arrColSum.Length > 0 Then
                    If L3FindInteger(arrColSum, C1Grid.Columns.IndexOf(C1Grid.Columns(e.FieldName))) Then
                        dr("IsSum") = "1"
                    Else
                        dr("IsSum") = "0"
                    End If
                End If
            ElseIf e.DataType = "Decimal" OrElse e.DataType = "Double" OrElse e.DataType = "Single" Then 'Kiểu số có format (có dấu - có dấu .)
                dr("DataType") = "N"
                If e.DataType.Contains("Decimal") Then 'Kiểu số->lấy Format
                    Dim sFormat As String = e.NumberFormat
                    If sFormat IsNot Nothing AndAlso sFormat <> "" Then
                        If sFormat = "Percent" Or sFormat.Contains("%") Then  'Thêm trường hợp format Percent - Update 06/11/2012
                            dr("DataType") = "Percent"
                        ElseIf sFormat.Contains(".") Then
                            Dim arr() As String = sFormat.Split(CType(".", Char))
                            dr("NumberFormat") = arr(arr.Length - 1).Length
                        ElseIf sFormat.Contains("N") Then 'Bổ sung 07/10/2011 TH định dạng trên cột "Nx"
                            dr("NumberFormat") = sFormat.Substring(1, sFormat.Length - 1)
                        End If
                    End If
                End If

                'Kiểm tra cột nào có tính tổng
                If arrColSum IsNot Nothing AndAlso arrColSum.Length > 0 Then
                    dr("IsSum") = e.IsSum
                End If

                If arrColSum IsNot Nothing AndAlso arrColSum.Length > 0 Then
                    If L3FindInteger(arrColSum, C1Grid.Columns.IndexOf(C1Grid.Columns(e.FieldName))) Then
                        dr("IsSum") = "1"
                    Else
                        dr("IsSum") = "0"
                    End If
                End If

            ElseIf e.DataType = "DateTime" Then 'Kiểu ngày
                dr("DataType") = "D"
                If arrColDateLong IsNot Nothing AndAlso arrColDateLong.Length > 0 Then
                    If L3FindInteger(arrColDateLong, C1Grid.Columns.IndexOf(C1Grid.Columns(e.FieldName))) Then
                        dr("IsDateTime") = "1"
                    Else
                        dr("IsDateTime") = "0"
                    End If
                End If
            Else 'Kiểu chuỗi
                dr("DataType") = "S"
            End If

            dt.Rows.Add(dr)
        Next

        Return dt

    End Function

    '''' <summary>
    '''' Tạo Table Caption cho UserControl và Grid Export Excel
    '''' </summary>
    '''' <param name="C1Grid">Lưới của form truyền vào</param>
    '''' <param name="ar">mảng các cột được hiển thị trên lưới</param>
    '''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Public Function CreateTableForExcel(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal arrColVisible As ArrayList) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("FieldName", GetType(System.String))
        dt.Columns.Add("Description", GetType(System.String))
        dt.Columns.Add("OrderNum", GetType(System.Int32))
        dt.Columns.Add("OrderNo", GetType(System.Int32))
        dt.Columns.Add("DataType", GetType(System.String))
        dt.Columns.Add("IsUsed", GetType(System.Boolean))
        dt.Columns.Add("IsUnicode", GetType(System.Boolean))
        dt.Columns.Add("NumberFormat", GetType(System.Byte))
        dt.Columns.Add("Obligatory", GetType(System.Byte))
        dt.Columns.Add("Grouped", GetType(System.Byte))
        dt.Columns.Add("IsSum", GetType(System.Byte))
        dt.Columns.Add("IsDateTime", GetType(System.Byte))
        dt.Columns.Add("IsExport", GetType(System.Byte))
        dt.Columns.Add("DataWidth", GetType(System.Int32))

        Dim dr As DataRow
        Dim iCount As Integer = 0

        For i As Integer = 0 To arrColVisible.Count - 1
            iCount += 1
            dr = dt.NewRow

            Dim e As ShowColumn = CType(arrColVisible(i), ShowColumn)
            dr(0) = e.FieldName
            dr(1) = e.Caption
            dr("Obligatory") = e.Obligatory
            dr("Grouped") = e.Grouped
            dr("IsSum") = e.IsSum
            dr("IsDateTime") = e.IsDateTime
            dr("IsExport") = 0
            dr("DataWidth") = e.DataWidth

            'Lay do dai max
            If dr(1).ToString.Length > giMaxLengthColumnCaption Then giMaxLengthColumnCaption = dr(1).ToString.Length
            dr(2) = iCount
            dr(3) = 0
            dr(4) = "S"
            dr(5) = 1
            dr(6) = 0
            dr(7) = 0
            dt.Rows.Add(dr)
        Next

        Return dt

    End Function

    '<DebuggerStepThrough()> _
    'Public Function L3FindInteger(ByVal arrCol() As Integer, ByVal iValue As Integer) As Boolean
    '    For i As Integer = 0 To arrCol.Length - 1
    '        If CInt(arrCol(i)) = iValue Then
    '            Return True
    '        End If
    '    Next
    '    Return False
    'End Function


    'Dim sValue As String = ""
    'Private Function ContainsValue(ByVal s As String) As Boolean
    '    'AndAlso prevents evaluation of the second Boolean
    '    'expression if the string is so short that an error
    '    'would occur.
    '    'Return s.Contains(sValue)
    '    Return s.Equals(sValue)

    'End Function

    'Public Function L3FindArrString(ByVal ArrString() As String, ByVal sValueFind As String) As Boolean
    '    sValue = sValueFind
    '    If Array.Exists(ArrString, AddressOf ContainsValue) Then
    '        Return True
    '    End If

    '    Return False
    'End Function

#End Region

#Region "Xuất Excel có định dạng"

    Public Sub CreateColStyleExcel(ByRef dtGrid As DataTable, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        If dtGrid.Columns.Contains(COL_StyleExcel) = False Then dtGrid.Columns.Add(COL_StyleExcel, GetType(System.String))
        If tdbg.Columns(tdbg.Columns.Count - 1).DataField <> COL_StyleExcel Then
            'Thêm cột StyleExcel trên lưới
            Dim dc As New C1.Win.C1TrueDBGrid.C1DataColumn
            dc.DataField = COL_StyleExcel
            tdbg.Columns.Add(dc)
            tdbg.Columns(dc.DataField).Caption = dc.DataField
            For i As Integer = 0 To tdbg.Splits.Count - 1
                tdbg.Splits(i).DisplayColumns(dc.DataField).Visible = False
            Next
        End If
        dtGrid.AcceptChanges()
    End Sub

    ''' <summary>
    ''' Add định dạng của dòng vào cột Style
    ''' Dữ liệu cột Style lưu định dạng của dòng(khi có sự kiện FechtRowStyle)và những cell có thuộc tính FetCellStyle=True
    ''' </summary>
    ''' <param name="tdbg">Lưới</param>
    ''' <param name="COL_Style">Cột cần gán giá trị</param>
    ''' <param name="bSetBackColor">Có gán màu nền không</param>
    ''' <param name="bSetForeColor">có gán màn chữ không</param>
    ''' <remarks>cách gọi tại sự kiện tdbg_FetchRowStyle:  AddRowStyleValue(e, tdbg, COL_Style, bSetBackColor, bSetForeColor)</remarks>
    Public Sub AddRowStyleValue(ByVal e As C1.Win.C1TrueDBGrid.FetchRowStyleEventArgs, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, Optional ByVal bSetBackColor As Boolean = False, Optional ByVal bSetForeColor As Boolean = False)
        '****************************
        Dim sFont As String = ""
        If e.CellStyle.Font.Bold Then sFont &= "B"
        If e.CellStyle.Font.Italic Then sFont &= "I"
        If e.CellStyle.Font.Underline Then sFont &= "U"
        '****************************
        Dim sForeColor As String = IIf(bSetForeColor = False, "", (e.CellStyle.ForeColor).ToArgb).ToString
        Dim sBlackColor As String = IIf(bSetBackColor = False, "", (e.CellStyle.BackColor).ToArgb).ToString
        '****************************
        Dim sStyle As String = "(" & sFont & "," & sForeColor & "," & sBlackColor & ")"
        'Nếu định dạng dòng này đã được thêm vào rồi thì không thêm nữa
        If tdbg(e.Row, COL_StyleExcel).ToString.Contains(sStyle) = False Then
            tdbg(e.Row, COL_StyleExcel) = tdbg(e.Row, COL_StyleExcel).ToString & sStyle
        End If
    End Sub

    ''' <summary>
    ''' Add định dạng của cell vào cột Style
    ''' Dữ liệu cột Style lưu định dạng của dòng(khi có sự kiện FechtRowStyle)và những cell có thuộc tính FetCellStyle=True
    ''' </summary>
    ''' <param name="tdbg">Lưới</param>
    ''' <param name="COL_Style">Cột cần gán giá trị</param>
    ''' <param name="bSetBackColor">Có gán màu nền không</param>
    ''' <param name="bSetForeColor">có gán màn chữ không</param>
    ''' <remarks>cách gọi tại sự kiện tdbg_FetchRowStyle:  AddRowStyleValue(e, tdbg, COL_Style, bSetBackColor, bSetForeColor)</remarks>
    Public Sub AddCellStyleValue(ByVal e As C1.Win.C1TrueDBGrid.FetchCellStyleEventArgs, ByRef tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, Optional ByVal bSetBackColor As Boolean = False, Optional ByVal bSetForeColor As Boolean = False)
        '****************************
        Dim sFont As String = ""
        If e.CellStyle.Font.Bold Then sFont &= "B"
        If e.CellStyle.Font.Italic Then sFont &= "I"
        If e.CellStyle.Font.Underline Then sFont &= "U"
        '****************************
        Dim sForeColor As String = IIf(bSetForeColor = False, "", (e.CellStyle.ForeColor).ToArgb).ToString
        Dim sBlackColor As String = IIf(bSetBackColor = False, "", (e.CellStyle.BackColor).ToArgb).ToString
        '****************************
        Dim sStyle As String = tdbg.Columns(e.Col).DataField & "(" & sFont & "," & sForeColor & "," & sBlackColor & ")"
        If tdbg(e.Row, COL_StyleExcel).ToString = "" Then
            tdbg(e.Row, COL_StyleExcel) = sStyle
        Else
            'Nếu định dạng cell này đã được thêm vào rồi thì không thêm nữa
            If tdbg(e.Row, COL_StyleExcel).ToString.Contains(tdbg.Columns(e.Col).DataField) = False Then
                tdbg(e.Row, COL_StyleExcel) = tdbg(e.Row, COL_StyleExcel).ToString & ";" & sStyle
            End If
        End If
    End Sub
#End Region

End Module
