''' <summary>
''' Các vấn đề liên quan đến Thông tin hệ thống và Tùy chọn
''' </summary>
Module D02X0004
    ''' <summary>
    ''' Load toàn bộ các thông số tùy chọn vào biến D02Options
    ''' </summary>
    Public Sub LoadOptions()
        With D02Options
            'Kiểm tra tồn tại đường dẫn mới lưu .Net thì lấy dữ liệu, ngược lại thì lấy theo đường dẫn cũ (Lemon3_Dxx)
            'Kiem tra ky cac ten luu xuong cua VB6 de gan vao NET

            If D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "MessageAskBeforeSave") = "" Then 'Lay duong dan cu VB6
                Dim D02LocalOptionsLocations As String = "D02"
                Dim Options As String = "Options"

                With D02Options
                    .DefaultDivisionID = GetSetting(D02LocalOptionsLocations, Options, "Division", "")
                    .MessageAskBeforeSave = CType(GetSetting(D02LocalOptionsLocations, Options, "AskBeforeSave", "True"), Boolean)
                    .MessageWhenSaveOK = CType(GetSetting(D02LocalOptionsLocations, Options, "MessageWhenSaveOK", "True"), Boolean)
                    .SaveLastRecent = CType(GetSetting(D02LocalOptionsLocations, Options, "SaveRecentValues", "False"), Boolean)
                    .RoundConvertedAmount = CType(GetSetting(D02LocalOptionsLocations, Options, "RoundConvertedAmount", "False"), Boolean)
                    .LockConvertedAmount = CType(GetSetting(D02LocalOptionsLocations, Options, "LockConvertedAmount", "False"), Boolean)
                    .ViewFormPeriodWhenAppRun = CType(GetSetting(D02LocalOptionsLocations, Options, "AcountingScreen", "False"), Boolean)
                    .ReportLanguage = CType(GetSetting(D02LocalOptionsLocations, Options, "nRPLang", "0"), Byte)
                    .ViewWorkflow = CType(GetSetting(D02LocalOptionsLocations, Options, "ShowDiagramTransaction", "False"), Boolean)
                End With
            Else 'Lấy đường dẫn mới .Net
                With D02Options
                    .DefaultDivisionID = D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "DefaultDivisionID", "")
                    .MessageAskBeforeSave = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "MessageAskBeforeSave", "True"), Boolean)
                    .MessageWhenSaveOK = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "MessageWhenSaveOK", "True"), Boolean)
                    .SaveLastRecent = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "SaveLastRecent", "False"), Boolean)
                    .RoundConvertedAmount = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "RoundConvertedAmount", "False"), Boolean)
                    .LockConvertedAmount = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "LockConvertedAmount", "False"), Boolean)
                    .ViewFormPeriodWhenAppRun = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "ViewFormPeriodWhenAppRun", "False"), Boolean)
                    .ReportLanguage = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "ReportLanguage", "0"), Byte)
                    .ViewWorkflow = CType(D99C0007.GetModulesSetting(D02, ModuleOption.lmOptions, "ViewWorkflow", "False"), Boolean)
                End With
            End If
        End With
    End Sub

    ''' <summary>
    ''' Load toàn bộ các thông số format cho biến DxxFormat theo chuẩn chung mới lấy từ D91P9300
    ''' </summary>
    Public Sub LoadFormatsNew()
        Const Number2 As String = "#,##0.00"
        Dim sSQL As String = "Exec D91P9300 "
        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        With D02Format
            If dt.Rows.Count > 0 Then
                .ExchangeRate = InsertFormat(dt.Rows(0).Item("ExchangeRateDecimals"))
                .DecimalPlaces = InsertFormat(dt.Rows(0).Item("DecimalPlaces"))
                .MyOriginal = .DecimalPlaces
                .D90_Converted = InsertFormat(dt.Rows(0).Item("D90_ConvertedDecimals"))
                .D07_Quantity = InsertFormat(dt.Rows(0).Item("D07_QuantityDecimals"))
                .D07_UnitCost = InsertFormat(dt.Rows(0).Item("D07_UnitCostDecimals"))
                .D08_Quantity = InsertFormat(dt.Rows(0).Item("D08_QuantityDecimals"))
                .D08_UnitCost = InsertFormat(dt.Rows(0).Item("D08_UnitCostDecimals"))
                .D08_Ratio = InsertFormat(dt.Rows(0).Item("D08_RatioDecimals"))
                .BOMQty = InsertFormat(dt.Rows(0).Item("BOMQtyDecimals"))
                .BOMPrice = InsertFormat(dt.Rows(0).Item("BOMPriceDecimals"))
                .BOMAmt = InsertFormat(dt.Rows(0).Item("BOMAmtDecimals"))
            Else
                .ExchangeRate = Number2
                .D90_Converted = Number2
                .D07_Quantity = Number2
                .D07_UnitCost = Number2
                .D08_Quantity = Number2
                .D08_UnitCost = Number2
                .D08_Ratio = Number2
                .BOMQty = Number2
                .BOMPrice = Number2
                .BOMAmt = Number2
            End If

            .DefaultNumber2 = Number2
            .DefaultNumber6 = "#,##0.000000"
        End With
    End Sub

    Private Function InsertFormat(ByVal ONumber As Object) As String
        Dim iNumber As Int16 = Convert.ToInt16(ONumber)
        Dim sRet As String = "#,##0"
        If iNumber = 0 Then
        Else
            sRet &= "." & Strings.StrDup(iNumber, "0")
        End If
        Return sRet
    End Function

    Public Function GetOriginalDecimal(ByVal sCurrencyID As String) As String

        Dim sSQL As String
        sSQL = "Select OriginalDecimal From D91V0010 Where CurrencyID = " & SQLString(sCurrencyID)
        Dim dt As DataTable
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            Return InsertFormat(dt.Rows(0).Item("OriginalDecimal"))
        Else
            Return D02Format.DecimalPlaces
        End If
    End Function

    ''' <summary>
    ''' Hỏi trước khi lưu tùy thuộc vào thiết lập ở phần Tùy chọn
    ''' </summary>
    Public Function AskSave() As DialogResult
        If D02Options.MessageAskBeforeSave Then
            Return D99C0008.MsgAskSave()
        Else
            Return DialogResult.Yes
        End If
    End Function

    ''' <summary>
    ''' Thông báo trước khi xóa
    ''' </summary>    
    Public Function AskDelete() As DialogResult
        If D02Options.MessageAskBeforeSave Then
            Return D99C0008.MsgAskDelete
        Else
            Return DialogResult.Yes
        End If
    End Function

    ''' <summary>
    ''' Thông báo khi lưu thành công tùy theo phần thiết lập ở tùy chọn
    ''' </summary>
    Public Sub SaveOK()
        If D02Options.MessageWhenSaveOK Then D99C0008.MsgSaveOK()
    End Sub

    ''' <summary>
    ''' Thông báo sau khi xóa thành công
    ''' </summary>
    Public Sub DeleteOK()
        If D02Options.MessageWhenSaveOK Then D99C0008.MsgL3(rl3("MSG000008"))

    End Sub

    ''' <summary>
    ''' Thông báo không lưu được dữ liệu
    ''' </summary>
    Public Sub SaveNotOK()
        D99C0008.MsgSaveNotOK()
    End Sub

    ''' <summary>
    ''' Thông báo không xóa được dữ liệu
    ''' </summary>
    Public Sub DeleteNotOK()
        'D99C0008.MsgL3("Không xóa được dữ liệu")
        D99C0008.MsgCanNotDelete()
    End Sub
End Module
