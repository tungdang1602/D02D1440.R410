''' <summary>
''' Module này liên qua đến các khai báo biến, enum, ... toàn cục
''' </summary>
''' <remarks>Các khai báo ở đây không được trùng với các khai báo ở các Module D99Xxxxx</remarks>
Module D02X0001

    ''' <summary>
    ''' Module đang coding D02E1440
    ''' </summary>
    Public Const MODULED02 As String = "D02E1440"
    ''' <summary>
    ''' Chuỗi D02
    ''' </summary>
    Public Const D02 As String = "D02"
    ''' <summary>
    ''' Dùng cho kiểm tra Security theo chuẩn của DIGINET
    ''' </summary>
    Public Const L3_APP_NAME As String = "Lemon3"
    ''' <summary>
    ''' Dùng cho kiểm tra Security theo chuẩn của DIGINET
    ''' </summary>
    Public Const L3_HS_SECTION As String = "Handshake"
    ''' <summary>
    ''' Dùng cho kiểm tra Security theo chuẩn của DIGINET
    ''' </summary>
    Public Const L3_HS_MODULE As String = "D02"
    ''' <summary>
    ''' Dùng cho kiểm tra Security theo chuẩn của DIGINET
    ''' </summary>
    Public Const L3_HS_VALUE As String = "R3.60.00.Y2007"
    ''' <summary>
    ''' Dùng cho kiểm tra luu thanh cong 
    ''' </summary>
    Public gbSaveOK As Boolean

    Public Structure StructureFormat
        Public Percentage As String
        Public NumberD02 As String
        ''' <summary>
        ''' format thành tiền
        ''' </summary>
        Public OriginalAmount As String
        ''' <summary>
        ''' Số làm tròn của thành tiền
        ''' </summary>
        Public OriginalAmountRound As Integer
        ''' <summary>
        ''' format thành tiền quy đổi
        ''' </summary>
        Public ConvertedAmount As String
        ''' <summary>
        ''' Số làm tròn của thành tiền quy đổi
        ''' </summary>
        Public ConvertedAmountRound As Integer
        ''' <summary>
        ''' format tỷ giá
        ''' </summary>
        Public ExchangeRate As String
        ''' <summary>
        ''' Số làm tròn của tỷ giá
        ''' </summary>
        Public ExchangeRateRound As Integer
        ''' <summary>
        ''' Nguyên tệ gốc
        ''' </summary>
        Public BaseCurrencyID As String
        ''' <summary>
        ''' Dấu phân cách thập phân
        ''' </summary>
        Public DecimalSeperator As String
        ''' <summary>
        ''' Dấu phân cách hàng ngàn
        ''' </summary>
        Public ThousandSeperator As String
        Public DefaultNumber2 As String
        '------------------------------------------------------------------------
        '  D91 Format here
        '------------------------------------------------------------------------

    End Structure
    ''' <summary>
    ''' Khai báo structure cho phần định dạng format theo chuẩn chung mới lấy từ D91P9300
    ''' </summary>
    Public Structure StructureFormatNew
        ''' <summary>
        ''' format tỷ giá
        ''' </summary>
        Public ExchangeRate As String
        ''' <summary>
        ''' format nguyên tệ
        ''' </summary>
        Public DecimalPlaces As String

        ''' <summary>
        ''' format nguyên tệ ứng với mỗi loại tiền
        ''' </summary>
        Public MyOriginal As String

        ''' <summary>
        ''' format tiền quy đổi
        ''' </summary>
        Public D90_Converted As String
        ''' <summary>
        ''' format số lượng, số lượng QĐ: nhóm sản xuất (D06, D07, D12, D37); nhóm kinh doanh
        ''' </summary>
        Public D07_Quantity As String
        ''' <summary>
        ''' format đơn giá: nhóm sản xuất (D06, D07, D12, D37); nhóm kinh doanh
        ''' </summary>
        ''' <remarks></remarks>
        Public D07_UnitCost As String
        ''' <summary>
        ''' format số lượng, số lượng QĐ: nhóm sản xuất (D08, D20, D30, D31, D32, D33, D34, D35, D36)
        ''' </summary>
        ''' <remarks></remarks>
        Public D08_Quantity As String
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public D08_UnitCost As String
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public D08_Ratio As String
        ''' <summary>
        ''' format số lượng, số lượng QĐ: nhóm sản xuất (danh mục Bộ định mức D08, danh mục Cấu trúc sản phẩm D32)
        ''' </summary>
        ''' <remarks></remarks>
        Public BOMQty As String
        ''' <summary>
        ''' format đơn giá: nhóm sản xuất (danh mục Bộ định mức D08, danh mục Cấu trúc sản phẩm D32)
        ''' </summary>
        ''' <remarks></remarks>
        Public BOMPrice As String
        ''' <summary>
        ''' format thành tiền: nhóm sản xuất (danh mục Bộ định mức D08, danh mục Cấu trúc sản phẩm D32)
        ''' </summary>
        ''' <remarks></remarks>
        Public BOMAmt As String
        ''' <summary>
        ''' Format 2 số lẻ (không theo quy tắc nào)
        ''' </summary>
        ''' <remarks></remarks>
        Public DefaultNumber2 As String
        Public DefaultNumber6 As String
    End Structure

    ''' <summary>
    ''' Khai bao de chon cac button tren luoi vd: Khoan muc, doi tuong,mat hang...
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum Button
        ObjectButton = 0
        Inventory = 1
        Ana = 2
    End Enum


    'Public D02Format As StructureFormat
    Public D02Format As StructureFormatNew

    ''' <summary>
    ''' Lưu trữ các thiết lập tùy chọn
    ''' </summary>
    Public D02Options As StructureOption

    ''' <summary>
    ''' Khai báo Structure cho phần Tùy chọn của Module
    ''' </summary>
    Public Structure StructureOption
        ''' <summary>
        ''' Hỏi trước khi lưu
        ''' </summary>
        Public MessageAskBeforeSave As Boolean
        ''' <summary>
        ''' Thông báo khi lưu thành công
        ''' </summary>
        Public MessageWhenSaveOK As Boolean
        ''' <summary>
        ''' Hiển thị form chọn kỳ kế toán khi chạy chương trình
        ''' </summary>
        Public ViewFormPeriodWhenAppRun As Boolean
        ''' <summary>
        ''' Lưu giá trị gần nhất
        ''' </summary>
        Public SaveLastRecent As Boolean
        ''' <summary>
        ''' Lưu đơn vị mặc định
        ''' </summary>
        Public DefaultDivisionID As String
        ''' <summary>
        ''' Khóa thành tiền quy đổi
        ''' </summary>
        Public LockConvertedAmount As Boolean
        ''' <summary>
        ''' Làm tròn thành tiền quy đổi
        ''' </summary>
        Public RoundConvertedAmount As Boolean
        ''' <summary>
        ''' Hiển thị quy trình sơ đồ nghiệp vụ
        ''' </summary>
        Public ViewWorkflow As Boolean
        ''' <summary>
        ''' Ngôn ngữ báo cáo
        ''' </summary>
        Public ReportLanguage As Byte
        '------------------------------------------------------------------------
        '  D02 Options here
        '------------------------------------------------------------------------
    End Structure
End Module
