'#------------------------------------------------------
'#Title: D99X0003
'#CreateUser: NGUYEN NGOC THANH
'#CreateDate: 24/03/2004
'#ModifiedUser: Hồ Ngọc Thoại
'#ModifiedDate: 10/04/2013
'#Description: Phân quyền In
'# Chứa các biến dùng chung (Print)
'# Sửa lại biến gcCon1 thành gcConPrint để khỏi trùng biến trong dll D99D0041
'#------------------------------------------------------

Imports System.Data.SqlClient

Friend Module D99X0003

    'Kết nối khi in báo cáo
    Public gcConPrint As New SqlConnection
    'Public gsServerName1 As String
    'Public gsPassword1 As String
    'Public gsDataBase1 As String
    'Public gsUser1 As String

    'Số dòng cho phép export của Excel
    Public Const RowMaximumExcel As Integer = 65528

    Public nRowMaximum As Double

    Public nCountPrint As Integer
    Public bCountPrint As Boolean

    Public rpt1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument

    'Kiểm tra in
    Public gbFlagPrint1 As Boolean

    'Chứa đường dẫn report chính
    Public gsMainReportName1 As String

    'Caption form báo cáo chính
    Public gsMainReportCaption1 As String

    'Nguồn dữ liệu thực thi chính
    Public gsMainStrData1 As String = ""
    'Nguồn dữ liệu thực thi chính
    Public dtReportMain As DataTable
    'Chứa tên máy in
    Public gsPrinterName1 As String
    'Đường dẫn chứa báo cáo tạm
    Public gsSaveFileName1 As String

    'Lưu lại tổng số lần in ra giấy của Hóa đơn
    Public giPrintNumber As Integer = 0
    'Phân quyền nút Print, PrintSetup
    Public giPermissionPrint As Integer = 4
End Module
