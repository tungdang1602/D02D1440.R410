'#------------------------------------------------------
'#Title: D99X4444
'#CreateUser: NGUYEN NGOC THANH
'#CreateDate: 30/07/2007
'#ModifiedUser: NGUYEN NGOC THANH
'#ModifiedDate: 31/12/2007
'#Description:
'#------------------------------------------------------

Imports System.IO
Imports D00D0041.D00C0001
Imports D99D0041.D99C0008

Module D99X4444

    Public Const sFileName As String = "\Xreports\D99R9x9x.rpt"

    Public Const sFirstStartDate As String = "30/07/2007"
    Public Const sFirstIncrementDate As String = "30/07/2007"
    Public Const sFirstEndDate As String = "30/06/2008"

    Public Const iMax As Integer = 10

    Public Const sAppName As String = "Application Name"
    Public Const sSection As String = "Check Info"

    Public Const sStart As String = "1"
    Public Const sPar02 As String = "2"
    Public Const sPar13 As String = "3"
    Public Const sPar04 As String = "4"
    Public Const sIncrementTime As String = "5" 'OK
    Public Const sPar06 As String = "6"
    Public Const sEnd As String = "7"
    Public Const sPar08 As String = "8"
    Public Const sPar19 As String = "9"
    Public Const sPar10 As String= "10"
    Public Const sPar11 As String = "11"
    Public Const sPar12 As String = "12"
    Public Const sIncrement As String = "13" 'OK
    Public Const sPar14 As String= "14"
    Public Const sPar15 As String = "15"
    Public Const sPar16 As String = "16"
    Public Const sPar17 As String = "17"
    Public Const sPar18 As String = "18"
    Public Const sInvalidNum As String = "19" 'OK
    Public Const sPar20 As String = "20"

    Public sRegIncrementDate As String
    Public sRegIncrementDateTime As String
    Public iRegInvalidNum As Integer

    Public sFileIncrementDate As String
    Public sFileIncrementDateTime As String
    Public iFileInvalidNum As Integer

    Public sSystemDate As String
    Public sSystemDateTime As String

    Public Sub CheckLimitDate()

        sSystemDate = Date.Now.ToShortDateString
        sSystemDateTime = Date.Now.ToString

        'Lấy thông tin từ file và từ Reg
        If GetReg() = False And OpenFile() = False Then

            '+++Ghi thông tin ban đầu xuống File và Reg++++
            sRegIncrementDate = sFirstIncrementDate
            sRegIncrementDateTime = sSystemDateTime
            iRegInvalidNum = 0

            sFileIncrementDate = sFirstIncrementDate
            sFileIncrementDateTime = sSystemDateTime
            iFileInvalidNum = 0

            SaveFile()
            SaveReg()
            '+++++++++++++++++++++++++++++++++++++++++

            'Kiểm tra sự hợp lệ giữa ngày hệ thống và ngày quy địh
            If CDate(sSystemDate) < CDate(sFirstStartDate) Then
                MsgL3("Invalid system date. Please check information again.", L3MessageBoxIcon.Err)
                End
            End If

            'Cập nhật lại ngày hệ thống hiện tại
            sRegIncrementDate = sSystemDate
            sFileIncrementDate = sSystemDate
            sRegIncrementDateTime = sSystemDateTime
            sFileIncrementDateTime = sSystemDateTime

            SaveFile()
            SaveReg()

        Else

            '++++++++++++++Kiểm tra đồng bộ IncrementDate tại Reg và File+++++++++++++++++

            If sRegIncrementDate = "" Then sRegIncrementDate = sFileIncrementDate
            If sFileIncrementDate = "" Then sFileIncrementDate = sRegIncrementDate
            If sRegIncrementDate <> sFileIncrementDate Then sRegIncrementDate = sFileIncrementDate

            If sRegIncrementDateTime = "" Then sRegIncrementDateTime = sFileIncrementDateTime
            If sFileIncrementDateTime = "" Then sFileIncrementDateTime = sRegIncrementDateTime
            If sRegIncrementDateTime <> sFileIncrementDateTime Then sRegIncrementDateTime = sFileIncrementDateTime

            If iRegInvalidNum <> iFileInvalidNum Then iRegInvalidNum = iFileInvalidNum

            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            'Kiểm tra sự hợp lệ giữa ngày hệ thống hiện tại và ngày cập nhật
            If CDate(sSystemDate) < CDate(sRegIncrementDate) Then
                MsgL3("Invalid system date. Please check information again.(1)", L3MessageBoxIcon.Err)
                End
            ElseIf CDate(sSystemDate) = CDate(sRegIncrementDate) Then
                If CDate(sSystemDateTime) <= CDate(sRegIncrementDateTime) Then
                    iRegInvalidNum = iRegInvalidNum + 1
                    iFileInvalidNum = iFileInvalidNum + 1
                End If
            End If

            'Gán ngày hệ thống
            If iRegInvalidNum >= iMax Then
                sRegIncrementDate = DateAdd("D", 1, sRegIncrementDate).ToString
                sFileIncrementDate = DateAdd("D", 1, sRegIncrementDate).ToString
                sRegIncrementDateTime = DateAdd("D", 1, sRegIncrementDateTime).ToString & " 00:00:00 AM"
                sFileIncrementDateTime = DateAdd("D", 1, sRegIncrementDateTime).ToString & " 00:00:00 AM"
                iRegInvalidNum = 0
                iFileInvalidNum = 0
            Else
                sRegIncrementDate = sSystemDate
                sFileIncrementDate = sSystemDate
                sRegIncrementDateTime = sSystemDateTime
                sFileIncrementDateTime = sSystemDateTime
            End If

            SaveFile()
            SaveReg()

            'Kiểm tra sự hợp lệ giữa ngày cập nhật và ngày kết thúc
            If CDate(sRegIncrementDate) >= CDate(sFirstEndDate) Then
                MsgL3("Your trial has expired. Please contact your vendor directly for support.", L3MessageBoxIcon.Err)
                End
            End If

        End If

    End Sub

    Private Function OpenFile() As Boolean
        Dim Arr() As String
        Dim sStr As String

        Try

            sStr = My.Computer.FileSystem.ReadAllText(Application.StartupPath & sFileName, System.Text.Encoding.Default).ToString 'Open file with Encoding ANSI 
            Arr = Split(sStr, Chr(9))
            sFileIncrementDate = EncryptString(Arr(1), False)
            sFileIncrementDateTime = EncryptString(Arr(3), False)
            iFileInvalidNum = CInt(IIf(EncryptString(Arr(4), False) = "", 0, EncryptString(Arr(4), False)))

            If sFileIncrementDate = "" Then Return False
            Return True

        Catch exc As Exception
            Return False
        End Try

    End Function

    Private Sub SaveFile()
        Dim PathFile As String = Application.StartupPath & sFileName
        Dim sCommand As String = EncryptString(DateAdd("D", 1, sSystemDateTime).ToString, True) & Chr(9) & _
                                    EncryptString(sFileIncrementDate, True) & Chr(9) & _
                                    EncryptString(DateAdd("D", 9, sSystemDate).ToString, True) & Chr(9) & _
                                    EncryptString(sFileIncrementDateTime, True) & Chr(9) & _
                                    EncryptString(CStr(iFileInvalidNum), True) & Chr(9) & _
                                    EncryptString(DateAdd("D", 19, sSystemDate).ToString, True) & Chr(9) & _
                                    EncryptString(DateAdd("D", 3, sSystemDateTime).ToString, True) & Chr(9) & _
                                    EncryptString(DateAdd("D", 22, sSystemDateTime).ToString, True) & Chr(9) & _
                                    EncryptString(DateAdd("D", 7, sSystemDate).ToString, True) '1,3,4
        Try

            My.Computer.FileSystem.WriteAllText(Application.StartupPath & sFileName, sCommand, False, System.Text.Encoding.Default)

        Catch exc As Exception
        End Try

    End Sub

    Private Sub SaveReg()

        SaveSetting(sAppName, sSection, sIncrement, EncryptString(sRegIncrementDate, True))
        SaveSetting(sAppName, sSection, sIncrementTime, EncryptString(sRegIncrementDateTime, True))
        SaveSetting(sAppName, sSection, sInvalidNum, EncryptString(CStr(iRegInvalidNum), True))

        SaveSetting(sAppName, sSection, sStart, EncryptString(DateAdd("D", 1, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sEnd, EncryptString(DateAdd("D", 9, sSystemDateTime).ToString, True))
        SaveSetting(sAppName, sSection, sPar02, EncryptString(DateAdd("D", 2, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar04, EncryptString(DateAdd("D", 4, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar06, EncryptString(DateAdd("D", 6, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar08, EncryptString(DateAdd("D", 8, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar10, EncryptString(DateAdd("D", 10, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar11, EncryptString(DateAdd("D", 11, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar12, EncryptString(DateAdd("D", 12, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar13, EncryptString(DateAdd("D", 13, sSystemDateTime).ToString, True))
        SaveSetting(sAppName, sSection, sPar14, EncryptString(DateAdd("D", 14, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar15, EncryptString(DateAdd("D", 15, sSystemDateTime).ToString, True))
        SaveSetting(sAppName, sSection, sPar16, EncryptString(DateAdd("D", 16, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar17, EncryptString(DateAdd("D", 17, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar18, EncryptString(DateAdd("D", 18, sSystemDate).ToString, True))
        SaveSetting(sAppName, sSection, sPar19, EncryptString(DateAdd("D", 19, sSystemDateTime).ToString, True))
        SaveSetting(sAppName, sSection, sPar20, EncryptString(DateAdd("D", 20, sSystemDate).ToString, True))

    End Sub

    Private Function GetReg() As Boolean

        GetReg = True

        sRegIncrementDate = EncryptString(GetSetting(sAppName, sSection, sIncrement), False)
        sRegIncrementDateTime = EncryptString(GetSetting(sAppName, sSection, sIncrementTime), False)
        iRegInvalidNum = CInt(IIf(EncryptString(GetSetting(sAppName, sSection, sInvalidNum), False) = "", 0, EncryptString(GetSetting(sAppName, sSection, sInvalidNum), False)))

        If sRegIncrementDate = "" Then GetReg = False

    End Function

End Module
