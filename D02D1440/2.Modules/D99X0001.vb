'#######################################################################################
'#                                     CHÚ Ý
'#--------------------------------------------------------------------------------------
'# Không được thay đổi bất cứ dòng code này trong module này, nếu muốn thay đổi bạn phải
'# liên lạc với Trưởng nhóm để được giải quyết.
'# Ngày cập nhật cuối cùng: 18/06/2013
'# Người cập nhật cuối cùng: Nguyễn Thị Minh Hòa
'# Diễn giả: bổ sung resource thay thế cho mỗi khách hàng: vd  replace tên "Phòng" thành "Xưởng"
'#######################################################################################
Imports System.Resources
Imports System.IO

''' <summary>
''' Module quản lý các vấn đề về Resource
''' </summary>
Public Module D99X0001
    ''' <summary>
    ''' Lưu trữ Resource tiếng Việt
    ''' </summary>
    Private rV As ResourceManager = ResourceManager.CreateFileBasedResourceManager("Vietnamese", Application.StartupPath, Nothing)
    ''' <summary>
    ''' Lưu trữ Resource tiếng Anh
    ''' </summary>
    Private rE As ResourceManager = ResourceManager.CreateFileBasedResourceManager("English", Application.StartupPath, Nothing)

    ''' <summary>
    ''' Trả về chuỗi resource ứng với ResourceID truyền vào
    ''' </summary>
    ''' <param name="ResourceID">Mã Resource</param>
    '''<remarks>Nếu không tìm thấy ResourceID ở file Tiếng Anh thì sẽ
    ''' trả về Resouce ở file Tiếng Việt, và nếu không tìm thấy Resource ở 
    ''' file Tiếng Việt sẽ trả về Mã ResourceID này
    ''' </remarks>
    Public Function r(ByVal ResourceID As String) As String
        Try

            Dim sRes As String = ""
            If geLanguage = EnumLanguage.Vietnamese Then
                sRes = rV.GetString(ResourceID).ToString
            Else
                sRes = rE.GetString(ResourceID).ToString
            End If

            'Update 17/06/2013: Nếu có thay đổi tên resource
            If giReplacResource <> 0 AndAlso gsConnectionString <> "" Then
                sRes = ReplaceResourceCustom(sRes)
            End If

            Return sRes

        Catch
            Try
                'Update 17/06/2013: Nếu đã thông báo lỗi 1 lần thì không cần thông báo nữa
                If Not gbResourceError Then
                    gbResourceError = True
                    If geLanguage = EnumLanguage.Vietnamese Then
                        If MessageBox.Show("˜Ò nghÜ li£n hÖ nhª cung cÊp ¢Ó cËp nhËt mìi 2 file ng¤n ngö: Vietnamese.resources vª English.resources" & vbCrLf & _
                            "BÁn câ muçn tiÕp tóc kh¤ng?" & vbCrLf & _
                            "Læi: [" + ResourceID + "]", MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                            Return "[" + rV.GetString(ResourceID).ToString + "]"
                        Else
                            End
                        End If
                    Else
                        If MessageBox.Show("Please contact to provider for update two resource file: Vietnamese.resource and English.resource" & vbCrLf & _
                            "Do you want to continue?" & vbCrLf & _
                            "Error: [" + ResourceID + "]", MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                            Return "[" + rV.GetString(ResourceID).ToString + "]"
                        Else
                            End
                        End If
                    End If
                Else
                    Return "[" + rV.GetString(ResourceID).ToString + "]"
                End If
            Catch ex As Exception
                If Len(ResourceID) > 74 Then
                    Return "[" + ResourceID.Substring(0, 70) + "...]"
                Else
                    Return "[" + ResourceID + "]"
                End If
            End Try
        End Try

    End Function

End Module
