Public Class Branch

    Private field_2 = False
    Private field_3 = False
    Private field_4 = False
    Private visibleField = "2字段"

    Private cdr_url = False
    Private cdr_bjnews = False

    Private cdr_mobile = False
    Private cdr_phone = False

    Private cdr_email = False
    Private cdr_qq = False

    '可用字段
    Public Function setField(key)
        Select Case key
            Case "2字段"
                field_2 = True
            Case "3字段"
                field_3 = True
            Case "4字段"
                field_4 = True
        End Select
    End Function


    '设置使用层级模板
    Public Function setVisibleField(cmdExternalData)

        '如果有4字段 显示层级4
        If field_4 = True Then
            If cmdExternalData("bjnews") <> "" Or cmdExternalData("url") <> "" Then
                visibleField = "4字段"
            End If
        End If

        '如果有3字段
        If field_3 = True Then
            '4字段的优先级更高
            If visibleField <> "4字段" Then
                If cmdExternalData("mail") <> "" Or cmdExternalData("qq") <> "" Then
                    visibleField = "3字段"
                End If
            End If

        End If
    End Function


    '获取字段
    Public Function getVisibleField()
        Return visibleField
    End Function



    Public Function setState(key)
        Select Case key
            Case "url"
                cdr_url = True
            Case "bjnews"
                cdr_bjnews = True
            Case "mobile"
                cdr_mobile = True
            Case "phone"
                cdr_phone = True
            Case "email"
                cdr_email = True
            Case "qq"
                cdr_qq = True
        End Select
    End Function

    Public Function getScope(key)
        If key = "url" Or key = "bjnews" Or key = "mobile" Or key = "phone" Or key = "email" Or key = "qq" Then
            Return True
        Else
            Return False
        End If
    End Function

    '设置数据，可能会存在合并的情况
    Public Function getValue(key, cmdExternalData)
        Dim newValue = cmdExternalData(key)
        Select Case key
                '网址/公众号
            Case "url"
                Dim user_bjnews = cmdExternalData("bjnews")
                If user_bjnews <> "" And cdr_bjnews = False Then
                    newValue = newValue + Chr(13) + user_bjnews
                End If
            Case "bjnews"
                Dim user_url = cmdExternalData("url")
                If user_url <> "" And cdr_url = False Then
                    newValue = newValue + Chr(13) + user_url
                End If
               '手机/固定电话
            Case "mobile"
                Dim user_phone = cmdExternalData("phone")
                '没有电话字段，但是用户设置了手机
                If user_phone <> "" And cdr_phone = False Then
                    newValue = newValue + Chr(13) + user_phone
                End If
            Case "phone"
                '没有手机字段，但是用户设置了电话
                Dim user_mobile = cmdExternalData("mobile")
                If user_mobile <> "" And cdr_mobile = False Then
                    newValue = newValue + Chr(13) + user_mobile
                End If
                '邮箱/QQ
            Case "email"
                Dim user_qq = cmdExternalData("qq")
                If user_qq <> "" And cdr_qq = False Then
                    newValue = newValue + Chr(13) + user_qq
                End If
            Case "qq"
                Dim user_email = cmdExternalData("email")
                If user_email <> "" And cdr_email = False Then
                    newValue = newValue + Chr(13) + user_email
                End If
        End Select
        Return newValue
    End Function

End Class
