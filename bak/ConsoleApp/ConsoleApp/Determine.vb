﻿Imports Corel.Interop.VGCore

'数据判断类，是否分行
Class Determine
    '模板存在的字段
    Private field_2 = False
    Private field_3 = False
    Private field_4 = False
    Private visibleField = "2字段"

    '组合字段状态定义
    Private cdr_url = False
    Private cdr_bjnews = False

    Private cdr_mobile = False
    Private cdr_phone = False

    Private cdr_email = False
    Private cdr_qq = False


    '可用字段
    Private Function setField(key)
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
    Private Function setVisibleField()
        '如果有4字段 显示层级4
        If field_4 = True Then
            If Param.hasValue("bjnews") Or Param.hasValue("url") Then
                visibleField = "4字段"
            End If
        End If

        '如果有3字段
        If field_3 = True Then
            '4字段的优先级更高
            If visibleField <> "4字段" Then
                If Param.hasValue("email") Or Param.hasValue("qq") Then
                    visibleField = "3字段"
                End If
            End If
        End If

    End Function

    '初始化字段的状态
    '涉及到状态合并的问题处理
    Private Function setState(key)
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


    '设置所有字段
    Private Function setAllField(tempShape)
        Dim key = Utils.getKeyEnglish(tempShape.Name)
        If key <> "" Then
            globalData.saveInputFiled(key)
        End If
    End Function


    '文本预处理
    Private Function proccessText(allShapes, pageIndex)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)
            Dim cdrTextShape As cdrShapeType = 6
            Dim cdrGroupShape As cdrShapeType = 7

            '保存所有字段
            setAllField(tempShape)

            '组
            If tempShape.Type = cdrGroupShape Then
                proccessText(tempShape.Shapes, pageIndex)
            End If

            '文字
            If tempShape.Type = cdrTextShape Then
                Dim key As String = Utils.getKeyEnglish(tempShape.Name)
                setState(key)
            End If
        Next k
    End Function



    '=================================== 对外接口 ===================================



    '初始化
    Function init(key, shapes, pageIndex)
        setField(key)
        setVisibleField()
        proccessText(shapes, pageIndex)
    End Function



    '判断是否需要合并的数据
    Public Function getRangeScope(key)
        If key = "url" Or key = "bjnews" Or key = "mobile" Or key = "phone" Or key = "email" Or key = "qq" Then
            Return True
        Else
            Return False
        End If
    End Function


    '设置数据，可能会存在合并的情况
    Public Function getMergeValue(key)
        Dim newValue = Param.getExternalValue(key)
        Select Case key
                '网址/公众号
            Case "url"
                'url + bjnews
                If Param.hasValue("bjnews") And cdr_bjnews = False Then
                    globalData.addRecord("bjnews被合并到url中")
                    Dim user_bjnews = Param.getExternalValue("bjnews")
                    newValue = newValue + Chr(13) + user_bjnews
                End If
            Case "bjnews"
                'url + bjnews
                If Param.hasValue("url") And cdr_url = False Then
                    globalData.addRecord("url被合并到bjnews中")
                    Dim user_url = Param.getExternalValue("url")
                    newValue = user_url + Chr(13) + newValue
                End If
               '手机/固定电话
            Case "mobile"
                '没有电话字段，但是用户设置了手机
                'mobile + phone
                If Param.hasValue("phone") And cdr_phone = False Then
                    globalData.addRecord("phone被合并到mobile中")
                    Dim user_phone = Param.getExternalValue("phone")
                    newValue = newValue + Chr(13) + user_phone
                End If
            Case "phone"
                '没有手机字段，但是用户设置了电话
                'mobile + phone
                If Param.hasValue("mobile") And cdr_mobile = False Then
                    globalData.addRecord("mobile被合并到phone中")
                    Dim user_mobile = Param.getExternalValue("mobile")
                    newValue = user_mobile + Chr(13) + newValue
                End If
                '邮箱/QQ
            Case "email"
                'email + qq
                If Param.hasValue("qq") And cdr_qq = False Then
                    globalData.addRecord("qq被合并到email中")
                    Dim user_qq = Param.getExternalValue("qq")
                    newValue = newValue + Chr(13) + user_qq
                End If
            Case "qq"
                'email + qq
                If Param.hasValue("email") And cdr_email = False Then
                    globalData.addRecord("email被合并到qq中")
                    Dim user_email = Param.getExternalValue("email")
                    newValue = user_email + Chr(13) + newValue
                End If
        End Select

        Return newValue
    End Function



    '////////////////////////////////////// 可见性 //////////////////////////////////////////////////



    '获取字段的状态
    Public Function getVisibleField()
        Return visibleField
    End Function


    Private Function setVisible(activeLayer, name, visibleLayerName)
        If name = visibleLayerName Then
            activeLayer.Visible = True
        Else
            activeLayer.Visible = False
        End If
    End Function


    '设置层级的可见性    
    '如果网址/公众号，都没有，那么要隐藏“4 字段”图层，显示“3 字段”图层。如果邮箱/QQ 号，也没有，那么就显示“2 字段图层
    Public Function setLayerVisible(activeLayer As Layer, visibleLayerName As String)
        Dim name = activeLayer.Name
        If name = "2字段" Then
            setVisible(activeLayer, name, visibleLayerName)
        End If
        If name = "3字段" Then
            setVisible(activeLayer, name, visibleLayerName)
        End If
        If name = "4字段" Then
            setVisible(activeLayer, name, visibleLayerName)
        End If
    End Function


End Class