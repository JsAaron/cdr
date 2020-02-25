﻿Imports Corel.Interop.VGCore

'input输入处理

Module Inputs

    '////////////////////////////////////// 文本 //////////////////////////////////////////////////

    '图片的读，创建基本结构
    Private Function getImage(tempShape As Shape, pageIndex As String, determine As Determine)
        Dim key = Utils.getKeyEnglish(tempShape.Name)
        If Len(key) > 0 Then
            globalData.saveValue(pageIndex, key, tempShape, determine, True)
        End If
        '
    End Function


    '获取文本
    Private Function getText(tempShape As Shape, pageIndex As String, determine As Determine)
        If tempShape.Text.Story.Text <> "" Then
            Dim key = Utils.getKeyEnglish(tempShape.Name)
            If Len(key) > 0 Then
                globalData.saveValue(pageIndex, key, tempShape, determine, False)
            Else
                ' Console.WriteLine("找不到对应的命名：" & tempShape.Name)
            End If
        End If
    End Function


    '设置文本
    Private Function setText(tempShape As Shape, pageIndex As String, determine As Determine)
        Dim key As String = Utils.getKeyEnglish(tempShape.Name)
        Dim value

        If key <> "" Then
            '是否是处理范围
            Dim hasRange = determine.getRangeScope(key)

            If hasRange = True Then
                '字段合并处理
                value = determine.getMergeValue(key)
            Else
                '单独字段
                value = Param.getExternalValue(key)
            End If

            '替换/清空
            If value <> "" Then
                If tempShape.Text.Story.Text = value Then
                    '不替换重复
                Else
                    tempShape.Text.Story.Delete()
                    tempShape.Text.Story.Replace(value)
                End If
            Else
                tempShape.Text.Story.Delete()
            End If

        End If
    End Function


    '递归检测形状
    Public Function accesstShape(doc, allShapes, determine, pageIndex)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)
            Dim cdrTextShape As cdrShapeType = 6
            Dim cdrGroupShape As cdrShapeType = 7
            Dim cdrBitmapShape As cdrShapeType = 5

            '组
            If tempShape.Type = cdrGroupShape Then
                accesstShape(doc, tempShape.Shapes, determine, pageIndex)
            End If

            '图片的读
            If tempShape.Type = cdrBitmapShape Then
                If Param.cmdCommand = "get:text" Then
                    getImage(tempShape, pageIndex, determine)
                End If
            End If

            '文字读写
            If tempShape.Type = cdrTextShape Then
                '读数据
                If Param.cmdCommand = "get:text" Then
                    getText(tempShape, pageIndex, determine)
                End If

                '写数据
                If Param.cmdCommand = "set:text" Then
                    setText(tempShape, pageIndex, determine)
                End If
            End If
        Next k
    End Function




    '////////////////////////////////////// 图片 //////////////////////////////////////////////////




    '替换图片
    Public Function replaceImage(doc, tempShape, key, typeName)
        globalData.steps = "开始logo图替换"
        '中心点
        doc.ReferencePoint = 9
        Dim centerX = tempShape.CenterX
        Dim centerY = tempShape.CenterY
        Dim SizeWidth = tempShape.SizeWidth
        Dim SizeHeight = tempShape.SizeHeight

        '返回或设置形状所在的图层
        Dim parentLayer As Layer = tempShape.Layer
        Dim imageType = 802
        Dim imagePath = Param.getExternalValue(key)
        parentLayer.Activate()

        'jpg类型
        Dim args() = Split(imagePath, ".jpg")

        If args.Count = 2 Then
            imageType = 774
        End If

        '修改图片必须是显示状态才可以
        Dim fixVisible
        If parentLayer.Visible = False Then
            fixVisible = True
            parentLayer.Visible = True
        End If

        Try
            parentLayer.Import(imagePath, imageType)
            globalData.steps = "替换" + key + "执行成功"
            globalData.state = "True"
            '重新设置图片
            Dim dfShapes = doc.Selection.Shapes
            '插入成功才删除图片
            If dfShapes.Count > 0 Then
                For j = 1 To dfShapes.Count
                    dfShapes.Item(j).Name = typeName
                    dfShapes.Item(j).SetSize(SizeWidth, SizeHeight)
                    dfShapes.Item(j).SetPositionEx(9, centerX, centerY)
                Next j
            End If
            tempShape.Delete()

            '如果修改了图片状态
            If fixVisible = True Then
                parentLayer.Visible = False
            End If

        Catch ex As Exception
            globalData.errorlog = "图片:" + imagePath + " 替换失败"
        End Try


    End Function



    '递归检测形状,并替换图片
    Public Function accessImage(doc, allShapes)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)

            Dim cdrGroupShape As cdrShapeType = 7
            Dim cdrBitmapShape As cdrShapeType = 5

            '组
            If tempShape.Type = cdrGroupShape Then
                accessImage(doc, tempShape.Shapes)
            End If

            If tempShape.Type = cdrBitmapShape Then
                If tempShape.Name = "二维码" And Param.hasValue("qrcode") Then
                    replaceImage(doc, tempShape, "qrcode", "二维码")
                ElseIf tempShape.Name = "Logo" And Param.hasValue("logo") Then
                    replaceImage(doc, tempShape, "logo", "Logo")
                ElseIf tempShape.Name = "Logo2" And Param.hasValue("logo2") Then
                    replaceImage(doc, tempShape, "logo2", "Logo2")
                End If

            End If
        Next k
    End Function




End Module