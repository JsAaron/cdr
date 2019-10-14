Imports Corel.Interop.VGCore

'input输入处理

Module Inputs

    '获取文本
    Private Function getText(tempShape, pageIndex)


        If tempShape.Text.Story.Text <> "" Then

            Dim key = Utils.getKeyEnglish(tempShape.Name)
            If Len(key) > 0 Then
                globalData.setValue(pageIndex, key, tempShape.Text.Story.Text)
            Else
                '    Console.WriteLine("找不到对应的命名：" & tempShape.Name)
            End If
        End If
    End Function


    '设置文本
    Private Function setText(tempShape, determine, pageIndex)
        Dim key As String = Utils.getKeyEnglish(tempShape.Name)
        If Param.cmdExternalData(key) <> "" Then
            '是否是处理范围
            Dim hasRange = determine.getScope(key)
            If hasRange = True Then
                '可能存在合并数据
                tempShape.Text.Story.Replace(determine.getValue(key))
            Else
                '正常处理
                tempShape.Text.Story.Replace(Param.getSetdata(key, pageIndex))
            End If
        End If
    End Function


    '递归检测形状
    Public Function accessText(doc, allShapes, determine, pageIndex)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)
            Dim cdrTextShape As cdrShapeType = 6
            Dim cdrGroupShape As cdrShapeType = 7

            '组
            If tempShape.Type = cdrGroupShape Then
                accessText(doc, tempShape.Shapes, determine, pageIndex)
            End If

            '文字
            If tempShape.Type = cdrTextShape Then
                '读数据
                If Param.cmdCommand = "get:text" Then
                    getText(tempShape, pageIndex)
                End If

                '写数据
                If Param.cmdCommand = "set:text" Then
                    setText(tempShape, determine, pageIndex)
                End If
            End If
        Next k
    End Function


End Module
