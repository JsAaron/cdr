Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq




Module App


    Dim lineCount = 2
    Dim mainCount = 2


    Class Pagesize
        Public width
        Public height
    End Class


    '获取文档所有页面、所有图层、所有图形对象
    Public Function accessExtractTextData(doc As Document, page As Page, determine As Determine)

        Dim k As Integer
        Dim m As Integer
        '当前页面的层
        Dim curLayer As Layer
        '当前页面所有层
        Dim allLayers As Layers = page.AllLayers
        Dim pageIndex = page.Index

        '预处理
        globalData.steps = "预处理"
        For k = 1 To allLayers.Count
            curLayer = allLayers.Item(k)
            '初始化预处理
            determine.init(curLayer.Name, curLayer.Shapes, pageIndex)
        Next k

        '读/取操作
        globalData.steps = "文本/图像读取操作"
        For k = 1 To allLayers.Count
            curLayer = allLayers.Item(k)
            Inputs.accesstShape(doc, curLayer.Shapes, determine, pageIndex)
        Next k

        '设置图片/层的可见性
        globalData.steps = "设置图片/层级可见性"
        If Param.cmdCommand = "set:text" Then
            Dim visibleLayerName = determine.getVisibleField()
            For m = 1 To allLayers.Count
                curLayer = allLayers.Item(m)
                '设置图片
                Inputs.accessImage(doc, curLayer.Shapes)
                '设置状态，处理层级可见性
                determine.setLayerVisible(curLayer, visibleLayerName)
            Next m
        End If

        globalData.steps = "文本处理完成"
        globalData.state = "True"

    End Function



    '========================================== 功能调用 ==========================================


    '文本处理
    Function fn_accessText(doc, page, determine)
        If Param.cmdCommand = "get:text" Or Param.cmdCommand = "set:text" Then
            accessExtractTextData(doc, page, determine)
            globalData.state = "True"
            Return True
        End If
    End Function


    '设置页面尺寸
    Function fn_pageSize(app, doc)
        If Param.cmdCommand = "get:pageSize" Then
            globalData.steps = "开始获取页面尺寸"
            '指定毫米
            doc.Unit = 3
            globalData.pagesize = New Pagesize()
            globalData.pagesize.width = app.ActivePage.SizeWidth
            globalData.pagesize.height = app.ActivePage.SizeHeight
            globalData.state = "True"
            globalData.steps = "获取页面尺寸完成"
            Return True
        End If
    End Function


    '打开文档功能
    Function fn_open()
        '如果只是打开文档，推出
        If Param.cmdCommand = "open" Then
            globalData.state = "True"
            Return True
        End If
    End Function


    '设置样式功能
    Function fn_setStyle(doc)
        '加载样式
        If Param.cmdCommand = "set:style" Then
            globalData.steps = "文档加载样式开始"
            If Len(Param.cmdStylePath) = 0 Then
                globalData.state = "False"
                globalData.errorlog = "必须传递样式路径参数"
                Return True
            End If
            doc.LoadStyleSheet(Param.cmdStylePath)
            globalData.state = "True"
            globalData.steps = "文档加载样式完成"
            Return True
        End If
    End Function


    '获取当前文字的json
    Function fn_fontJson(doc)
        If Param.cmdCommand = "get:fontJson" Then
            Utils.createFontJson(doc)
            globalData.state = "True"
            Return True
        End If
    End Function


    '===================== 功能链接 =====================

    '开始执行操作
    Function execFn(app, doc, page, determine)

        globalData.steps = "文档打开完成"

        '打开文档
        If fn_open() = True Then
            Exit Function
        End If

        '设置样式
        If fn_setStyle(doc) = True Then
            Exit Function
        End If

        '页面尺寸
        If fn_pageSize(app, doc) = True Then
            Exit Function
        End If

        '获取字体文件
        If fn_fontJson(doc) = True Then
            Exit Function
        End If

        '文本处理
        If fn_accessText(doc, page, determine) = True Then
            Exit Function
        End If
    End Function


    '===================== 建立链接 =====================


    '建立链接
    Sub openLink(app As Application)
        Try
            If Len(Param.cmdPath) > 2 Then
                globalData.steps = "开始打开文档"
                app.OpenDocument(Param.cmdPath)
            End If

            Dim doc As Document = app.ActiveDocument
            If app.Documents.Count = 0 Then
                globalData.errorlog = "没有找到活动文档"
                Exit Sub
            End If

        Catch ex As Exception
            If lineCount = 0 Then
                globalData.errorlog = "CorelDRAW打开文档错误"
                Exit Sub
            End If
            lineCount = lineCount - 1
            Threading.Thread.Sleep(3000)
            openLink(app)
            Exit Sub
        End Try

        Try

            Dim doc As Document = app.ActiveDocument
            Dim pages = doc.Pages

            If app.Documents.Count = 0 Then
                globalData.errorlog = "没有找到活动文档"
                Exit Sub
            End If

            globalData.totalPages = pages.Count


            '单独设置字体
            If Param.cmdCommand = "set:font" Then
                globalData.steps = "设置字体"

                '没有任何焦点
                If TypeName(app.ActiveShape) = "Nothing" Then
                    globalData.errorlog = "没有任何选中"
                    globalData.steps = "设置字体失败"
                Else
                    '如果是文本类型
                    If app.ActiveShape.Type = 6 Then
                        app.ActiveShape.Text.Story.Font = Param.cmdFontName
                        globalData.steps = "设置字体完成"
                        globalData.state = "True"
                        globalData.textOverflow = app.ActiveShape.Text.Overflow
                    Else
                        globalData.errorlog = "没有选中文本类型"
                        globalData.steps = "设置字体失败"
                    End If
                End If
                Exit Sub
            End If

            '遍历页面
            If cmdActivePagte <> "" Then
                Dim determine As Determine = New Determine()
                execFn(app, doc, pages.Item(cmdActivePagte), determine)
            Else
                For i = 1 To pages.Count
                    Dim determine As Determine = New Determine()
                    execFn(app, doc, pages.Item(i), determine)
                Next
            End If

        Catch ex As Exception
            Console.WriteLine(ex)
            Exit Sub
        End Try

    End Sub


    Sub Main()

        Console.OutputEncoding = Encoding.UTF8

        '如果有外部命令
        If Len(Command) > 0 Then
            Try
                parseCommand(Command)
            Catch ex As Exception
                globalData.errorlog = "命令参数解析错误"
            End Try

        End If

        '没有解析错误的情况
        If Len(globalData.errorlog) = 0 Then
            globalData.steps = "开始连接CorelDRAW"
            Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
            Dim app As Application = Activator.CreateInstance(pia_type)
            app.Visible = True
            globalData.steps = "连接CorelDRAW成功"
            openLink(app)
        End If

        Console.WriteLine(globalData.retrunData())

        ' MsgBox(1)

    End Sub

End Module

