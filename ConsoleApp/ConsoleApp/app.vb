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


    '////////////////////////////////////// 逻辑 //////////////////////////////////////////////////


    '替换图片
    Public Function replaceImage(doc, tempShape, type, typeName)
        globalData.steps = "开始logo图替换"
        '中心点
        doc.ReferencePoint = 9
        Dim centerX = tempShape.CenterX
        Dim centerY = tempShape.CenterY
        Dim SizeWidth = tempShape.SizeWidth
        Dim SizeHeight = tempShape.SizeHeight

        Dim activeLayer As Layer = tempShape.Layer

        Dim imageType = 802
        activeLayer.Activate()

        'jpg类型
        Dim args() = Split(cmdExternalData(type), ".jpg")
        If args.Count = 2 Then
            imageType = 774
        End If

        activeLayer.Import(cmdExternalData(type), imageType)
        globalData.steps = "替换" + type + "执行成功"

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
    End Function




    '递归检测形状,并替换图片
    Public Function processImage(doc, allShapes)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)

            Dim cdrGroupShape As cdrShapeType = 7
            Dim cdrBitmapShape As cdrShapeType = 5

            '组
            If tempShape.Type = cdrGroupShape Then
                processImage(doc, tempShape.Shapes)
            End If

            If tempShape.Type = cdrBitmapShape Then
                '二维码
                If tempShape.Name = "二维码" And cmdExternalData("qrcode") <> "" Then
                    replaceImage(doc, tempShape, "qrcode", "二维码")
                End If

                'logo图片
                If tempShape.Name = "Logo" And cmdExternalData("logo") <> "" Then
                    replaceImage(doc, tempShape, "logo", "Logo")
                End If
            End If

        Next k
    End Function


    Function setVisible(activeLayer, name, visibleLayerName)
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
        If Param.cmdCommand = "set:text" Then
            For k = 1 To allLayers.Count
                curLayer = allLayers.Item(k)
                '初始化预处理
                determine.init(curLayer.Name, curLayer.Shapes)
            Next k
        End If


        '文本读取操作
        For k = 1 To allLayers.Count
            curLayer = allLayers.Item(k)
            Inputs.accessText(doc, curLayer.Shapes, determine, pageIndex)
        Next k


        If True Then
            Return True
        End If



        '设置图片/层的可见性
        If cmdCommand = "set:text" Then
            Dim visibleLayerName = determine.getVisibleField()
            For m = 1 To allLayers.Count
                curLayer = allLayers.Item(m)
                '设置图片
                processImage(doc, curLayer.Shapes)
                '设置状态，处理层级可见性
                setLayerVisible(curLayer, visibleLayerName)
            Next m
        End If


        If cmdCommand = "get:text" Then
            globalData.steps = "获取文本信息完成"
        Else
            globalData.steps = "设置文本信息完成"
        End If

        globalData.state = "True"

    End Function






    '===================== 功能调用 =====================


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

            For i = 1 To pages.Count
                Dim determine As Determine = New Determine()
                execFn(app, doc, pages.Item(i), determine)
            Next

        Catch ex As Exception
            ' log("error", "CorelDRAW执行功能错误")
            Exit Sub
        End Try

    End Sub


    Sub Main()
        Console.OutputEncoding = Encoding.UTF8
        '如果有外部命令
        If Len(Command) > 0 Then
            parseCommand(Command)
        End If

        globalData.steps = "开始连接CorelDRAW"
        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)
        app.Visible = True
        globalData.steps = "连接CorelDRAW成功"

        openLink(app)

        Console.WriteLine(globalData.retrunData())

        MsgBox(1)
    End Sub

End Module

