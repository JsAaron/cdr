Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq



Module App

    Dim lineCount = 2
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


    '===================== 导入 =====================

    Function setImport(doc As Document)
        Dim parentLayer As Layer = doc.ActiveLayer
        parentLayer.Activate()
        '修改图片必须是显示状态才可以
        Dim fixVisible
        If parentLayer.Visible = False Then
            fixVisible = True
            parentLayer.Visible = True
        End If

        Try
            If cmdExternalData("type") = "" Then
                parentLayer.Import(cmdExternalData("path"), 0)
            Else
                parentLayer.Import(cmdExternalData("path"), cmdExternalData("type"))
            End If

            '如果修改了图片状态
            If fixVisible = True Then
                parentLayer.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Function


    '===================== 打印 =====================


    '设置打印普通参数
    Function setPrintValue(doc, key, value)
        Dim active_doc As Document = doc

        If key = "Collate" Then
            active_doc.PrintSettings.Collate = value
        End If

        If key = "FileName" Then
            active_doc.PrintSettings.FileName = value
        End If

        If key = "Copies" Then
            active_doc.PrintSettings.Copies = value
        End If

        If key = "PrintRange" Then
            active_doc.PrintSettings.PrintRange = value
        End If

        If key = "PageRange" Then
            active_doc.PrintSettings.PageRange = value
        End If

        If key = "ShowDialog" Then
            If value = True Then
                active_doc.PrintSettings.ShowDialog()
            End If
        End If

        If key = "PageSet" Then
            active_doc.PrintSettings.PageSet = value
        End If

        If key = "PaperOrientation" Then
            active_doc.PrintSettings.PaperOrientation = value
        End If

        If key = "PrintToFile" Then
            active_doc.PrintSettings.PrintToFile = value
        End If

        If key = "SelectPrinter" Then
            active_doc.PrintSettings.SelectPrinter(value)
        End If

        If key = "PaperSize" Then
            active_doc.PrintSettings.PaperSize = value
        End If

        '保存样式方法
        If key = "Save" Then
            active_doc.PrintSettings.Save(value)
        End If

        '加载样式
        If key = "Load" Then
            active_doc.PrintSettings.Load(value)
        End If

        '将打印设置重置为默认值
        If key = "Reset" Then
            If value = True Then
                active_doc.PrintSettings.Reset()
            End If
        End If

    End Function


    '设置打印数组参数
    Function setPrintArrayValue(doc, key, value)
        Dim active_doc As Document = doc
        If key = "SetPaperSize" Then
            Dim v As JArray = value
            active_doc.PrintSettings.SetPaperSize(v.First.ToString(), v.Last.ToString())
        End If

        If key = "SetCustomPaperSize" Then
            Dim v As JArray = value
            active_doc.PrintSettings.SetCustomPaperSize(v.Item(0).ToString(), v.Item(1).ToString(), v.Item(2).ToString())
        End If

    End Function


    '设置打印对象参数
    Function setPrintObjectValue(doc, key, value)
        Dim active_doc As Document = doc
        If key = "Printer" Then
            Dim v As JObject = value
            '打印机的颜色输出
            If v("ColorEnabled") = "True" Then
                ' active_doc.PrintSettings.Printer.ColorEnabled = True
            End If
        End If


    End Function


    Function setPrint(doc, key, value)
        Dim active_doc As Document = doc
        Dim JVType = TypeName(value)

        '普通类型 布尔，字符串，数字
        If JVType = "JValue" Then
            '如果没有参数
            If value.ToString().Length = 0 Then
                Return False
            Else
                Dim typeValue
                Select Case value.ToString()
                    Case "0"
                        typeValue = 0
                    Case "True"
                        typeValue = True
                    Case "False"
                        typeValue = False
                    Case Else
                        typeValue = value.ToString()
                End Select
                setPrintValue(active_doc, key, typeValue)
            End If
        End If

        '数组类型
        If JVType = "JArray" Then
            Dim jvArray As JArray = value
            If jvArray.Count > 0 Then
                setPrintArrayValue(active_doc, key, value)
            End If
        End If

        '对象类型
        If JVType = "JObject" Then
            Dim jvObject As JObject = value
            '如果有属性
            If jvObject.Count > 0 Then
                setPrintObjectValue(active_doc, key, value)
            End If
        End If

        'Console.WriteLine(JVType)

    End Function


    Function getSettingsValue(key)
        Return cmdExportSettings(key).ToString()
    End Function


    '找到母版页面
    Function findMasterLayer(doc, layerName)
        Dim activeDoc As Document = doc
        For k = 1 To activeDoc.MasterPage.Layers.Count
            Dim curLayer As Layer = activeDoc.MasterPage.Layers.Item(k)
            If curLayer.Name = layerName Then
                Return curLayer
            End If
        Next k
        Return False
    End Function


    '设置状态
    Function setExportImageStatus(layer, status)
        If TypeName(layer) = "Layer" Then
            layer.Visible = status
            layer.Printable = status
        End If
    End Function


    Function hasShape(currPage)
        Dim allLayers = currPage.AllLayers
        For m = 1 To allLayers.Count
            Dim curLayer = allLayers.Item(m)
            If curLayer.Shapes.Count > 0 Then
                Return True
            End If
        Next m
        Return False
    End Function

    '导出图片
    Function exportImage(pageIndex, doc)


        Dim activeDoc As Document = doc

        Dim FileName = getSettingsValue("FileName")
        Dim ImageType = getSettingsValue("ImageType")
        Dim keepCanvas = getSettingsValue("keepCanvas")

        Dim Width
        Dim Height
        Dim coverLayer
        Dim footerLayer
        Dim middleLayer
        Dim exportName As String

        '激活当前到导出页面
        Dim currPage As Page = activeDoc.Pages.Item(pageIndex)
        currPage.Activate()


        Dim mode = getSettingsValue("mode")

        '分离模式，有首位
        If mode = 1 Then

            '如果是封面
            If pageIndex = 1 Then
                coverLayer = findMasterLayer(doc, "封面导出")
                setExportImageStatus(coverLayer, True)
                exportName = "cover"
                Width = getSettingsValue("CoverWidth")
                Height = getSettingsValue("CoverHeight")
            ElseIf pageIndex = activeDoc.Pages.Count Then
                '如果是封尾
                footerLayer = findMasterLayer(doc, "封底导出")
                setExportImageStatus(footerLayer, True)
                exportName = "back"
                Width = getSettingsValue("BackWidth")
                Height = getSettingsValue("BackHeight")
            Else
                '中间页面
                Dim v = pageIndex / 2
                Dim s = 0
                For i = 1 To Len(v)
                    If Mid(v, i, 1) = "." Then
                        s = s + 1
                    End If
                Next

                If s = 1 Then
                    Return False
                End If

                middleLayer = findMasterLayer(doc, "对页导出")
                setExportImageStatus(middleLayer, True)
                Width = getSettingsValue("MiddleWidth")
                Height = getSettingsValue("MiddleHeight")
                exportName = pageIndex.ToString() + "-" + (pageIndex + 1).ToString()
            End If

        End If


        '合并模式
        If mode = 2 Then
            '全部页面
            Dim v = pageIndex / 2
            Dim s = 0
            For i = 1 To Len(v)
                If Mid(v, i, 1) = "." Then
                    s = s + 1
                End If
            Next
            If s = 0 Then
                Return False
            End If
            middleLayer = findMasterLayer(doc, "对页导出")
            setExportImageStatus(middleLayer, True)
            Width = getSettingsValue("MiddleWidth")
            Height = getSettingsValue("MiddleHeight")
            exportName = pageIndex.ToString() + "-" + (pageIndex + 1).ToString()
        End If


        '正常模式
        If mode = 3 Then
            middleLayer = findMasterLayer(doc, "对页导出")
            setExportImageStatus(middleLayer, True)
            Width = getSettingsValue("MiddleWidth")
            Height = getSettingsValue("MiddleHeight")
            exportName = pageIndex.ToString()
        End If


        '独立模式，单独生成指定的图片名
        If mode = 4 Then
            If keepCanvas Then
                middleLayer = findMasterLayer(doc, "对页导出")
                setExportImageStatus(middleLayer, True)
            End If
            Width = getSettingsValue("MiddleWidth")
            Height = getSettingsValue("MiddleHeight")
            exportName = getSettingsValue("ImageName")
        End If

        Dim filePath = FileName + "\" + exportName + ".jpg"

        Try

            Dim structOpts As StructExportOptions = New StructExportOptions
            structOpts.Overwrite = True
            structOpts.SizeX = Width
            structOpts.SizeY = Height
            activeDoc.Export(filePath, 774, 1, structOpts)
            'Dim efilter As ExportFilter = activeDoc.ExportBitmap(filePath, 774, 1, ImageType, Width, Height, 0, 0, 0, True, True, False, False, 8)
        Catch ex As Exception

        End Try


        If mode = 1 Then
            '复位
            If pageIndex = 1 Then
                setExportImageStatus(coverLayer, False)
            ElseIf pageIndex = activeDoc.Pages.Count Then
                setExportImageStatus(footerLayer, False)
            Else
                setExportImageStatus(middleLayer, False)
            End If
        End If


        If mode = 2 Or mode = 3 Then
            setExportImageStatus(middleLayer, False)
        End If

        If keepCanvas Then
            setExportImageStatus(middleLayer, False)
        End If

    End Function



    '================== 插入图片处理 =====================

    '等比缩小
    Function equalDecrease(theimage, imgWidth, imgHeight, tgwidth, tgheight, ratio)
        For k = 1 To 1000
            Dim value = ratio - (k / 1000)
            imgWidth = imgWidth * value
            imgHeight = imgHeight * value
            If imgWidth < tgwidth And imgHeight < tgheight Then
                theimage.SizeWidth = imgWidth
                theimage.SizeHeight = imgHeight
                Return True
            End If
        Next k
    End Function


    '等比放大
    Function equalZoom(theimage, imgWidth, imgHeight, tgwidth, tgheight, ratio)
        For k = 1 To 1000
            Dim value = ratio + (k / 1000)
            imgWidth = imgWidth * value
            imgHeight = imgHeight * value
            If imgWidth >= tgwidth Or imgHeight >= tgheight Then
                theimage.SizeWidth = imgWidth
                theimage.SizeHeight = imgHeight
                Return True
            End If
        Next k
    End Function


    '图片尺寸
    Function coverSize(theimage, tgwidth, tgheight)

        Dim imgWidth = theimage.SizeWidth
        Dim imgHeight = theimage.SizeHeight

        '图片正方形
        If imgWidth = imgHeight Then
            '尺寸是最小边就行了
            If tgwidth > tgheight Or tgwidth = tgheight Then
                theimage.SizeWidth = tgheight
                theimage.SizeHeight = tgheight
                Return True
            End If
            If tgwidth < tgheight Then
                theimage.SizeWidth = tgwidth
                theimage.SizeHeight = tgwidth
                Return True
            End If
        End If

        '宽高都没有溢出，等比放大
        If imgWidth < tgwidth And imgHeight < tgheight Then
            Return equalZoom(theimage, imgWidth, imgHeight, tgwidth, tgheight, 1)
        End If

        '如果宽/高溢出,缩小
        If imgWidth > tgwidth Or imgHeight > tgheight Then
            Return equalDecrease(theimage, imgWidth, imgHeight, tgwidth, tgheight, 1)
        End If

    End Function


    Function containSize(obj, tgwidth, tgheight)
        Dim wfactor = tgwidth / obj.SizeWidth
        Dim hfactor = tgheight / obj.SizeHeight
        Dim factor = hfactor
        If wfactor < hfactor Then
            factor = wfactor
        End If
        Dim newwidth = factor * obj.SizeWidth
        Dim newheight = factor * obj.SizeHeight
        obj.SizeWidth = newwidth
        obj.SizeHeight = newheight
    End Function


    '设置图片对象尺寸
    Function setImageShapeSize(theimage, Shape, groupName)
        If InStr(groupName, "图标") > 0 Then
            Dim iconWidth = getSettingsValue("iconWidth")
            Dim iconHeight = getSettingsValue("iconHeight")
            containSize(theimage, iconWidth, iconHeight)
        Else
            Dim SizeWidth = Shape.SizeWidth
            Dim SizeHeight = Shape.SizeHeight

            '如果有指定X坐标
            Dim coordX = getSettingsValue("LeftX")
            If Len(coordX) > 0 Then
                Dim w = Shape.LeftX + Shape.SizeWidth
                Dim h = Shape.TopY - Shape.SizeHeight
                If coordX > Shape.LeftX And coordX < w Then
                    Dim diffX = coordX - Shape.LeftX
                    SizeWidth = Shape.SizeWidth - diffX
                End If
            End If

            Dim coordY = getSettingsValue("TopY")
            If Len(coordY) > 0 Then
                Dim diffY = Shape.TopY - coordY
                SizeHeight = Shape.SizeHeight - diffY
            End If

            coverSize(theimage, SizeWidth, SizeHeight)
        End If
    End Function


    '找到插入的图片对象
    Function findImortImage(doc, imageName)
        For k = 1 To doc.Selection.Shapes.Count
            Dim theimage = doc.Selection.Shapes.Item(k)
            If imageName = theimage.Name Then
                Return theimage
            End If
        Next k
    End Function

    '删除占位符
    Function deleteImageShape(Shape)
        Dim clipShapes = Shape.PowerClip.Shapes
        For n = 1 To clipShapes.Count
            Dim clipSshape = clipShapes.Item(n)
            If clipSshape.Name = "delete" Then
                clipSshape.Delete()
            End If
        Next n
    End Function


    '处理图片
    Function prcessImage(theimage, Shape, shapeName)
        setImageShapeSize(theimage, Shape, shapeName)
        theimage.AddToPowerClip(Shape, -1)
        '对齐
        Dim coordX = getSettingsValue("LeftX")
        If Len(coordX) > 0 Then
            Dim w = Shape.LeftX + Shape.SizeWidth
            Dim h = Shape.TopY - Shape.SizeHeight
            If coordX > Shape.LeftX And coordX < w Then
                Dim diffX = coordX - Shape.LeftX
                Dim SizeWidth = Shape.SizeWidth - diffX
                theimage.CenterX = coordX + (SizeWidth / 2)
            End If
        End If

        Dim coordY = getSettingsValue("TopY")
        If Len(coordY) > 0 Then
            Dim diffY = Shape.TopY - coordY
            Dim SizeHeight = Shape.SizeHeight - diffY
            theimage.CenterY = coordY - (SizeHeight / 2)
        End If


        deleteImageShape(Shape)
    End Function

    '处理图片，'
    '最大查找3次
    Dim findImageCount = 3
    Function refImage(doc, imageName, Shape, shapeName, rotationAngle)

        If findImageCount = 0 Then
            Exit Function
        End If

        Dim theimage = findImortImage(doc, imageName)
        If TypeName(theimage) IsNot "Nothing" Then
            '设置旋转角度
            If rotationAngle > 0 Then
                theimage.RotationAngle = rotationAngle
            End If
            prcessImage(theimage, Shape, shapeName)
            Exit Function
        Else
            '如果没有找到图片，休眠后开始循环查找
            Sleep(500)
            findImageCount = findImageCount - 1
            refImage(doc, imageName, Shape, shapeName, rotationAngle)
        End If
    End Function

    '插入图片
    Function insertImage(doc)

        Dim activeDoc As Document = doc
        Dim FileName = getSettingsValue("FileName")
        Dim parentName = getSettingsValue("parentName")
        Dim shapeName = getSettingsValue("shapeName")
        Dim imageName = getSettingsValue("imageName")
        Dim layerName = getSettingsValue("layerName")
        Dim StaticID = getSettingsValue("StaticID")
        Dim rotationAngle = getSettingsValue("RotationAngle")

        Dim activeLayer = activeDoc.ActivePage.AllLayers.Find(layerName)

        '修改图片必须是显示状态才可以
        Dim fixVisible
        If activeLayer.Visible = False Then
            fixVisible = True
            activeLayer.Visible = True
        End If

        '没有包含容器
        If layerName = parentName Then
            Dim shapes = activeLayer.FindShapes(shapeName)
            For i = 1 To shapes.Count
                Dim shape = shapes.Item(i)
                If shape.StaticID = StaticID Then
                    activeLayer.Import(FileName, 0)
                    refImage(doc, imageName, shape, shapeName, rotationAngle)
                End If
            Next i
        Else
            '中间图
            '保证找到的对象一定是目标
            Dim shapeGroup = activeLayer.Shapes.FindShapes(shapeName)
            For m = 1 To shapeGroup.Count
                Dim shape = shapeGroup.Item(m)
                If shape.StaticID = StaticID Then
                    activeLayer.Import(FileName, 0)
                    refImage(doc, imageName, shape, shapeName, rotationAngle)
                End If
            Next m
        End If


        '如果修改了图片状态
        If fixVisible = True Then
            activeLayer.Visible = False
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


            '保存文档
            If Param.cmdCommand = "SaveAsCopy" Then
                doc.SaveAsCopy(Param.cmdSavePath)
                Exit Sub
            End If


            '插入图片
            If Param.cmdCommand = "insert-image" Then
                insertImage(doc)
            End If

            '导出图片
            If Param.cmdCommand = "export-image" Then

                '缓存区域
                Dim cacheIndex = doc.ActivePage.Index
                Dim pageIndex = getSettingsValue("Page")

                If pageIndex = "all" Then
                    For i = 1 To doc.Pages.Count
                        exportImage(doc.Pages.Item(i).Index, doc)
                    Next
                Else
                    exportImage(pageIndex, doc)
                End If

                '恢复之前的处理页面
                doc.Pages.Item(cacheIndex).Activate()

                Exit Sub
            End If


            '独立命令，打印
            If Param.cmdCommand = "print" Then
                Dim settingsObject As JObject = cmdPrintSettings
                For Each item In settingsObject
                    setPrint(doc, item.Key, item.Value)
                Next
                doc.PrintOut()
                Exit Sub
            End If


            '独立命令，保存文件
            If Param.cmdCommand = "save" Then
                doc.SaveAs(cmdExternalData("path"))
                Exit Sub
            End If


            '如果是单独的入命令
            If Param.cmdCommand = "import" Then
                setImport(doc)
                Exit Sub
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

        'Console.WriteLine(cmdExternalData)

        '没有解析错误的情况
        If Len(globalData.errorlog) = 0 Then
            globalData.steps = "开始连接CorelDRAW"
            Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
            Dim app As Application = Activator.CreateInstance(pia_type)
            app.Visible = True
            globalData.steps = "连接CorelDRAW成功"
            openLink(app)

        End If

        'Console.WriteLine(globalData.retrunData())


    End Sub

End Module

