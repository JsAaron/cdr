﻿Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json


Module Module1

    Dim lineCount = 2
    Dim mainCount = 2

    Dim cmdCommand As String
    Dim cmdPath As String
    Dim cmdStylePath As String
    Dim cmdExternalData

    Class Pagesize
        Public width
        Public height
    End Class

    'json返回数据类
    Class ReturnData
        Public state = False '状态'
        Public pagesize
        Public text
        Public errorlog '错误日志
        Public steps '步骤
    End Class

    Dim globalData As ReturnData = New ReturnData()


    '数据判断类，是否分行
    Class BranchData

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
        Public Function setVisibleField()

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
        Public Function getValue(key)
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

    Dim branchObject As BranchData = New BranchData()

    '////////////////////////////////////// 功能 //////////////////////////////////////////////////


    Function GetJSON(myrange)
        Dim returnStr As String
        Dim count As Integer
        Dim colunms As Integer
        count = UBound(myrange, 1)
        colunms = UBound(myrange, 2)

        returnStr = "{["

        For i = 2 To count
            returnStr = returnStr + "{"
            For j = 1 To colunms
                returnStr = returnStr + """" & myrange(1, j) & """:""" & Replace(myrange(i, j), """", "\""") & """"

                If j <> colunms Then
                    returnStr = returnStr + ","
                End If

                If i = count And j = colunms Then
                    returnStr = returnStr + "}"
                ElseIf j = colunms Then
                    returnStr = returnStr + "},"
                End If
            Next
        Next
        returnStr = returnStr + "]}"
        GetJSON = returnStr
    End Function

    Function parseJson(jsonString As String)
        Dim strFunc, objSC
        objSC = CreateObject("ScriptControl")
        objSC.Language = "JScript"
        strFunc = "function jsonParse(s) { return eval('(' + s + ')'); }"
        objSC.AddCode(strFunc)
        parseJson = objSC.CodeObject.jsonParse(jsonString)
    End Function


    Function getKeyEnglish(str)
        Dim e = ""
        Select Case str
            Case "公司地址"
                e = "address"
            Case "地址"
                e = "address"
            Case "姓名"
                e = "name"
            Case "电话"
                e = "mobile"
            Case "网址"
                e = "url"
            Case "职务"
                e = "job"
            Case "公司英文名称"
                e = "companyname"
            Case "标语"
                e = "slogan"
            Case "公司名称"
                e = "company"
            Case "邮箱"
                e = "email"
            Case "Logo"
                e = "logo"
            Case "Logo2"
                e = "logo2"
            Case "二维码"
                e = "qrcode"
            Case "QQ"
                e = "qq"
            Case "公众号"
                e = "bjnews"
            Case "固定电话"
                e = "phone"
        End Select
        getKeyEnglish = e
    End Function

    '////////////////////////////////////// 逻辑 //////////////////////////////////////////////////

    '递归检测形状
    Public Function recurveText(doc, allShapes, infoArr)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)
            Dim cdrTextShape As cdrShapeType = 6
            Dim cdrGroupShape As cdrShapeType = 7

            '组
            If tempShape.Type = cdrGroupShape Then
                recurveText(doc, tempShape.Shapes, infoArr)
            End If

            '文字
            If tempShape.Type = cdrTextShape Then
                '读数据
                If cmdCommand = "get:text" Then
                    If tempShape.Text.Story.Text <> "" Then
                        Dim t As New ArrayList
                        t.Add(getKeyEnglish(tempShape.Name))
                        t.Add(tempShape.Text.Story.Text)
                        infoArr.Add(t)
                    End If
                End If

                '写数据
                If cmdCommand = "set:text" Then
                    Dim key As String = getKeyEnglish(tempShape.Name)
                    If cmdExternalData(key) <> "" Then
                        '是否是处理范围
                        Dim hasRange = branchObject.getScope(key)
                        If hasRange = True Then
                            '可能存在合并数据
                            tempShape.Text.Story.Replace(branchObject.getValue(key))
                        Else
                            '正常处理
                            tempShape.Text.Story.Replace(cmdExternalData(key))
                        End If
                    End If
                End If
            End If
        Next k
    End Function


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



    '文本预处理
    Public Function preproccessText(allShapes)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)
            Dim cdrTextShape As cdrShapeType = 6
            Dim cdrGroupShape As cdrShapeType = 7

            '组
            If tempShape.Type = cdrGroupShape Then
                preproccessText(tempShape.Shapes)
            End If

            '文字
            If tempShape.Type = cdrTextShape Then
                Dim key As String = getKeyEnglish(tempShape.Name)
                branchObject.setState(key)
            End If
        Next k
    End Function



    '获取文档所有页面、所有图层、所有图形对象
    Public Function accessExtractTextData(doc)

        Dim infoArr As New ArrayList
        Dim k As Integer
        Dim m As Integer
        Dim allLayers As Layers
        Dim activeLayer As Layer
        allLayers = doc.ActivePage.AllLayers


        '预处理
        If cmdCommand = "set:text" Then
            For k = 1 To allLayers.Count
                activeLayer = allLayers.Item(k)
                '字段
                branchObject.setField(activeLayer.Name)
                '获取显示的层级
                branchObject.setVisibleField()
                '文本名
                preproccessText(activeLayer.Shapes)
            Next k
        End If

        '文本
        For k = 1 To allLayers.Count
            activeLayer = allLayers.Item(k)
            recurveText(doc, activeLayer.Shapes, infoArr)
        Next k
        globalData.text = infoArr


        '设置图片/层的可见性
        If cmdCommand = "set:text" Then
            Dim visibleLayerName = branchObject.getVisibleField()
            For m = 1 To allLayers.Count
                activeLayer = allLayers.Item(m)
                '设置图片
                processImage(doc, activeLayer.Shapes)
                '设置状态，处理层级可见性
                setLayerVisible(activeLayer, visibleLayerName)
            Next m
        End If


        If cmdCommand = "get:text" Then
            globalData.steps = "获取文本信息完成"
        Else
            globalData.steps = "设置文本信息完成"
        End If

        globalData.state = "True"

    End Function


    '获取文档所有页面、所有图层、所有图形对象
    Public Function getFontNames(doc) As ArrayList
        Dim list As New ArrayList
        ' 定义循环变量
        Dim i As Integer, j As Integer, k As Integer
        Dim allPages As Pages, allShapes As Shapes, allLayers As Layers
        ' 定义临时变量
        Dim tempPage As Page, tempLayer As Layer, tempShape As Shape
        allPages = doc.Pages
        For i = 1 To allPages.Count
            tempPage = allPages.Item(i)
            ' 遍历页面中的所有图层
            allLayers = tempPage.Layers
            For j = 1 To allLayers.Count
                tempLayer = allLayers.Item(j)
                ' 遍历图层中的所有形状（对象）
                allShapes = tempLayer.Shapes
                For k = 1 To allShapes.Count
                    ' 得到这个形状
                    tempShape = allShapes.Item(k)
                    Dim cdrTextShape As cdrShapeType = 6

                    '如果是文本形状
                    If tempShape.Type = cdrTextShape Then
                        list.Add(tempShape.Text.Selection.Font)
                    End If
                Next k
            Next j
        Next i
        Return list
    End Function


    '创建字体文件
    Public Function createFontJson(doc)
        Dim names = getFontNames(doc)
        '获取当前的应用的字体
        Dim i As Integer
        Dim fontList As New ArrayList()
        Dim str As String = ""

        For i = 0 To names.Count - 1
            Dim IsExist As Boolean = True
            For j As Integer = 0 To fontList.Count - 1
                If fontList(j).ToString() = names(i) Then
                    IsExist = False
                    Exit For
                End If
            Next
            If IsExist Then
                fontList.Add(names(i))
                Dim empty As String = ""
                str = str + "{" + """fontname""" + ":""" + empty + """," + """familyname""" + ":""" + Replace(names(i), Chr(10), "") + """," + """postscriptname""" + ":""" + empty + """},"
            End If
        Next
        '去掉最后一个，
        str = Left(str, Len(str) - 1)
        str = "[" + str + "]"
        Dim p() = Split(doc.FileName, ".")
        'log("log", "开始处理字体FileName")
        Dim fs As FileStream = File.Create(doc.FilePath + p(0) + ".json")
        'log("log", "开始处理字体FileStream")
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(str)
        'log("log", "开始处理字体UTF8Encoding")
        fs.Write(info, 0, info.Length)
        'log("log", "开始处理字体Write")
        fs.Close()
        globalData.state = "True"
    End Function


    '主体方法
    'pagesize 获取页面尺寸
    'fontJson 启动字体json
    'extract 提取文本数据（名片） 
    Sub execMain(app, doc)
        '页面尺寸
        If cmdCommand = "get:pageSize" Then
            globalData.steps = "开始获取页面尺寸"
            '指定毫米
            doc.Unit = 3
            globalData.pagesize = New Pagesize()
            globalData.pagesize.width = app.ActivePage.SizeWidth
            globalData.pagesize.height = app.ActivePage.SizeHeight
            globalData.state = "True"
            globalData.steps = "获取页面尺寸完成"
        End If

        '获取当前文字的json
        If cmdCommand = "get:fontJson" Then
            createFontJson(doc)
        End If

        '获取数据
        If cmdCommand = "get:text" Then
            globalData.steps = "开始获取页面文本内容"
            accessExtractTextData(doc)
        End If

        If cmdCommand = "set:text" Then
            globalData.steps = "开始设置页面文本内容"
            accessExtractTextData(doc)
        End If

    End Sub


    Sub checkLine(app)
        Try
            If Len(cmdPath) > 2 Then
                globalData.steps = "开始打开文档"
                app.OpenDocument(cmdPath)
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
            checkLine(app)
            Exit Sub
        End Try

        Try
            Dim doc As Document = app.ActiveDocument
            If app.Documents.Count = 0 Then
                globalData.errorlog = "没有找到活动文档"
                Exit Sub
            End If

            globalData.steps = "文档打开完成"

            '如果只是打开文档，推出
            If cmdCommand = "open" Then
                globalData.state = "True"
                Exit Sub
            End If

            '加载样式
            If cmdCommand = "set:style" Then
                globalData.steps = "文档加载样式开始"
                If Len(cmdStylePath) = 0 Then
                    globalData.state = "False"
                    globalData.errorlog = "必须传递样式路径参数"
                    Exit Sub
                End If
                doc.LoadStyleSheet(cmdStylePath)
                globalData.state = "True"
                globalData.steps = "文档加载样式完成"
                Exit Sub
            End If

            execMain(app, doc)

        Catch ex As Exception
            ' log("error", "CorelDRAW执行功能错误")
            Exit Sub
        End Try

    End Sub


    Function decodeURI(cmdExternalData, key)
        If cmdExternalData(key) <> "" Then
            Dim e = CreateObject("MSScriptControl.ScriptControl")
            e.Language = "javascript"
            cmdExternalData(key) = e.Eval("decodeURI('" & cmdExternalData(key) & "')")
        End If
    End Function

    Public Function parseCommand()
        Dim args() = Split(Command, " ")
        Dim count = args.Count
        globalData.steps = "开始解析参数"
        cmdCommand = args(0)
        If cmdCommand = "open" Then
            If count = 2 Then
                cmdPath = args(1)
            End If
        ElseIf cmdCommand = "get:pageSize" Then
            If count = 2 Then
                cmdPath = args(1)
            End If
        ElseIf cmdCommand = "get:fontJson" Then
            If count = 2 Then
                cmdPath = args(1)
            End If
        ElseIf cmdCommand = "get:text" Then
            If count = 2 Then
                cmdPath = args(1)
            End If
        ElseIf cmdCommand = "set:text" Then
            If count = 1 Then
                globalData.errorlog = "必须传递设置参数"
            ElseIf count = 2 Then
                cmdExternalData = JsonConvert.DeserializeObject(args(1))
                decodeURI(cmdExternalData, "logo")
                decodeURI(cmdExternalData, "qrcode")
            ElseIf count = 3 Then
                cmdExternalData = JsonConvert.DeserializeObject(args(1))
                decodeURI(cmdExternalData, "logo")
                decodeURI(cmdExternalData, "qrcode")
                cmdPath = args(2)
            End If
        ElseIf cmdCommand = "set:style" Then
            '参数不够
            If count = 1 Then
                globalData.errorlog = "必须传递样式路径参数"
            End If

            If count = 2 Then
                cmdStylePath = args(1)
            End If

            '设置样式
            If count = 3 Then
                cmdPath = args(1)
            End If

        End If

        globalData.steps = "解析参数完成"
    End Function

    Sub Main()
        Console.OutputEncoding = Encoding.UTF8

        '如果有外部命令
        If Len(Command) > 0 Then
            parseCommand()
        End If


        globalData.steps = "开始连接CorelDRAW"

        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)
        app.Visible = True

        globalData.steps = "连接CorelDRAW成功"

        Try
            checkLine(app)
        Catch ex As Exception
            If mainCount = 0 Then
                globalData.errorlog = "CorelDRAW软件无法链接"
                Exit Sub
            End If
            mainCount = mainCount - 1
            Threading.Thread.Sleep(3000)
            checkLine(app)
            Exit Sub
        End Try

        Console.WriteLine(JsonConvert.SerializeObject(globalData))

        ' MsgBox(1)

    End Sub

End Module

