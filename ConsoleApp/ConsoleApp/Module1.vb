Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json


Module Module1

    Dim lineCount = 2
    Dim mainCount = 2

    Dim cmdCommand As String
    Dim cmdPath As String
    Dim cmdExternalData

    Class Pagesize
        Public width
        Public height
    End Class

    'json返回数据类
    Class CreateData
        Public state = False '状态'
        Public pagesize
        Public text
        Public errorlog '错误日志
        Public steps '步骤
    End Class

    Dim globalData As CreateData = New CreateData()


    '////////////////////////////////////// 功能 //////////////////////////////////////////////////

    Public Function Debug(value)
        '   Console.WriteLine("debug: " & value)
    End Function


    Public Function log(name, value)
        Console.WriteLine("{""" + name + """:""" + value + """}")
    End Function

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
            Case "公司名称"
                e = "company"
            Case "邮箱"
                e = "email"
            Case "Logo"
                e = "logo"
            Case "二维码"
                e = "qrcode"
            Case "QQ"
                e = "qq"
            Case "公众号"
                e = "bjnews"
            Case "固定电话"
                e = "phone"
            Case Else
                e = str
        End Select
        getKeyEnglish = e
    End Function



    '////////////////////////////////////// 逻辑 //////////////////////////////////////////////////

    '递归检测形状
    Public Function recurveShape(doc, allShapes, infoArr, ActiveLayer)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)
            Dim cdrTextShape As cdrShapeType = 6
            Dim cdrGroupShape As cdrShapeType = 7
            Dim cdrBitmapShape As cdrShapeType = 5
            Dim cdrJPEG As cdrFilter = 774

            '位图
            If tempShape.Type = cdrBitmapShape Then
                'logo图
                If tempShape.Name = "Logo" Then
                    '位图处理
                    If cmdCommand = "get:text" Then
                        '获取
                    ElseIf cmdCommand = "set:text" Then
                        Dim io As New StructImportOptions
                        With io
                            .Mode = 0
                            .CombineMultilayerBitmaps = False
                            .DetectWatermark = True
                            .ExtractICCProfile = False
                            .ICCFileName = "C:\Test.icc"
                            .ImageIndex = 3
                            .LinkBitmapExternally = False
                            .MaintainLayers = True
                            .UseOPILinks = False
                        End With
                        tempShape.Delete()
                        Dim a = ActiveLayer.ImportEx("C:\Users\Administrator\Desktop\Logo.jpg", 774, io)
                    End If
                End If
            End If

            '组
            If tempShape.Type = cdrGroupShape Then
                recurveShape(doc, tempShape.Shapes, infoArr, ActiveLayer)
            End If

            '文字
            If tempShape.Type = cdrTextShape Then
                '如果有值
                If tempShape.Text.Story.Text <> "" Then

                    If cmdCommand = "get:text" Then
                        Dim t As New ArrayList
                        t.Add(getKeyEnglish(tempShape.Name))
                        t.Add(tempShape.Text.Story.Text)
                        infoArr.Add(t)
                    ElseIf cmdCommand = "set:text" Then
                        Dim key As String = getKeyEnglish(tempShape.Name)
                        If cmdExternalData(key) <> "" Then
                            tempShape.Text.Story.Replace(cmdExternalData(key))
                        End If
                    End If

                End If
            End If
        Next k
    End Function


    '获取文档所有页面、所有图层、所有图形对象
    Public Function accessExtractTextData(doc)
        Dim infoArr As New ArrayList
        Dim k As Integer
        Dim allLayers As Layers
        Dim tempLayer As Layer
        allLayers = doc.ActivePage.AllLayers
        For k = 1 To allLayers.Count
            tempLayer = allLayers.Item(k)
            recurveShape(doc, tempLayer.Shapes, infoArr, tempLayer)
        Next k
        globalData.text = infoArr
        globalData.state = "True"
        If cmdCommand = "get:text" Then
            globalData.steps = "获取文本信息完成"
        Else
            globalData.steps = "设置文本信息完成"
        End If
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
                'Console.WriteLine(names(i))
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

            execMain(app, doc)

        Catch ex As Exception
            ' log("error", "CorelDRAW执行功能错误")
            Exit Sub
        End Try

    End Sub


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
            ElseIf count = 3 Then
                cmdPath = args(2)
            End If
        End If
        globalData.steps = "解析参数完成"
    End Function

    Sub Main()

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

        Console.OutputEncoding = Encoding.UTF8
        Console.WriteLine(JsonConvert.SerializeObject(globalData))

        'MsgBox(1)
    End Sub

End Module

