Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json


Module Module1

    Dim lineCount = 2
    Dim mainCount = 2

    Dim cmdPath As String
    Dim cmdConfig As Object
    Dim cmdExternalData As Object

    Dim Test = True

    '////////////////////////////////////// 功能 //////////////////////////////////////////////////

    Public Function Debug(value)
        Console.WriteLine("debug: " & value)
    End Function


    Public Function log(name, value)
        Console.WriteLine("{""" + name + """:""" + value + """}")
        Return False
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


    '////////////////////////////////////// 逻辑 //////////////////////////////////////////////////

    '递归检测形状
    Public Function recurveShape(allShapes, infoArr)
        Dim tempShape As Shape
        For k = 1 To allShapes.Count
            ' 得到这个形状
            tempShape = allShapes.Item(k)
            Dim cdrTextShape As cdrShapeType = 6
            Dim cdrGroupShape As cdrShapeType = 7

            '组
            If tempShape.Type = cdrGroupShape Then
                recurveShape(tempShape.Shapes, infoArr)
            End If

            '文字
            If tempShape.Type = cdrTextShape Then
                '如果有值
                If tempShape.Text.Story.Text <> "" Then
                    Dim t As New ArrayList
                    t.Add(tempShape.Name)
                    t.Add(tempShape.Text.Story.Text)
                    infoArr.Add(t)
                End If
            End If
        Next k
    End Function




    '获取文档所有页面、所有图层、所有图形对象
    Public Function getExtractData(doc)
        Dim infoArr As New ArrayList
        Dim i As Integer, k As Integer
        Dim tempLayer, allLayers

        allLayers = doc.ActivePage.AllLayers
        For k = 1 To allLayers.Count
            tempLayer = allLayers.Item(k)
            recurveShape(tempLayer.Shapes, infoArr)
        Next k

        If infoArr.Count > 0 Then
            Dim str = "{""extract"":["
            For i = 0 To infoArr.Count
                str = str & "{" & """" & (infoArr.Item(i).item(0)) & """" & ":" & """" & (infoArr.Item(i).item(1)) & """" & "}"
                If i = (infoArr.Count - 1) Then
                    str = str + "]}|"
                    Console.Write(str)
                Else
                    str = str + ","
                End If
            Next i
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
        Dim msg As String
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
        'log("log", "开始搜索文档字体")
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
        Console.Write("{""fontjson"":""True""}|")
    End Function


    '主体方法
    Sub execMain(app, doc)
        Dim setPagesize, setFontjson, setExtract

        If cmdConfig <> "" Then
            setPagesize = cmdConfig.pagesize
            setFontjson = cmdConfig.fontjson
            setExtract = cmdConfig.extract
        End If

        If Test = True Then
            setPagesize = "True"
            setFontjson = "True"
            setExtract = "True"
        End If

        '页面尺寸
        If setPagesize = "True" Then
            '指定毫米
            doc.Unit = 3
            Dim width = app.ActivePage.SizeWidth
            Dim height = app.ActivePage.SizeHeight
            Dim tW = """width"""
            Dim tH = """height"""
            Dim data = "{" & tW & ":""" & width & """," & tH & ":""" & height & """}"
            '单独输出结构
            Console.Write("{""pagesize"":" + data + "}|")
        End If

        '创建当前文字的json
        If setFontjson = "True" Then
            createFontJson(doc)
        End If


        '获取数据
        If setExtract = "True" Then
            getExtractData(doc)
        End If

    End Sub


    Sub checkLine(app)
        Try
            If Len(cmdPath) > 2 Then
                app.OpenDocument(cmdPath)
            End If
            Dim doc As Document = app.ActiveDocument
            If app.Documents.Count = 0 Then
                log("error", "没有找到活动文档")
                Exit Sub
            End If
        Catch ex As Exception
            If lineCount = 0 Then
                log("error", "CorelDRAW打开文档错误")
                Exit Sub
            End If
            log("error", "CorelDRAW文档打开失败,尝试重新打开文档")
            lineCount = lineCount - 1
            Threading.Thread.Sleep(3000)
            checkLine(app)
            Exit Sub
        End Try

        Try
            Dim doc As Document = app.ActiveDocument
            If app.Documents.Count = 0 Then
                log("error", "没有找到活动文档")
                Exit Sub
            End If
            Debug("执行主方法操作")
            execMain(app, doc)
        Catch ex As Exception
            log("error", "CorelDRAW执行功能错误")
            Exit Sub
        End Try

    End Sub


    Public Function parseCommand()
        Dim args() = Split(Command, " ")
        If args.Count < 3 Then
            log("error", "外部传入参数解析错误")
        End If
        Debug("参数解析开始")
        cmdPath = args(0)
        cmdConfig = parseJson(args(1))
        cmdExternalData = parseJson(args(2))
        Debug("参数解析完毕")
    End Function

    Sub Main()

        Debug("外部参数 - " & Command())

        '如果有外部命令
        If Len(Command) > 0 Then
            parseCommand()
            If Len(cmdPath) <= 2 Then
                log("error", "文档路径为空")
                Exit Sub
            End If
        End If

        Debug("CorelDRAW开始链接")

        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)
        app.Visible = True

        Debug("CorelDRAW链接成功")

        Try
            checkLine(app)
        Catch ex As Exception
            If mainCount = 0 Then
                log("error", "CorelDRAW软件无法链接")
                Exit Sub
            End If
            log("log", "CorelDRAW链接失败,开始下一次链接")
            mainCount = mainCount - 1
            Threading.Thread.Sleep(3000)
            checkLine(app)
            Exit Sub
        End Try

        MsgBox("结束")
    End Sub

End Module

