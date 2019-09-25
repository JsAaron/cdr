Imports Corel.Interop.VGCore
Imports System.IO
Imports System.Text

Module Module1

    Dim lineCount = 2
    Dim cmdPath
    Dim cmdFontJson = "True"
    Dim mainCount = 2


    Public Function log(name, value)
        Console.WriteLine("{""" + name + """:""" + value + """}")
        Return False
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
                Console.WriteLine(tempLayer.Name)
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


    Sub outputResult(app, doc)

        '指定毫米
        doc.Unit = 3
        Dim width = app.ActivePage.SizeWidth
        Dim height = app.ActivePage.SizeHeight
        Dim tW = """width"""
        Dim tH = """height"""
        Dim data = "{" & tW & ":""" & width & """," & tH & ":""" & height & """}"
        '单独输出结构
        'Console.Write("{""pagesize"":" + data + "}|")

        '创建当前文字的json
        If cmdFontJson = "True" Then
            log("log", "开始搜索文档字体")
            '获取当前的应用的字体
            Dim names = getFontNames(doc)
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
                    Console.WriteLine(names(i))
                    fontList.Add(names(i))
                    Dim empty As String = ""
                    str = str + "{" + """fontname""" + ":""" + empty + """," + """familyname""" + ":""" + Replace(names(i), Chr(10), "") + """," + """postscriptname""" + ":""" + empty + """},"
                End If
            Next

            log("log", "开始处理字体")

            '去掉最后一个，
            str = Left(str, Len(str) - 1)
            str = "[" + str + "]"
            Dim p() = Split(doc.FileName, ".")
            log("log", "开始处理字体FileName")
            Dim fs As FileStream = File.Create(doc.FilePath + p(0) + ".json")
            log("log", "开始处理字体FileStream")
            Dim info As Byte() = New UTF8Encoding(True).GetBytes(str)
            log("log", "开始处理字体UTF8Encoding")
            fs.Write(info, 0, info.Length)
            log("log", "开始处理字体Write")
            fs.Close()
            log("fontjson", "true")

        End If

    End Sub



    Function jsonDecode(jsonString)
        Dim L = Len(jsonString)
        Dim str = Mid(jsonString, 2, L - 2)
        Dim args() = Split(str, ",")

        Dim reg As Object = CreateObject("VBScript.Regexp")
        reg.Global = True
        reg.Pattern = """(.*?)"": ""(.*?)"""

        For i = 0 To UBound(args)
            Dim matches = reg.Execute(args(i))
            '遍历所有匹配到的结果'    
            For Each match In matches
                If match.SubMatches(0) = "path" Then
                    cmdPath = match.SubMatches(1)
                End If
                If match.SubMatches(0) = "fontJson" Then
                    cmdFontJson = match.SubMatches(1)
                End If
            Next
        Next i

    End Function

    Function getJob(jsonString As String)
        Dim strFunc, objSC, objJSON
        objSC = CreateObject("ScriptControl")
        objSC.Language = "JScript"
        strFunc = "function jsonParse(s) { return eval('(' + s + ')'); }"
        objSC.AddCode(strFunc)
        objJSON = objSC.CodeObject.jsonParse(jsonString)
        getJob = jsonDecode(objJSON)
    End Function


    Sub checkLine(app)

        Try
            log("log", "CorelDRAW开始连接文档")

            If Len(cmdPath) > 2 Then
                app.OpenDocument(cmdPath)
            End If

            Dim doc As Document = app.ActiveDocument

            '如果没有文档
            If app.Documents.Count = 0 Then
                log("error", "没有找到活动文档")
                Exit Sub
            End If

        Catch ex As Exception
            'log("log", "CorelDRAW打开文档失败，休眠3秒后重新链接")
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

            '如果没有文档
            If app.Documents.Count = 0 Then
                log("error", "没有找到活动文档")
                Exit Sub
            End If

            log("log", "文档连接成功")
            outputResult(app, doc)
        Catch ex As Exception
            log("error", "CorelDRAW执行功能错误")
            Exit Sub
        End Try

    End Sub


    Sub Main()

        '如果有外部命令
        If Len(Command) > 0 Then
            getJob(Command)
            If Len(cmdPath) <= 2 Then
                log("error", "文档路径为空")
                Exit Sub
            End If
            If cmdFontJson = "True" Then
                log("log", "启用了字体采集功能")
            End If
        End If

        log("log", "CorelDRAW开始链接")
        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)
        app.Visible = True
        log("log", "CorelDRAW链接成功")

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

