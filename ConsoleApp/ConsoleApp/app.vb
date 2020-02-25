Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Module App

    Dim lineCount = 2


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
            Console.WriteLine(1)
            active_doc.PrintSettings.Collate = value
        End If

        If key = "FileName" Then
            Console.WriteLine(2)
            active_doc.PrintSettings.FileName = value
        End If

        If key = "Copies" Then
            Console.WriteLine(3)
            active_doc.PrintSettings.Copies = value
        End If

        If key = "PrintRange" Then
            Console.WriteLine(4)
            active_doc.PrintSettings.PrintRange = value
        End If

        If key = "PageRange" Then
            Console.WriteLine(6)
            active_doc.PrintSettings.PageRange = value
        End If

        If key = "ShowDialog" Then
            Console.WriteLine(5)
            active_doc.PrintSettings.ShowDialog()
        End If

        If key = "PageSet" Then
            Console.WriteLine(7)
            active_doc.PrintSettings.PageSet = value
        End If

        If key = "PaperOrientation" Then
            Console.WriteLine(8)
            active_doc.PrintSettings.PaperOrientation = value
        End If

        If key = "PrintToFile" Then
            Console.WriteLine(9)
            active_doc.PrintSettings.PrintToFile = True
        End If

        If key = "SelectPrinter" Then
            Console.WriteLine(10)
            active_doc.PrintSettings.SelectPrinter(value)
        End If

        If key = "PaperSize" Then
            Console.WriteLine(1)
            active_doc.PrintSettings.PaperSize = value
        End If

    End Function


    '设置打印数组参数
    Function setPrintArrayValue(doc, key, value)
        Dim active_doc As Document = doc
        If key = "SetPaperSize" Then
            Dim v As JArray = value
            active_doc.PrintSettings.SetPaperSize(v.First.ToString(), v.Last.ToString())
        End If

        ' Console.WriteLine(key)
        'active_doc.PrintSettings.Printer.ShowDialog()

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

            '如果参数是0
            If value.ToString() = "0" Then
                setPrintValue(active_doc, key, value.ToString())
            Else
                '如果参数是空
                If value.ToString().Length = 0 Then
                    Return False
                Else
                    If value = False Then
                        Return False
                    Else
                        setPrintValue(active_doc, key, value.ToString())
                    End If
                End If
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


            '独立命令，打印
            If Param.cmdCommand = "print" Then
                Dim settingsObject As JObject = cmdPrintSettings
                For Each item In settingsObject
                    setPrint(doc, item.Key, item.Value)
                Next
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

