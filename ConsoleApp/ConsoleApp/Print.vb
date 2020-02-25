Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module Print

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

        globalData.steps = "开始连接CorelDRAW"
        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)

        If Len(Param.cmdPath) > 2 Then
            globalData.steps = "开始打开文档"
            app.OpenDocument(Param.cmdPath)
        End If

        Dim doc As Document = app.ActiveDocument
        If app.Documents.Count = 0 Then
            globalData.errorlog = "没有找到活动文档"
            Exit Sub
        End If



        Console.WriteLine(doc)


        MsgBox(1)

    End Sub

End Module

