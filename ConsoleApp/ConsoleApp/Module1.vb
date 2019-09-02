﻿Imports Corel.Interop.VGCore


Module Module1


  Sub createDocumnet()
    Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
    Dim app As Application = Activator.CreateInstance(pia_type)

    '如果有命令路径参数，打开对应的cdr
    If Command() <> "" Then
      app.OpenDocument(Command)
    End If

    'app.Visible = True
    Dim doc As Document = app.ActiveDocument

    '如果没有文档
    If app.Documents.Count = 0 Then
      Console.Write("error:0002")
      Exit Sub
    End If


    '指定毫米
    doc.Unit = 3

    Dim width = app.ActivePage.SizeWidth
    Dim height = app.ActivePage.SizeHeight

    Console.Write("{" & "width:" & width & "," & "height:" & height & "}")

  End Sub


  Sub Main()
    Try
      createDocumnet()
    Catch ex As Exception
      Console.Write("error:0001")
    End Try
  End Sub

End Module

