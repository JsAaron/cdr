Imports Corel.Interop.VGCore

Module Module1

  '====================================================================================================================================================================
  '@desc: 在一组形状中找出尺寸最大的一个图形
  '@return: 返回一组形状中尺寸最大的一个图形
  '====================================================================================================================================================================
  Public Function getMaxSizeShapeInShapes(sh As Shapes) As Shape
    Dim resultShape As Shape
    Dim i As Integer
    Dim tempShape As Shape
    If sh.Count > 0 Then
      For i = 1 To sh.Count
        tempShape = sh.Item(i)
        If i = 1 Then
          resultShape = tempShape
        Else
          If tempShape.SizeWidth > resultShape.SizeWidth And tempShape.SizeHeight > resultShape.SizeHeight Then
            resultShape = tempShape
          End If
        End If
      Next i
      getMaxSizeShapeInShapes = resultShape
    End If

  End Function

  Sub createDocumnet()
    Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
    Dim app As Application = Activator.CreateInstance(pia_type)
    'app.Visible = True
    Dim doc As Document = app.ActiveDocument

    '如果没有文档
    If app.Documents.Count = 0 Then
      MsgBox("There aren't any open documents")
      Exit Sub
    End If

    '指定毫米
    doc.Unit = 3
    Dim pageWidth = app.PageSizes(1).Width
    Dim pageHeight = app.PageSizes(1).Height
    'doc.ActivePage.CreateLayer("MY A MU")
    ' Console.WriteLine(doc.ActiveLayer.Shapes.Count)

    Dim resultShape = getMaxSizeShapeInShapes(doc.ActiveLayer.Shapes)
    Console.WriteLine(resultShape.SizeWidth)
    Console.WriteLine(resultShape.SizeHeight)
    MsgBox(1)

  End Sub


  Sub Main()
    Try
      createDocumnet()
    Catch ex As Exception
      Console.WriteLine("open error")
    End Try
  End Sub

End Module

