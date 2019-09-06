Imports Corel.Interop.VGCore
Imports System.IO
Imports System.Text

Module Module1

  Public Function log(name, value)
    Console.WriteLine("{""" & name & """:""" & value & """}")
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

  Sub createDocumnet()
    Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
    Dim app As Application = Activator.CreateInstance(pia_type)

    app.Visible = True

    '启动休眠
    Threading.Thread.Sleep(3000)

    Dim b() = Split(Command, " ")
    Dim path As String
    Dim createJson As Boolean = False
    If b.Count = 1 Then
      path = b(0)
    End If
    If b.Count = 2 Then
      path = b(0)
      If b(1) = "fontJson:true" Then
        createJson = True
      End If
    End If

    '如果有命令路径参数，打开对应的cdr
    If path <> "" Then
      app.OpenDocument(path)
    End If

    Dim doc As Document = app.ActiveDocument

    '如果没有文档
    If app.Documents.Count = 0 Then
      log("error", "0002")
      Exit Sub
    End If


    '指定毫米
    doc.Unit = 3
    Dim width = app.ActivePage.SizeWidth
    Dim height = app.ActivePage.SizeHeight
    Dim tW = """width"""
    Dim tH = """height"""
    Dim data = "{" & tW & ":""" & width & """," & tH & ":""" & height & """}"

    '创建当前文字的json
    If createJson = True Then

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
          fontList.Add(names(i))
          str = str + "{" + """fontname""" + ":""" + names(i) + """," + """familyname""" + ":""" + names(i) + """," + """postscriptname""" + ":""" + names(i) + """},"
        End If
      Next

      '去掉最后一个，
      str = Left(str, Len(str) - 1)
      str = "[" + str + "]"
      Dim fs As FileStream = File.Create(doc.FilePath + "font.json")
      Dim info As Byte() = New UTF8Encoding(True).GetBytes(str)
      fs.Write(info, 0, info.Length)
      fs.Close()
      '文件写入完成
      log("finish", "0001")
    End If

    '输出参数
    Console.Write(data)

  End Sub


  Sub Main()
    Try
      createDocumnet()
    Catch ex As Exception
      log("error", "0001")
    End Try
  End Sub

End Module

