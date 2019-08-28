Imports Corel.Interop.VGCore


Module Module1


    '获取文档所有页面、所有图层、所有图形对象
    Public Function getFiles() As String
        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)
        'app.Visible = True
        Dim doc As Document = app.ActiveDocument

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
                    ' 根据形状的类型，输出不同的信息
                    msg = "在页面" & i & "的图层" & j & "中，找到了一个："
                    Dim cdrTextShape As cdrShapeType = Nothing

                    Debug.Print(tempShape.SizeWidth)

                    ' 如果是文本形状
                    If tempShape.Type = cdrTextShape Then
                        msg = msg & "文本"
                    End If
                    Dim cdrBitmapShape As cdrShapeType = Nothing
                    ' 如果是位图形状
                    If tempShape.Type = cdrBitmapShape Then
                        msg = msg & "位图"
                    End If
                    ' 打印调试消息到本地调试窗口
                    Debug.Print(msg)
                Next k
            Next j
        Next i
        MsgBox("遍历文档完成！请查看调试窗口")


    End Function


    '====================================================================================================================================================================
    '@desc: 在一组形状中找出尺寸最大的一个图形
    '@author: Zebe
    '@url: http://www.cdrvba.com
    '@param sh: 图形对象集合
    '@return: 返回一组形状中尺寸最大的一个图形
    '====================================================================================================================================================================
    Public Function getMaxSizeShapeInShapes(sh As Collection) As Shape
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
        End If
 
        Console.WriteLine(resultShape)
    End Function

    Sub createDocumnet()
        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)
        'app.Visible = True
        Dim doc As Document = app.ActiveDocument

        If doc.Pages.Count = 0 Then
            MsgBox("There aren't any open documents")
            Exit Sub
        End If
        
        
        
        'Console.WriteLine(doc.Unit )

        'doc.AddPages(1)
        'doc.ActivePage.Color.RGBAssign(255,255,0)
        'Console.WriteLine(app.Application )
       ' Debug.print(doc)
        'getMaxSizeShapeInShapes(doc.Pages)
        'For Each Item In doc.ActiveLayer.Shapes
       '     Console.WriteLine(Item)
       ' Next

        MsgBox(123)

        'getFiles()


      


        'Console.WriteLine(allPages)



        'doc.ActivePage.Color.RGBAssign(255, 0, 0)


        'Dim doc As Document = app.CreateDocument()
        'doc.ActiveLayer.CreateEllipse2(3, 2, 1)
        'doc.ActiveDocument.Selection.Shapes(1).Fill.UniformColor.CMYKAssign(155, 222, 333, 0)

        'doc = app.CreateDocument()
        'doc.ActiveLayer.CreateEllipse2(3, 4, 5)

        'doc = app.CreateDocument()
        'doc.ActiveLayer.CreateEllipse2(5, 6, 7)

        ' app.Documents(1).Activate()
        'app.ActiveDocument.AddPages(3)


        ' MsgBox(app.AppWindow().Left)




        Console.WriteLine("cateate new document")


        ' Dim shape As Shape = doc.ActiveLayer.CreateArtisticText(
        '   0.0, 0.0, text, cdrTextLanguage.cdrLanguageMixed,
        '    cdrTextCharSet.cdrCharSetMixed, fontName, fontSize,
        '   cdrTriState.cdrUndefined, cdrTriState.cdrUndefined,
        '    cdrFontLine.cdrMixedFontLine, cdrAlignment.cdrLeftAlignment)
        Console.WriteLine("complete")
    End Sub







    Sub CreateTextInCorelDRAW(text As String, fontName As String, fontSize As Single)
        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)
        app.Visible = True
        Dim doc As Document = app.ActiveDocument
        Console.WriteLine("activate document")
        If doc Is Nothing Then
            doc = app.CreateDocument()
            Console.WriteLine("cateate new document")
        End If

        Dim sPath As Shape = doc.ActiveLayer.CreateEllipse(0, 10, 60, 60)
        Dim sh As Shape = doc.ActiveLayer.CreateArtisticText(1, 4, "这是沿着形状路径排列的美术字")
        sh.Text.FitToPath(sPath)


        ' Dim shape As Shape = doc.ActiveLayer.CreateArtisticText(
        '   0.0, 0.0, text, cdrTextLanguage.cdrLanguageMixed,
        '    cdrTextCharSet.cdrCharSetMixed, fontName, fontSize,
        '   cdrTriState.cdrUndefined, cdrTriState.cdrUndefined,
        '    cdrFontLine.cdrMixedFontLine, cdrAlignment.cdrLeftAlignment)
        Console.WriteLine("complete")
    End Sub




    '==================================================================================
    '文件夹选择函数
    '@描述：调用文件夹选择对话框来获得选择的文件夹路径，如果没有选择则返回空字符串
    '==================================================================================
    Public Function chooseFolder() As String
        ' 创建Shell对象，用来浏览系统文件夹
        Dim shell = CreateObject("Shell.Application")
        Dim folder = shell.BrowseForFolder(0, "选择文件夹", 0, 0)
        ' 判断是否选择了文件夹
        If Not folder Is Nothing Then
            chooseFolder = folder.self.path ' 返回文件夹自身的路径
        Else
            chooseFolder = ""
        End If
        ' 释放内存
        folder = Nothing
        shell = Nothing
    End Function


    '========================================================================================
    ' 创建图形（这个方法中没有用到 title 参数，可根据需要使用，例如设置备注）
    ' @author: Zebe
    ' @date: 2017/12/11
    '========================================================================================
    Private Sub createShape(width As Integer, height As Integer, title As String)
        Dim pia_type As Type = Type.GetTypeFromProgID("CorelDRAW.Application")
        Dim app As Application = Activator.CreateInstance(pia_type)
        Dim doc As Document = app.ActiveDocument
        ' 如果没有活动文档，则自动创建一个文档，并设置文档单位为mm
        If doc Is Nothing Then
            doc.CreateDocument()
            Dim cdrMillimeter As cdrUnit = Nothing
            doc.Unit = cdrMillimeter
        End If
        ' 在页面左下角（坐标0,0）开始，创建指定宽高的矩形
        doc.ActivePage.ActiveLayer.CreateRectangle2(0, 0, width, height)
    End Sub

    '========================================================================================
    ' 读取XML文件并创建图形
    '========================================================================================
    Private Sub readXMLAndCreateShape(filePath As String)
        ' 载入XML文件
        Dim xmlDom = CreateObject("MSXML.DOMDocument")
        xmlDom.Load(filePath)
        xmlDom.async = False ' 关闭异步读取，设置为同步读取（即：这句代码会阻塞，直到文件读取完）

        ' 节点变量声明

        Dim shapeNodes = xmlDom.SelectNodes("//shape")
        Dim widthNodes = xmlDom.SelectSingleNode("//width")
        Dim heightNodes = xmlDom.SelectSingleNode("//height")
        Dim titleNodes = xmlDom.SelectSingleNode("//title")

        ' 遍历所有shape节点
        Dim i As Integer, j As Integer
        For i = 0 To shapeNodes.Length - 1
            ' 取出每个shape节点下的子节点（根据索引序号去取）
            Dim width As Integer, height As Integer, title As String
            width = shapeNodes.Item(i).ChildNodes(0).Text
            height = shapeNodes.Item(i).ChildNodes(1).Text
            title = shapeNodes.Item(i).ChildNodes(2).Text
            'MsgBox ("width=" & width & ", height=" & height & ", title=" & title)
            createShape(width, height, title)
        Next i

        ' 释放已经加载的DOM对象所占用的内存
        xmlDom = Nothing
    End Sub


    '*********************************************************************************************
    ' 文档回调事件
    ' 每当新的文档被创建时，此过程被触发
    ' @author Zebe
    '*********************************************************************************************
    Private Sub GlobalMacroStorage_DocumentNew(ByVal doc As Document, ByVal FromTemplate As Boolean,
                                               ByVal Template As String, ByVal IncludeGraphics As Boolean)
        MsgBox("检测到新文档被创建，文档名称：" & doc.Title)
    End Sub





    Sub Main()
        Try

            createDocumnet()
            'readXMLAndCreateShape("C:\Users\Administrator\source\repos\app\app\shape.xml")' 调用过程（请注意XML文件路径）


            'CreateTextInCorelDRAW("激活文档写数据1", "Arial", 24.0F)
            ' invisibleMethod()

            '  Dim path As String = chooseFolder '调用文件夹选择函数，将其返回值传递给 path 变量
            '  If path <> "" Then
            '      MsgBox("你选择的是：" & path)
            '    Else
            '      MsgBox("你什么都没有选择")
            '  End If
        Catch ex As Exception
            Console.WriteLine("open error")
        End Try
    End Sub


End Module