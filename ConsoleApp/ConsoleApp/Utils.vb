Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module Utils



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
        Return e
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




End Module
