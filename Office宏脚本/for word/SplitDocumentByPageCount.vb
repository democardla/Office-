Sub SplitDocumentByPageCount()
    Dim doc As Document
    Dim newDoc As Document
    Dim pageCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim startPage As Integer
    Dim endPage As Integer
    Dim totalPages As Integer
    Dim pageRange As Range
    Dim folderPath As String

    ' 这里填写你要储存的位置
    folderPath = "/Users/democardla/Desktop/测试用数据材料/"


    ' 这里表示为`每隔pageCount页`将页面储存为新的文档
    pageCount = 10 ' 替换为实际的每个文档包含的页数
    
    ' 获取当前打开的文档
    Set doc = ActiveDocument
    
    ' 计算文档的总页数
    totalPages = doc.ComputeStatistics(wdStatisticPages)
    
    ' 计算分割后的文件数量
    numFiles = Int(totalPages / pageCount)
    If totalPages Mod pageCount <> 0 Then
        numFiles = numFiles + 1
    End If
    
    ' 循环创建新文件并复制页面
    For i = 1 To numFiles
        ' 计算本次复制的起始页和结束页
        startPage = (i - 1) * pageCount + 1
        endPage = i * pageCount
        If endPage > totalPages Then
            endPage = totalPages
        End If
        
        Set pageRange = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=startPage)
        pageRange.End = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=endPage + 1).End

        ' 创建新的Word文档
        Set newDoc = Documents.Add

        ' 将原文档激活，以便复制内容
        doc.Activate

        pageRange.Copy
        ' 修改为您想要粘贴的内容类型，比如 wdPasteFormat 用于保留格式
        newDoc.Range.PasteSpecial DataType:=wdPasteRTF 
        
        ' 保存新文档，文件名为`folderPath`+`SplitFile`+`i`+`.docx`
        newDoc.SaveAs2 FileName:=folderPath & "SplitFile" & i & ".docx"

        newDoc.Close
    Next i
    
    ' 提示分割完成
    MsgBox "文档已成功分割为" & numFiles & "个文件。"
End Sub


