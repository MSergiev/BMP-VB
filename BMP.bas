Attribute VB_Name = "BMP"
'-------------------------------------------
'---------- VBA EXCEL BMP PARSER -----------
'----------- (24-bit BMP Files) ------------
'-------------------------------------------
'-------------------------------------------
'- Created by: Miroslav Sergiev, MI, 13568 -
'---------------- FMI 2015 -----------------
'-------------------------------------------


Private sizeArr(15) As Byte     'Array with header chunk sizes
Private signature As String     'File signature
Private valArr(14) As Long      'Array with header data
Private pixArr() As Long        'Array with pixel color data

'Main subroutine
    Sub BMP()
        'Stop screen updating
        UpdateScreen = Application.ScreenUpdating
        Application.ScreenUpdating = False
        
        SizeArrInit     'Initialize size array
        Reset           'Reset worksheet
        BMPOpen         'Open and read BMP file
        GetData         'Draw header data
        BMPDraw         'Draw pixel data
            
        'Resume screen updating
        Application.ScreenUpdating = UpdateScreen
    End Sub

'Open and read BMP file
    Private Sub BMPOpen()
        Sheet1.Name = "Data"
        Worksheets("Data").Activate
        Dim bytTemp As Byte, tmpArr(3) As Byte, padding As Byte
        InitArray (tmpArr)
        Sheet1.[A1] = Application.GetOpenFilename(filefilter:="Bitmap Files (*.bmp), *.bmp")
        Open Sheet1.[A1] For Binary Access Read As #1
            
        Get 1, , tmpArr(0)
        Get 1, , tmpArr(1)
        signature = StrConv(tmpArr, vbUnicode)
        InitArray (tmpArr)
        
        For i = 0 To UBound(valArr)
            For j = 0 To sizeArr(i + 1) - 1
                Get 1, , tmpArr(j)
            Next
            valArr(i) = ByteConv(tmpArr)
            InitArray (tmpArr)
        Next
        
        ReDim pixArr(valArr(6) - 1, valArr(5) - 1)
        padding = (4 - (valArr(5) * 3 Mod 4)) Mod 4
        
        For i = 0 To valArr(6) - 1 Step 1
            For j = 0 To valArr(5) - 1 Step 1
                For k = 0 To 2 Step 1
                    Get 1, , tmpArr(k)
                Next
                pixArr(i, j) = CellColor(tmpArr)
                InitArray (tmpArr)
            Next
            For k = 1 To padding
                Get 1, , bytTemp
            Next
        Next
        
        Close 1
    End Sub

'Draw pixel data
    Private Sub BMPDraw()
        Sheet2.Name = "Image"
        Worksheets("Image").Activate
        With Range(Cells(1, 1), Cells(valArr(6), valArr(5)))
            .ColumnWidth = 0.1
            .RowHeight = .Cells(1).Width
        
            For i = 1 To valArr(6)
                For j = 1 To valArr(5)
                    .Cells(i, j).Interior.Color = pixArr(valArr(6) - 1 - (i - 1), (j - 1))
                Next
            Next
        End With
    End Sub

'Draw header data
    Private Sub GetData()
        Sheet1.Activate
        Range("A1:C1").Cells.Merge
        Range("A2:C2").Cells.Merge
        Range("A8:C8").Cells.Merge
        With Range("A1:A19")
            .ColumnWidth = 15
            .Columns.HorizontalAlignment = xlCenter
            .Rows(2) = "Bitmap File Header"
            .Rows(3) = "Signature"
            .Rows(4) = "Size"
            .Rows(5) = "Reserved 1"
            .Rows(6) = "Reserved 2"
            .Rows(7) = "Offset"
            .Rows(8) = "DIB Header"
            .Rows(9) = "Header Size"
            .Rows(10) = "Image Width"
            .Rows(11) = "Image Height"
            .Rows(12) = "Color Planes"
            .Rows(13) = "Bits Per Pixel"
            .Rows(14) = "Compression"
            .Rows(15) = "Image Size"
            .Rows(16) = "H.Resolution"
            .Rows(17) = "V.Resolution"
            .Rows(18) = "Pallette Colors"
            .Rows(19) = "Important Colors"
        End With
        With Range("B1:B19")
            .ColumnWidth = 2
            .Columns.HorizontalAlignment = xlCenter
            .Rows(3) = sizeArr(0)
            .Rows(4) = sizeArr(1)
            .Rows(5) = sizeArr(2)
            .Rows(6) = sizeArr(3)
            .Rows(7) = sizeArr(4)
            .Rows(9) = sizeArr(5)
            .Rows(10) = sizeArr(6)
            .Rows(11) = sizeArr(7)
            .Rows(12) = sizeArr(8)
            .Rows(13) = sizeArr(9)
            .Rows(14) = sizeArr(10)
            .Rows(15) = sizeArr(11)
            .Rows(16) = sizeArr(12)
            .Rows(17) = sizeArr(13)
            .Rows(18) = sizeArr(14)
            .Rows(19) = sizeArr(15)
        End With
        With Range("C1:C19")
            .ColumnWidth = 7
            .Columns.HorizontalAlignment = xlCenter
            .Rows(3) = signature
            .Rows(4) = valArr(0)
            .Rows(5) = valArr(1)
            .Rows(6) = valArr(2)
            .Rows(7) = valArr(3)
            .Rows(9) = valArr(4)
            .Rows(10) = valArr(5)
            .Rows(11) = valArr(6)
            .Rows(12) = valArr(7)
            .Rows(13) = valArr(8)
            .Rows(14) = valArr(9)
            .Rows(15) = valArr(10)
            .Rows(16) = valArr(11)
            .Rows(17) = valArr(12)
            .Rows(18) = valArr(13)
            .Rows(19) = valArr(14)
        End With
        
        With Range("A2:C2")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Cells.Interior.Color = RGB(0, 51, 0)
        End With
        With Range("A3:C7")
            .Cells.Interior.Color = RGB(204, 255, 204)
        End With
        With Range("A8:C8")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Cells.Interior.Color = RGB(0, 0, 128)
        End With
        With Range("A9:C19")
            .Cells.Interior.Color = RGB(153, 204, 255)
        End With
        Range("B1:B19").Cells.BorderAround (xlContinuous)
            
        Range("A8").Font.Bold = True
        
        Range("A1:C19").Borders.LineStyle = xlContinuous
    End Sub

'Reset worksheet
    Private Sub Reset()
        Sheet1.Activate
        Cells.Clear
        Cells.ClearFormats
        Rows.RowHeight = Rows(Rows.Count).RowHeight
        Columns.ColumnWidth = Columns(Columns.Count).ColumnWidth
            
        Sheet2.Activate
        Cells.Clear
        Cells.ClearFormats
        Rows.RowHeight = Rows(Rows.Count).RowHeight
        Columns.ColumnWidth = Columns(Columns.Count).ColumnWidth
    End Sub

'Append byte array into a Long
    Private Function ByteConv(Arr) As Long
        For i = UBound(Arr) To 0 Step -1
            ByteConv = ByteConv + Arr(i) * 2 ^ (i * 8)
        Next
    End Function

'Convert RGB triplet to RGB Long
    Private Function CellColor(Arr) As Long
        Dim red As Integer, green As Integer, blue As Integer
        For i = 0 To UBound(Arr)
            red = Arr(2)
            green = Arr(1)
            blue = Arr(0)
            CellColor = RGB(red, green, blue)
        Next
    End Function

'Initialize size array
    Private Function SizeArrInit()
        sizeArr(0) = 2
        sizeArr(1) = 4
        sizeArr(2) = 2
        sizeArr(3) = 2
        sizeArr(4) = 4
        sizeArr(5) = 4
        sizeArr(6) = 4
        sizeArr(7) = 4
        sizeArr(8) = 2
        sizeArr(9) = 2
        sizeArr(10) = 4
        sizeArr(11) = 4
        sizeArr(12) = 4
        sizeArr(13) = 4
        sizeArr(14) = 4
        sizeArr(15) = 4
    End Function

'Clear array
    Private Function InitArray(ByRef Arr)
        For i = 0 To UBound(Arr)
            Arr(i) = 0
        Next
    End Function
