Sub pecahString()
    Dim HighWordArray() As String
    Dim MedWordArray() As String
    Dim LowWordArray() As String
    
    Dim textStringHigh As String
    Dim textStringMed As String
    Dim textStringLow As String
    Dim categoryBarang As String
    
    Dim buatArray As Long
    Dim baris As Long
    Dim barisLanjut As Long
    Dim batesBawah As Long
    Dim barisLoop As Long
    Dim kolomLoop As Long
    
    'For i = LBound(WordArray) To UBound(WordArray)
    '    strr = strr & vbNewLine & "Isi " & i & " - " & Trim(WordArray(i))
    'Next i
    
    'MsgBox strr
    
    Sheets("Data").Cells.Clear
    
    Sheets("Data").Cells(2, 1).Value = "Category"
    Sheets("Data").Cells(2, 2).Value = "Priority"
    Sheets("Data").Cells(2, 3).Value = "Brand"
    
    barisLoop = 1
    kolomLoop = 1
    batesBawah = 0
    
    Do Until IsEmpty(Cells(barisLoop, 1).Value) And IsEmpty(Cells(barisLoop, 1 + 1).Value)
    
        categoryBarang = Cells(barisLoop, 1).Value
    
        textStringHigh = Cells(barisLoop, 2).Value
        HighWordArray() = Split(textStringHigh, ",")
    
        textStringMed = Cells(barisLoop, 3).Value
        MedWordArray() = Split(textStringMed, ",")
    
        textStringLow = Cells(barisLoop, 4).Value
        LowWordArray() = Split(textStringLow, ",")
        
        If batesBawah = 0 Then
            For j = 3 To (UBound(HighWordArray) + 3)
        
                buatArray = j - 3
                Sheets("Data").Cells(j, 1).Value = categoryBarang
                Sheets("Data").Cells(j, 2).Value = "High"
                Sheets("Data").Cells(j, 3).Value = Trim(HighWordArray(buatArray))
                baris = j
            Next j
            
        Else
            
            buatArray = 0
            For j = batesBawah To (UBound(HighWordArray) + batesBawah)
        
                Sheets("Data").Cells(j, 1).Value = categoryBarang
                Sheets("Data").Cells(j, 2).Value = "High"
                Sheets("Data").Cells(j, 3).Value = Trim(HighWordArray(buatArray))
                baris = j
                
                If buatArray < UBound(HighWordArray) Then
                    buatArray = buatArray + 1
                End If
            Next j
        End If
        
        
    
        baris = baris + 1
    
        batesBawah = UBound(MedWordArray) + baris
        buatArray = 0
        For j = baris To (batesBawah)
            Sheets("Data").Cells(j, 1).Value = categoryBarang
            Sheets("Data").Cells(j, 2).Value = "Medium"
            Sheets("Data").Cells(j, 3).Value = Trim(MedWordArray(buatArray))
            barisLanjut = baris
        
            If buatArray < UBound(MedWordArray) Then
                buatArray = buatArray + 1
            End If
        Next j
    
        barisLanjut = batesBawah + 1
    
        batesBawah = UBound(LowWordArray) + barisLanjut
        buatArray = 0
        For j = barisLanjut To (batesBawah)
            Sheets("Data").Cells(j, 1).Value = categoryBarang
            Sheets("Data").Cells(j, 2).Value = "Low"
            Sheets("Data").Cells(j, 3).Value = Trim(LowWordArray(buatArray))
        
            If buatArray < UBound(LowWordArray) Then
                buatArray = buatArray + 1
            End If
        Next j
        
        batesBawah = batesBawah + 1
        
        Erase HighWordArray, MedWordArray, LowWordArray
        barisLoop = barisLoop + 1
    Loop
    
End Sub
