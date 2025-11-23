Private Sub copi()

    Dim Var1 As Shape
    Dim Var2 As Shape
    Dim Var3 As Shape
    Dim Var4 As Shape
    
    ' Pastikan ada objek yang dipilih
    If ActiveShape Is Nothing Then
        MsgBox "Tidak ada objek yang dipilih.", vbExclamation
        Exit Sub
    End If
    
    ' copi objek
    Set Var1 = ActiveShape
    Set Var2 = Var1.Duplicate
    Set Var3 = Var1.Duplicate
    Set Var4 = Var1.Duplicate
    
    ActiveDocument.Unit = cdrMillimeter         ' merubah satuan menjadi mm
    ActiveDocument.ReferencePoint = cdrCenter   ' fokus objek berada di center
    
    ' menyusun objek
    Offset = Var1.SizeWidth + 2
    
    Var2.Move Offset, 0
    Var3.Move (Offset * 2), 0
    Var4.Move (Offset * 3), 0
    
    ' memilih objek
    ActiveDocument.ClearSelection   ' kosongkan selection dulu
    ActiveDocument.AddToSelection Var1, Var2, Var3, Var4
    
    ' menyatukan objek
    Dim Varall As Shape
    Set Varall = ActiveSelection.Group
    
    ' set objek ketengah halaman
    Varall.AlignAndDistribute 3, 3, 2, 0, False, 2

End Sub

Sub Cetak_Pas_Foto()

    If ActiveShape Is Nothing Then
        MsgBox "Tidak ada objek yang dipilih.", vbExclamation
        Exit Sub
    End If
    
    ActiveDocument.Unit = cdrMillimeter         ' merubah satuan menjadi mm
    
    Dim f As form_4x6
    
    If ActiveShape.SizeWidth = 40 And ActiveShape.SizeHeight = 60 Then
    
        Set f = New form_4x6
        f.Show
        
        If f.IsCancelled Then
            Exit Sub
        End If
        
        Call Empat_Enam(f.Jumlah)
        
    Else
        MsgBox "Ukuran Salah.", vbExclamation
    End If
    
End Sub

Private Sub Empat_Enam(Jumlah As Integer)
    MsgBox Jumlah, vbExclamation
End Sub
