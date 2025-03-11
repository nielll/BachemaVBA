Attribute VB_Name = "BachemaFunctions"
Function KontrollmessungInnerhalb(KontrollName As String, Sollwert As Range, SuchSpalteRange As Range, ProbenNameRange As Range, MesswertRange As Range, MissingIsValid As Boolean, AbwProzent As Double) As Boolean
    Dim ProbenRow As Long
    Dim i As Long, j As Long
    Dim Messwert As Double
    Dim KontrollAbove As Integer, KontrollBelow As Integer
    Dim Suchwert As String
    Dim Probenwert As String
    
    ' Werte initialisieren
    KontrollmessungInnerhalb = False
    FoundAbove = False
    FoundBelow = False
    
    ' Zeile der Probe in SuchSpalteRange suchen, indem für jede Zeile der kombinierte Suchwert erstellt wird
    Probenwert = ""
    For i = 1 To ProbenNameRange.Columns.Count
        Probenwert = Probenwert & ProbenNameRange.Cells(1, i).value
    Next i
    
    ' Probenwert erstellen durch Kombination der Werte in jeder Spalte der ProbenNameRange für die gleiche Zeile
    For i = 1 To SuchSpalteRange.Rows.Count
        Suchwert = ""
        For j = 1 To SuchSpalteRange.Columns.Count
            Suchwert = Suchwert & SuchSpalteRange.Cells(i, j).value
        Next j
        
        If Suchwert = Probenwert Then
            ProbenRow = SuchSpalteRange.Cells(i, 1).row
            Exit For
        End If
    Next i
        
    ' Wenn die ProbenRow gefunden wurde, überprüfen wir die Kontrollmessungen
    If Probenwert <> "" And ProbenRow > 0 Then
        ' Überprüfen der nächstgelegenen Kontrollmessung oberhalb
        'Debug.Print KontrollName & " ; " & Probenwert & " : " & ProbenRow & " : " & SuchSpalteRange.Cells(1, 1).row
        For j = ProbenRow To SuchSpalteRange.Cells(1, 1).row Step -1
            i = j - SuchSpalteRange.row
            
            If SuchSpalteRange.Cells(i, 1).value = "" Then
                Exit For
            End If
            
            If SuchSpalteRange.Cells(i, 1).value = KontrollName Then
                Messwert = MesswertRange.Cells(i, 1).value
                If Messwert >= Sollwert.value * (1 - (AbwProzent / 100)) And Messwert <= Sollwert.value * (1 + (AbwProzent / 100)) Then
                    KontrollAbove = 2
                Else
                    KontrollAbove = 1
                End If
                Exit For
            End If
        Next j

        ' Überprüfen der nächstgelegenen Kontrollmessung unterhalb
        'Debug.Print KontrollName & " ; " & Probenwert & " : " & ProbenRow & " : " & SuchSpalteRange.Cells(SuchSpalteRange.Rows.Count, 1).row
        For j = ProbenRow + 1 To SuchSpalteRange.Cells(SuchSpalteRange.Rows.Count, 1).row
            i = j - SuchSpalteRange.row + 1
            
            ' Wir gehen aus dem loop sobald eine leere zelle kommt
            If SuchSpalteRange.Cells(i, 1).value = "" Then
                Exit For
            End If
            
            
            If SuchSpalteRange.Cells(i, 1).value = KontrollName Then
                Messwert = MesswertRange.Cells(i, 1).value
                If Messwert >= Sollwert.value * (1 - (AbwProzent / 100)) And Messwert <= Sollwert.value * (1 + (AbwProzent / 100)) Then
                    KontrollBelow = 2
                Else
                    KontrollBelow = 1
                End If
                Exit For
            End If
        Next j

        ' Wenn sowohl oberhalb als auch unterhalb Werte gefunden wurden, kann die Funktion beendet werden
        KontrollmessungInnerhalb = IIf(MissingIsValid, (KontrollAbove = 0 Or KontrollAbove = 2) And (KontrollBelow = 0 Or KontrollBelow = 2), KontrollAbove = 2 And KontrollBelow = 2)
        Exit Function
    End If
End Function
Function KontrollmessungInnerhalbNeu(SuchSpalteRange As Range, ProbenNameRange As Range, MesswertRange As Range, MissingIsValid As Boolean, ToleranzSubstanzen As Range) As Boolean
    Dim ProbenRow As Long
    Dim i As Long, j As Long
    Dim Messwert As Double
    Dim KontrollAbove As Integer, KontrollBelow As Integer
    Dim Suchwert As String
    
    ' Werte initialisieren
    KontrollmessungInnerhalbNeu = False
    FoundAbove = False
    FoundBelow = False
    
    ' Zeile der Probe in SuchSpalteRange suchen, indem für jede Zeile der kombinierte Suchwert erstellt wird
    Probenwert = ""
    For i = 1 To ProbenNameRange.Columns.Count
        Probenwert = Probenwert & ProbenNameRange.Cells(1, i).value
    Next i
    
    ' Probenwert erstellen durch Kombination der Werte in jeder Spalte der ProbenNameRange für die gleiche Zeile
    For i = 1 To SuchSpalteRange.Rows.Count
        Suchwert = ""
        For j = 1 To SuchSpalteRange.Columns.Count
            Suchwert = Suchwert & SuchSpalteRange.Cells(i, j).value
        Next j
        
        If Suchwert = Probenwert Then
            ProbenRow = SuchSpalteRange.Cells(i, 1).row
            Exit For
        End If
    Next i
    
    ' Wenn die ProbenRow gefunden wurde, überprüfen wir die Kontrollmessungen
    If Probenwert <> "" And ProbenRow > 0 Then
    
        ' Überprüfen der nächstgelegenen Kontrollmessung unterhalb
        'Debug.Print KontrollName & " ; " & Probenwert & " : " & ProbenRow & " : " & SuchSpalteRange.Cells(SuchSpalteRange.Rows.Count, 1).row
        For j = ProbenRow + 1 To SuchSpalteRange.Cells(SuchSpalteRange.Rows.Count, 1).row
            i = j - SuchSpalteRange.row + 1
            
            ' Wir gehen aus dem loop sobald eine leere zelle kommt
            If SuchSpalteRange.Cells(i, 1).value = "" Then
                Exit For
            End If
            
            For ii = 1 To ToleranzSubstanzen.Columns(1).Cells().Count
                If ToleranzSubstanzen.Columns(1).Cells(ii, 1).value = SuchSpalteRange.Cells(i, 1).value And ToleranzSubstanzen.Columns(2).Cells(ii, 1).value <> "" And ToleranzSubstanzen.Columns(3).Cells(ii, 1).value <> "" Then
                    Messwert = MesswertRange.Cells(i, 1).value
                    
                    If Messwert >= ToleranzSubstanzen.Columns(2).Cells(ii, 1).value * 1 And Messwert <= ToleranzSubstanzen.Columns(3).Cells(ii, 1).value * 1 Then
                        KontrollBelow = 2
                    Else
                        KontrollBelow = 1
                    End If
                    
                    Exit For
                End If
            Next ii
            
            If KontrollBelow > 0 Then
                Exit For
            End If
        Next j
        
        
        If KontrollBelow <> 1 Then
            ' Überprüfen der nächstgelegenen Kontrollmessung oberhalb
            'Debug.Print KontrollName & " ; " & Probenwert & " : " & ProbenRow & " : " & SuchSpalteRange.Cells(1, 1).row
            For j = ProbenRow To SuchSpalteRange.Cells(1, 1).row Step -1
                i = j - SuchSpalteRange.row
                
                If SuchSpalteRange.Cells(i, 1).value = "" Then
                    Exit For
                End If
                
                For ii = 1 To ToleranzSubstanzen.Columns(1).Cells().Count
                    If ToleranzSubstanzen.Columns(1).Cells(ii, 1).value = SuchSpalteRange.Cells(i, 1).value And ToleranzSubstanzen.Columns(2).Cells(ii, 1).value <> "" And ToleranzSubstanzen.Columns(3).Cells(ii, 1).value <> "" Then
                        Messwert = MesswertRange.Cells(i, 1).value
                        
                        If Messwert >= ToleranzSubstanzen.Columns(2).Cells(ii, 1).value * 1 And Messwert <= ToleranzSubstanzen.Columns(3).Cells(ii, 1).value * 1 Then
                            KontrollAbove = 2
                        Else
                            KontrollAbove = 1
                        End If
                        
                        Exit For
                    End If
                Next ii
                
                If KontrollAbove > 0 Then
                    Exit For
                End If
            Next j
        End If

        ' Wenn sowohl oberhalb als auch unterhalb Werte gefunden wurden, kann die Funktion beendet werden
        KontrollmessungInnerhalbNeu = IIf(MissingIsValid, (KontrollAbove = 0 Or KontrollAbove = 2) And (KontrollBelow = 0 Or KontrollBelow = 2), KontrollAbove = 2 And KontrollBelow = 2)
        Exit Function
    End If
End Function
Function IsOrgC(SuchSpalteRange As Range, ProbenNameRange As Range, ToleranzSubstanzen As Range) As Boolean
    Dim ProbenRow As Long
    Dim i As Long, j As Long
    Dim FoundAbove As Boolean, FoundBelow As Boolean
    Dim Suchwert As String
    Dim Probenwert As String
    
    ' Werte initialisieren
    FoundAbove = False
    FoundBelow = False
    
    ' Zeile der Probe in SuchSpalteRange suchen, indem für jede Zeile der kombinierte Suchwert erstellt wird
    Probenwert = ""
    For i = 1 To ProbenNameRange.Columns.Count
        Probenwert = Probenwert & ProbenNameRange.Cells(1, i).value
    Next i
    
    ' Probenwert erstellen durch Kombination der Werte in jeder Spalte der ProbenNameRange für die gleiche Zeile
    For i = 1 To SuchSpalteRange.Rows.Count
        Suchwert = ""
        For j = 1 To SuchSpalteRange.Columns.Count
            Suchwert = Suchwert & SuchSpalteRange.Cells(i, j).value
        Next j
        
        If Suchwert = Probenwert Then
            ProbenRow = SuchSpalteRange.Cells(i, 1).row
            Exit For
        End If
    Next i
        
    ' Wenn die ProbenRow gefunden wurde, überprüfen wir die Kontrollmessungen
    If Probenwert <> "" And ProbenRow > 0 Then
        ' Überprüfen der nächstgelegenen Kontrollmessung unterhalb
        'Debug.Print KontrollName & " ; " & Probenwert & " : " & ProbenRow & " : " & SuchSpalteRange.Cells(SuchSpalteRange.Rows.Count, 1).row
        For j = ProbenRow + 1 To SuchSpalteRange.Cells(SuchSpalteRange.Rows.Count, 1).row
            i = j - SuchSpalteRange.row + 1
            
            ' Wir gehen aus dem loop sobald eine leere zelle kommt
            If SuchSpalteRange.Cells(i, 1).value = "" Then
                Exit For
            End If
            
            For ii = 1 To ToleranzSubstanzen.Columns(1).Cells().Count
                If ToleranzSubstanzen.Columns(1).Cells(ii, 1).value = SuchSpalteRange.Cells(i, 1).value And ToleranzSubstanzen.Columns(2).Cells(ii, 1).value <> "" And ToleranzSubstanzen.Columns(3).Cells(ii, 1).value <> "" Then
                    FoundBelow = True
                    Exit For
                End If
            Next ii
        Next j
        
        
        If FoundBelow Then
            ' Überprüfen der nächstgelegenen Kontrollmessung oberhalb
            'Debug.Print KontrollName & " ; " & Probenwert & " : " & ProbenRow & " : " & SuchSpalteRange.Cells(1, 1).row
            For j = ProbenRow To SuchSpalteRange.Cells(1, 1).row Step -1
                i = j - SuchSpalteRange.row
                
                If SuchSpalteRange.Cells(i, 1).value = "" Then
                    Exit For
                End If
            
                For ii = 1 To ToleranzSubstanzen.Columns(1).Cells().Count
                    If ToleranzSubstanzen.Columns(1).Cells(ii, 1).value = SuchSpalteRange.Cells(i, 1).value And ToleranzSubstanzen.Columns(2).Cells(ii, 1).value <> "" And ToleranzSubstanzen.Columns(3).Cells(ii, 1).value <> "" Then
                        FoundAbove = True
                        Exit For
                    End If
                Next ii
            Next j
        End If

        ' Wenn sowohl oberhalb als auch unterhalb Werte gefunden wurden, kann die Funktion beendet werden
        IsOrgC = FoundAbove And FoundBelow
        Exit Function
    End If
    
    IsOrgC = False
End Function
