Attribute VB_Name = "ModSetCellDataBar"
Option Explicit

'SetCellDataBarEEE³êFFukamiAddins3.ModCell



Public Sub SetCellDataBar(TargetCell As Range, Ratio As Double, Color As Long)
'ZÌ®ÝèÅ0`1ÌlÉîÃ¢ÄAf[^o[ðÝè·é
'20210820

'TargetCell :ÎÛÌZ
'Ratio      :i0`1j
'Color      :o[ÌFiRGBlj

    Dim Gosa As Double
    Gosa = 10 ^ (-10) '©©©©©©©©©©©©©©©©©©©©©©©
    
    With TargetCell
        .Interior.Pattern = xlPatternLinearGradient
        .Interior.Gradient.Degree = 0
        
        With .Interior.Gradient.ColorStops
            If Ratio > Gosa Then
                .Add(0).Color = Color
                .Add(Gosa).Color = Color
                .Add(Gosa * 2).Color = Color
                
                If Gosa * 3 < Ratio Then
                    .Add(Ratio).Color = Color
                Else
                    .Add(Gosa * 3).Color = Color
                End If
            End If
            
            If Ratio < 1 Then
                If Ratio + Gosa > 1 Then
                    .Add((1 + Ratio) / 2).Color = Color
                Else
                    .Add(Ratio + Gosa).Color = rgbWhite
                End If
                .Add(1).Color = rgbWhite
            End If
        End With
    End With

End Sub


