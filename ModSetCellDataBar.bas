Attribute VB_Name = "ModSetCellDataBar"
Option Explicit

'SetCellDataBar・・・元場所：FukamiAddins3.ModCell



Public Sub SetCellDataBar(TargetCell As Range, Ratio As Double, Color As Long)
'セルの書式設定で0～1の値に基づいて、データバーを設定する
'20210820

'TargetCell :対象のセル
'Ratio      :割合（0～1）
'Color      :バーの色（RGB値）

    Dim Gosa As Double
    Gosa = 10 ^ (-10) '←←←←←←←←←←←←←←←←←←←←←←←
    
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


