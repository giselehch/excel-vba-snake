Attribute VB_Name = "Module4"
' this module is used to generate food
Dim foodR As Long
Dim foodC As Long

Public foodColle As New Collection      'a collection to write down all the positions of the food
Public threeFood As Integer         'Snake will increase length when eat three fruits




Sub generateFood()
    Randomize
    
    foodR = Int((posiRF - posiRI + 1) * Rnd) + posiRI
    foodC = Int((posiCF - posiCI + 1) * Rnd) + posiCI
    
    Cells(foodR, foodC).Interior.Color = vbGreen
    foodColle.Add foodR & ", " & foodC
End Sub




