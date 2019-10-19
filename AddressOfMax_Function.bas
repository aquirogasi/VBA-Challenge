Attribute VB_Name = "AddressOfMax_Function"
'Funtion to find the address of the Max Value in a specific range

Function AddressOfMax(rng As Range) As Range
    
    Set AddressOfMax = rng.Cells(WorksheetFunction.Match(WorksheetFunction.Max(rng), rng, 0))

End Function
