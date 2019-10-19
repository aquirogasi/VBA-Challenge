Attribute VB_Name = "AddressOfMin_Function"
'Funtion to find the address of the Min Value in a specific range


Function AddressOfMin(rng As Range) As Range
    
    Set AddressOfMin = rng.Cells(WorksheetFunction.Match(WorksheetFunction.Min(rng), rng, 0))

End Function
