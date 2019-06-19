Function getOptionType(option_data As String)
    ' Error handling.
    If option_data = "" Then GoTo ErrorHandl

    ' Create array to hold option types.
    Dim optionTypes As Variant
    Dim optionType As Variant 

    ' Populate array with option types.
    optionTypes = Array("IRON CONDOR","BUTTERFLY","Calendar","Diagonal","Combo","Vertical","Call","Put")

    ' Search each option type for a match in the option_data string that indicates the option type.
    For Each optionType In optionTypes
        ' InStr returns the position of the matching keyword in the string. If the return is 0, no match was found.
        If InStr(option_data, optionType) Then
            oType = optionType
            Exit For
        End If
    Next

    ' Example of for loop logic:
    ' If InStr(option_data, "IRON CONDOR") Then
    '   optionType = "Iron Condor"
    ' ElseIf InStr(option_data, "BUTTERFLY")
    '   optionType = "Butterfly"
    ' End If

    getOptionType = oType
    Exit Function
ErrorHandl:
    getOptionType = ""
End Function