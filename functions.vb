Function getOptionType(option_data As String)
    ' If argument is null, return null.
    If option_data = "" Then GoTo ErrorHandl

    ' Create array to hold option types.
    Dim optionTypes As Variant
    Dim optionType As Variant 

    ' Populate array with option types.
    optionTypes = Array("IRON CONDOR","BUTTERFLY","CALENDAR","DIAGONAL","COMBO","VERTICAL","CALL","PUT")

    ' Search each option type for a match in the option_data string that indicates the option type.
    For Each optionType In optionTypes
        ' InStr returns the position of the matching keyword in the string. If the return is 0, no match was found.
        If InStr(option_data, optionType) Then
            oType = optionType
            Exit For
        End If
    Next

    ' If there are no matches in option_data, return null.
    If oType = 0 Then GoTo ErrorHandl

    getOptionType = oType
    Exit Function
ErrorHandl:
    getOptionType = ""
End Function