Function getOptionType(trade_order As String)
    ' If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl

    ' Create array to hold option types.
    Dim optionTypes As Variant
    Dim optionType As Variant 

    ' Populate array with option types.
    optionTypes = Array("IRON CONDOR","BUTTERFLY","CALENDAR","DIAGONAL","COMBO","VERTICAL","CALL","PUT")

    ' Search each option type for a match in the trade_order string that indicates the option type.
    For Each optionType In optionTypes
        ' InStr returns the position of the matching keyword in the string. If the return is 0, no match was found.
        If InStr(trade_order, optionType) Then
            oType = optionType
            Exit For
        End If
    Next

    ' If there are no matches in trade_order, return null.
    If oType = 0 Then GoTo ErrorHandl

    ' Convert string to sentence case.
    oType = WorksheetFunction.Proper(oType)

    ' Return option type.
    getOptionType = oType
    Exit Function
ErrorHandl:
    getOptionType = ""
End Function

Function getNthWord(text As String, start_num As Integer, Optional num_words As Integer=1)
Dim current_pos As Long
Dim char_len As Integer
Dim current_word_num As Integer
 
getNthWord = ""
current_word_num = 1
 
'Remove leading spaces.
text = Trim(text)

' Get the number of characters in the text.
char_len = Len(text)

' Find the character position of the nth word.
For current_pos = 1 To char_len
    ' If this is the nth word, concatinate each character until it is no longer the nth + num_words word.
    If (current_word_num >= start_num) And (current_word_num <= start_num + num_words - 1) Then
        getNthWord = getNthWord & Mid(text, current_pos, 1)
    End If

    ' If there is a space after this character, increment the current_word_num by 1.
    If (Mid(text, current_pos, 1) = " ") Then
      current_word_num = current_word_num + 1
    End If
Next current_pos
 
'Remove the rightmost space.
getNthWord = Trim(getNthWord)
 
End Function

Function getExpiration(trade_order As String, option_type As String)
    Select Case option_type
    ' If option type matches case strings, return the 6th-8th word.
    Case "Combo", "Vertical", "Butterfly"
      expDate = getNthWord(trade_order, 6, 3)
    ' If option type matches case strings, parse the 8th word and return it with the 9th and 10th word concatenated.
    Case "Diagonal", "Calendar"
      expDay = getNthWord(trade_order, 8)
      expDay = Split(expDay, "/")
      expDate = expDay(1) & " " & getNthWord(trade_order, 9, 2)
    ' If option type matches case strings, return the 5th-7th word.
    Case "Call", "Put"
      expDate = getNthWord(trade_order,5,3)
    ' If option type is an Iron Condor, return the 7th-9th word.
    Case "Iron Condor"
      expDate = getNthWord(trade_order,7,3)
    ' If there is no match, return null.
    Case Else
      GoTo ErrorHandl
    End Select

    ' Return the concatonated string as a date value.
    getExpiration = DateValue(expDate)
    Exit Function
ErrorHandl:
    getExpiration = ""
End Function

Function getSymbol(trade_order As String, option_type As String)
    ' If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    
    Select Case option_type
    ' If option type is an Iron Condor, return the 5th word.
    Case "Iron Condor"
      Symbol = getNthWord(trade_order, 5)
    ' Otherwise return the 4th word.
    Case "Call", "Put"
      Symbol = getNthWord(trade_order, 3)
    Case Else
      Symbol = getNthWord(trade_order, 4)
    End Select

    ' Return the concatonated string as a date value.
    getSymbol = Symbol
    Exit Function
ErrorHandl:
    getSymbol = ""
End Function

Function daysTillExp(trade_date As Date, expiration_date As Date)
    ' If argument is null, return null.
    ' TODO: Handle #VALUE error from wrong datatype.
    If trade_date = 0 Then GoTo ErrorHandl

    ' Return the concatonated string as a date value.
    daysTillExp = expiration_date - trade_date
    Exit Function
ErrorHandl:
    daysTillExp = ""
End Function

Function getStrategy(trade_order As String, option_type As String)
    ' If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl

      ' Test market position.
      If InStr(trade_order, "BOT") Then
        position = "Long"
      Else
        position = "Short"
      End If

      ' Test option type.
      If InStr(trade_order, "CALL") Then
        side = "Call"
      ElseIf InStr(trade_order, "PUT") Then
        side = "Put"
      End If

    Select Case option_type
    ' Match strategy to posture.
    Case "Vertical", "Diagonal"
     ' Concatenate market position, option side, and option type to determine strategy.
      strategy = position & " " & side &  " " & option_type
    Case "Call", "Put"
      strategy = position & " " & option_type
    Case "Combo"
      strategy = position & " Synthetic"
    Case Else
      ' For Iron Condor, Butterfly, and Calendar.
      strategy = option_type
    End Select

    ' Return the concatonated string as a date value.
    getStrategy = strategy
    Exit Function
ErrorHandl:
    getStrategy = ""
End Function