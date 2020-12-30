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

Function getPosture(option_strategy As String)
    ' If argument is null, return null.
    If option_strategy = "" Then GoTo ErrorHandl

    Select Case option_strategy
    ' Match option strategy to its corresponding market posture.
    Case "Iron Condor", "Calendar", "Butterfly"
      posture = "Neutral"
    Case "Long Call", "Short Put", "Long Call Vertical", "Short Put Vertical", "Long Call Diagonal", "Long Synthetic"
      posture = "Bullish"
    Case "Short Call", "Long Put", "Long Put Diagonal", "Short Synthetic"
      posture = "Bearish"
    End Select

    ' Return the posture of the trade.
    getPosture = posture
    Exit Function
ErrorHandl:
    getPosture = ""
End Function

Function getStockQuote(symobol As String, trade_date As Date)
    ' If argument is null, return null.
    If symobol = "" Then GoTo ErrorHandl
    
    ' Get date as YYYYMMDD.
    dateAsString = Replace(Format(trade_date, "yyyy/mm/dd"), "/", "")

    ' Define iex's defaul strings.
    site = "https://sandbox.iexapis.com/stable/stock/"
    Chart = "/chart/"
    byDate = "/chart/date/"
    chartByDate = "?chartByDay=true"
    token = "?token=Tpk_46cc7266d1374035bb1cb11d639476c4"

    ' Create array to hold acceptable range parameters.
    Dim ranges As Variant
    Dim range As Variant

    ' Populate array with range values.
    ' TODO: Test to see if date falls within range, and use select it accordingly.
    ranges = Array("5d", "1m", "3m", "6m", "1y", "2y", "5y", "max")

    ' Concatonate strings to formulate url.
    ' TODO: Troubleshoot chartByDate booleon failure and use it instead of a range (see documentation).
    Url = site & symobol & Chart & ranges(5) & token

    ' Get json object from url.
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", Url, False
    objHTTP.send
    
    ' Format date to search the JSON by.
    dateHyphenated = Format(trade_date, "yyyy-mm-dd")
    regPattern = "(({""date"":"")" & dateHyphenated & "[^}]*})"

    ' Use regex to isolate single stock quote.
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = regPattern: regex.Global = False
    Set matches = regex.Execute(objHTTP.responseText)
    stockQuote = matches(0)

    ' TODO: hande #VALUE error due to no regex matches.

    ' Return the stock quote.
    getStockQuote = stockQuote
    Exit Function
ErrorHandl:
    getStockQuote = ""
End Function

Function getQuoteValue(quote_key As String, stock_quote As String)
    ' If argument is null, return null.
    If stock_quote = "" Then GoTo ErrorHandl

    ' Concatonate desired JSON key with regex commands to trap the value in a capturing group.
    regPattern = """" & quote_key & """:([^,]*)"

    ' Use regex to return the value of the key passed in.
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = regPattern: regex.Global = False
    Set matches = regex.Execute(stock_quote)
    quoteValue = matches(0).SubMatches(0)

    getQuoteValue = quoteValue
    Exit Function
ErrorHandl:
    getQuoteValue = ""
End Function

Function getPrem(trade_order As String, option_type As String)
    ' If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    
    ' Get the nth word based on option type.
    Select Case option_type
    Case "Iron Condor"
      premium = getNthWord(trade_order, 12)
    Case "Calendar", "Diagonal"
      premium = getNthWord(trade_order, 13)
    Case "Call", "Put"
      premium = getNthWord(trade_order, 10)
    Case Else
      premium = getNthWord(trade_order, 11)
    End Select

    ' Remove @ symbol and add 0 to cast as numeric.
    premium = Replace(premium,"@","") + 0

    ' Return the option premium.
    getPrem = premium
    Exit Function
ErrorHandl:
    getPrem = ""
End Function

Function getMaxProfit(trade_order As String, option_type As String, qty As String, prem As String, Optional comm As Currency = 0#, Optional risk As Currency = 0#)
    ' If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    
    ' Convert strings to numbers. (Currency and Integer won't accept "" as an argument, which results in a #VALUE error.)
    Dim cPrem As Currency
    cPrem = CCur(prem)
    
    ' Get the nth word based on option type.
    Select Case option_type
    Case "Iron Condor"
      maxProfit = cPrem
    Case "Butterfly"
      ' Get strikes.
      strikes = Split(getNthWord(trade_order, 9), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      strike3 = strikes(2)

      ' Get the distance between the short strike and long strike.
      spread1 = Abs(strike1 - strike2)
      spread2 = Abs(strike3 - strike2)

      ' Whichever spread is smaller is the max profit.
      If spread1 <= spread2 Then
        maxProfit = spread1 - cPrem
      Else
        maxProfit = spread2 - cPrem
      End If
    Case "Calendar"
      maxProfit = risk * 2
      ' TODO: getMaxProfit = maxProfit, Exit Function
    Case "Diagonal"
      ' If debit paid was negitive (actually a credit).
      If cPrem < 0 Then
        maxProfit = Abs(cPrem)
      Else
        ' Get strikes.
        strikes = Split(getNthWord(trade_order, 11), "/")
        strike1 = strikes(0)
        strike2 = strikes(1)

        ' Spread Width - Premium + Extrinsic Value.
        ' TODO: Use Black-scholes Model to calculate extrinsic value.
        maxProfit = Abs(strike1 - strike2) - cPrem
      End If
    Case "Combo"
        ' If long, stock has unlimited profit potential.
        If InStr(trade_order, "BOT") Then
          getMaxProfit = "Undefined"
          Exit Function
        ' If short, stock can only drop to 0 (finite profits).
        Else
          maxProfit = getNthWord(trade_order, 9) + cPrem
        End If
    Case "Call"
        ' If long, call has unlimited profit potential.
        If InStr(trade_order, "BOT") Then
          getMaxProfit = "Undefined"
          Exit Function
        ' If short, call can only expire worthless.
        Else
          maxProfit = cPrem
        End If
    Case "Put"
      ' If long, put has unlimited profit potential.
        If InStr(trade_order, "BOT") Then
          maxProfit = getNthWord(trade_order, 8) - cPrem
        ' If short, put can only expire worthless.
        Else
          maxProfit = cPrem
        End If
    ' Case "Vertical"
    Case Else
      ' If short, max profit is credit recieved.
      If InStr(trade_order, "SOLD") Then
        maxProfit = cPrem
      ' If short, call can only expire worthless.
      Else
        ' Get strikes.
        strikes = Split(getNthWord(trade_order, 9), "/")
        strike1 = strikes(0)
        strike2 = strikes(1)

        maxProfit = Abs(strike1 - strike2) - cPrem
      End If
    End Select
    
    ' Multiply by shares being controlled and subtract commissions.
    maxProfit = maxProfit * Abs(qty) * 100 - comm

    ' Return the option maxProfit.
    getMaxProfit = maxProfit
    Exit Function
ErrorHandl:
    getMaxProfit = ""
End Function

Function getRisk(trade_order As String, option_type As String, qty As String, prem As String, max_profit As String, Optional comm As Currency = 0)
    ' If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    
    ' TODO: Replace arguments with functions.
    ' qty = getNthWord(trade_order,2)
    ' prem = getPrem(trade_order, option_type)

    ' Covert strings to numbers. (Currency and Integer won't accept "" as an argument, which results in a #VALUE error.)
    Dim premium As Currency
    premium = CCur(prem)
    Dim maxProfit As Currency
    If IsNumeric(max_profit) Then maxProfit = CCur(max_profit)
    
    ' Get the nth word based on option type.
    Select Case option_type
    Case "Iron Condor"
      strikes = Split(getNthWord(trade_order, 10), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      strike3 = strikes(2)
      strike4 = strikes(3)

      ' Get the distance between the short strike and long strike.
      spread1 = Abs(strike2 - strike1)
      spread2 = Abs(strike3 - strike4)

      If spread1 > spread2 Then
        risk = spread1
        GoTo IronCondor
      Else
        risk = spread2
        GoTo IronCondor
      End If
    Case "Vertical"
      ' If short, risk is spread width less credit recieved.
      If InStr(trade_order, "SOLD") Then
        ' Get strikes.
        strikes = Split(getNthWord(trade_order, 9), "/")
        strike1 = strikes(0)
        strike2 = strikes(1)

        risk = Abs(strike1 - strike2) - premium
      ' If long, risk is debit paid.
      Else
        risk = premium
      End If
    Case "Diagonal"
      If premium < 0 Then
        ' Get strikes.
        strikes = Split(getNthWord(trade_order, 11), "/")
        strike1 = strikes(0)
        strike2 = strikes(1)
        spread = Abs(strike1 - strike2)

        risk = spread - Abs(premium)
      Else
        risk = premium
      End If
    Case "Combo"
      ' If short, risk is unlimited.
      If InStr(trade_order, "SOLD") Then
        getRisk = "Undefined"
        Exit Function
      ' If long, risk is capped by the floor.
      Else
        strike1 = getNthWord(trade_order, 9)
        risk = strike1 - premium
      End If
    Case "Call"
      ' If short, risk is unlimited.
      If InStr(trade_order, "SOLD") Then
        getRisk = "Undefined"
        Exit Function
      ' If long, risk is debit paid.
      Else
        risk = premium
      End If
    Case "Put"
      ' If short, risk is capped by the floor.
      If InStr(trade_order, "SOLD") Then
        strike1 = getNthWord(trade_order, 9)
        risk = strike1 - premium
      ' If long, risk is debit paid.
      Else
        risk = premium
      End If
    Case Else
      risk = premium
    End Select
    
    ' Multiply by shares being controlled and add commissions.
    risk = risk * Abs(qty) * 100 + comm

    ' Return the option risk.
    getRisk = risk
    Exit Function
IronCondor:
    risk = risk * Abs(qty) * 100 + comm
    getRisk = risk - maxProfit
    Exit Function
Standard:
    risk = premium
    risk = risk * Abs(qty) * 100 + comm
    getRisk = risk
ErrorHandl:
    getRisk = ""
End Function

Function GetPlCLose(trade_order As String, option_type As String, prem As String, max_profit As String, Optional comm As Currency = 0#)
    ' If argument is null, return null.
    If trade_order = "" Or prem = "" Then GoTo ErrorHandl
    
    ' Covert strings to numbers.
    Dim premium As Currency
    premium = CCur(prem)
    Dim maxProfit As Currency
    If IsNumeric(max_profit) Then maxProfit = CCur(max_profit)

    ' How many shares being controlled?
    contracts = Abs(getNthWord(trade_order,2))
    shares = contracts * 100
    debit = premium * shares

    ' Get the nth word based on option type.
    Select Case option_type
    Case "Iron Condor"
      ' Opening Credit - Closing Debit
      plClose = maxProfit - debit
      ' Skip comissions -- Max profits accounts for it.
    Case "Combo"
      credit = Replace(getNthWord(trade_order, 11),"@","") + 0
      plClose = debit + (credit*shares) - comm
    Case "Vertical"
      If InStr(trade_order, "SOLD") Then
        plClose = maxProfit - debit
      Else
        entryDebit = getPrem(trade_order, option_type)
        risk = getRisk(trade_order, option_type, CStr(contracts), CStr(entryDebit), max_profit, comm)
        credit = debit
        plClose = credit - risk
      End If
    Case Else
      credit = getPrem(trade_order, option_type)
      ' If credit was recieved...
      If (InStr(trade_order, "BOT") And credit < 0) Or InStr(trade_order, "SOLD") Then
        plClose = maxProfit - debit
      Else
        risk = getRisk(trade_order, option_type, CStr(contracts), CStr(credit), max_profit, comm)
        credit = debit
        plClose = credit - risk
      End If
    End Select

    ' Return the option plClose.
    GetPlCLose = plClose
    Exit Function
ErrorHandl:
    GetPlCLose = ""
End Function

Function GetPlPercent(pl_closed As String, max_profit As String, max_risk As String)
 
    ' If argument is null, return null.
    If pl_closed = "" Then GoTo ErrorHandl

    ' Covert strings to numbers.
    Dim plClose As Currency
    plClose = CCur(pl_closed)
    Dim maxProfit As Currency
    If IsNumeric(max_profit) Then maxProfit = CCur(max_profit) Else GoTo Undefined
    Dim risk As Currency
    If IsNumeric(max_risk) Then risk = CCur(max_risk) Else GoTo Undefined

    ' If trade was profitable, show percent of max profit
    If plClose > 0 Then
      plPercent = plClose / maxProfit
    ' If trade lost money, show percent of max loss
    Else
      plPercent = plClose / risk
    End If

    ' Return the option plPercent.
    GetPlPercent = plPercent
    Exit Function
Undefined:
    GetPlPercent = 0
    Exit Function
ErrorHandl:
    GetPlPercent = ""
End Function

Function GetMonth(mon_abr As String)
  monthNumber = Application.Match(mon_abr, Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"), 0)
  getMonth = Evaluate(TEXT(monthNumber, "00"))
End Function

Function GetOptionSignature(trade_order As String)
    If trade_order = "" Then GoTo ErrorHandl

    option_type = getOptionType(trade_order)

    Select Case option_type
    ' Single Leg Options (AND P/L%)
    Case "Call", "Put"
      optionSignature = "." & getNthWord(trade_order, 3) & getNthWord(trade_order, 7) & GetMonth(getNthWord(trade_order, 6)) & getNthWord(trade_order, 5) & Left(getNthWord(trade_order, 9), 1) & getNthWord(trade_order, 8)
    Case "Iron Condor"
      signatureBase = getNthWord(trade_order, 5) & getNthWord(trade_order, 9) & GetMonth(getNthWord(trade_order, 8)) & getNthWord(trade_order, 7)
      types = Split(getNthWord(trade_order, 11), "/")
      type1 = Left(types(0), 1)
      type2 = Left(types(1), 1)
      strikes = Split(getNthWord(trade_order, 10), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      strike3 = strikes(2)
      strike4 = strikes(3)
      optionSignature = "." & signatureBase & type1 & strike1 & "-." & signatureBase & type1 & strike2 & "+." & signatureBase & type2 & strike3 & "-." & signatureBase & type2 & strike4
    Case "Vertical"
      signatureBase = getNthWord(trade_order, 4) & getNthWord(trade_order, 8) & GetMonth(getNthWord(trade_order, 7)) & getNthWord(trade_order, 6) & Left(getNthWord(trade_order, 10), 1)
      strikes = Split(getNthWord(trade_order, 9), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      optionSignature = "." & signatureBase & strike1 & "-." & signatureBase & strike2
    Case "Diagonal"
      dates = Split(getNthWord(trade_order, 8), "/")
      date1 = dates(0) & GetMonth(getNthWord(trade_order, 7)) & getNthWord(trade_order, 6)
      date2 = getNthWord(trade_order, 10) & GetMonth(getNthWord(trade_order, 9)) & dates(1)
      tickerSymbol = getNthWord(trade_order, 4)
      optionType = Left(getNthWord(trade_order, 12), 1)
      strikes = Split(getNthWord(trade_order, 11), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      optionSignature = "." & tickerSymbol & date1 & optionType & strike1 & "-." & tickerSymbol & date2 & optionType & strike2
    Case "Calendar"
      dates = Split(getNthWord(trade_order, 8), "/")
      date1 = dates(0) & GetMonth(getNthWord(trade_order, 7)) & getNthWord(trade_order, 6)
      date2 = getNthWord(trade_order, 10) & GetMonth(getNthWord(trade_order, 9)) & dates(1)
      tickerSymbol = getNthWord(trade_order, 4)
      signaturePostfix = Left(getNthWord(trade_order, 12), 1) & getNthWord(trade_order, 11)
      optionSignature = "." & tickerSymbol & date1 & signaturePostfix & "-." & tickerSymbol & date2 & signaturePostfix
    Case "Butterfly"
      signatureBase = getNthWord(trade_order, 4) & getNthWord(trade_order, 8) & GetMonth(getNthWord(trade_order, 7)) & getNthWord(trade_order, 6) & Left(getNthWord(trade_order, 10), 1)
      strikes = Split(getNthWord(trade_order, 9), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      strike3 = strikes(2)
      optionSignature = "." & signatureBase & strike1 & "-." & signatureBase & strike2 & "-." & signatureBase & strike2 & "+." & signatureBase & strike3
    Case "Combo"
      signatureBase = getNthWord(trade_order, 4) & getNthWord(trade_order, 8) & GetMonth(getNthWord(trade_order, 7)) & getNthWord(trade_order, 6)
      types = Split(getNthWord(trade_order, 10), "/")
      type1 = Left(types(0), 1)
      type2 = Left(types(1), 1)
      strike = getNthWord(trade_order, 9)
      optionSignature = "." & signatureBase & type1 & strike & "-." & signatureBase & type2 & strike
    Case Else
    End Select

    getOptionSignature = optionSignature
    Exit Function
ErrorHandl:
    getOptionSignature = ""
End Function