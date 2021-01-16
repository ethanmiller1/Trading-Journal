Enum StopLossRule 
  FIXED_RESISTANCE
  FIXED_SUPPORT
  FIXED_PROTECTION
  PERCENT_RESISTANCE
  PERCENT_SUPPORT
  PERCENT_MAX_LOSS
End Enum

Function GetOptionType(trade_order As String)
    'If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl

    'Create array to hold option types.
    Dim optionTypes As Variant
    Dim optionType As Variant 

    'Populate array with option types.
    optionTypes = Array("IRON CONDOR","BUTTERFLY","CALENDAR","DIAGONAL","COMBO","VERTICAL","CALL","PUT")

    'Search each option type for a match in the trade_order string that indicates the option type.
    For Each optionType In optionTypes
        'InStr returns the position of the matching keyword in the string. If the return is 0, no match was found.
        If InStr(trade_order, optionType) Then
            oType = optionType
            Exit For
        End If
    Next

    'If there are no matches in trade_order, return null.
    If oType = 0 Then GoTo ErrorHandl

    'Convert string to sentence case.
    oType = WorksheetFunction.Proper(oType)

    'Return option type.
    GetOptionType = oType
    Exit Function
ErrorHandl:
    GetOptionType = ""
End Function

Function GetNthWord(text As String, start_num As Integer, Optional num_words As Integer=1)
Dim current_pos As Long
Dim char_len As Integer
Dim current_word_num As Integer
 
GetNthWord = ""
current_word_num = 1
 
'Remove leading spaces.
text = Trim(text)

'Get the number of characters in the text.
char_len = Len(text)

'Find the character position of the nth word.
For current_pos = 1 To char_len
    'If this is the nth word, concatinate each character until it is no longer the nth + num_words word.
    If (current_word_num >= start_num) And (current_word_num <= start_num + num_words - 1) Then
        GetNthWord = GetNthWord & Mid(text, current_pos, 1)
    End If

    'If there is a space after this character, increment the current_word_num by 1.
    If (Mid(text, current_pos, 1) = " ") Then
      current_word_num = current_word_num + 1
    End If
Next current_pos
 
'Remove the rightmost space.
GetNthWord = Trim(GetNthWord)
 
End Function

Function GetExpiration(trade_order As String, Optional option_type As String = "false")
    If option_type = "false" Then option_type = GetOptionType(trade_order)
    Select Case option_type
    'If option type matches case strings, return the 6th-8th word.
    Case "Combo", "Vertical", "Butterfly"
      expDate = GetNthWord(trade_order, 6, 3)
    'If option type matches case strings, parse the 8th word and return it with the 9th and 10th word concatenated.
    Case "Diagonal", "Calendar"
      expDay = GetNthWord(trade_order, 8)
      expDay = Split(expDay, "/")
      expDate = expDay(1) & " " & GetNthWord(trade_order, 9, 2)
    'If option type matches case strings, return the 5th-7th word.
    Case "Call", "Put"
      expDate = GetNthWord(trade_order,5,3)
    'If option type is an Iron Condor, return the 7th-9th word.
    Case "Iron Condor"
      expDate = GetNthWord(trade_order,7,3)
    'If there is no match, return null.
    Case Else
      GoTo ErrorHandl
    End Select

    'Return the concatonated string as a date value.
    GetExpiration = DateValue(expDate)
    Exit Function
ErrorHandl:
    GetExpiration = ""
End Function

Function GetSymbol(trade_order As String, option_type As String)
    'If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    
    Select Case option_type
    'If option type is an Iron Condor, return the 5th word.
    Case "Iron Condor"
      Symbol = GetNthWord(trade_order, 5)
    'Otherwise return the 4th word.
    Case "Call", "Put"
      Symbol = GetNthWord(trade_order, 3)
    Case Else
      Symbol = GetNthWord(trade_order, 4)
    End Select

    'Return the concatonated string as a date value.
    GetSymbol = Symbol
    Exit Function
ErrorHandl:
    GetSymbol = ""
End Function

Function DaysTillExp(trade_date As Date, expiration_date As Date)
    'If argument is null, return null.
    'TODO: Handle #VALUE error from wrong datatype.
    If trade_date = 0 Then GoTo ErrorHandl

    'Return the concatonated string as a date value.
    DaysTillExp = expiration_date - trade_date
    Exit Function
ErrorHandl:
    DaysTillExp = ""
End Function

Function GetStrategy(trade_order As String, Optional option_type As String = "false")
    'If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    If option_type = "false" Then option_type = GetOptionType(trade_order)

      'Test market position.
      If InStr(trade_order, "BOT") Then
        position = "Long"
      Else
        position = "Short"
      End If

      'Test option type.
      If InStr(trade_order, "CALL") Then
        side = "Call"
      ElseIf InStr(trade_order, "PUT") Then
        side = "Put"
      End If

    Select Case option_type
    Case "Vertical", "Diagonal"
     'Concatenate market position, option side, and option type to determine strategy.
      strategy = position & " " & side &  " " & option_type
    Case "Call", "Put"
      strategy = position & " " & option_type
    Case "Combo"
      strategy = position & " Synthetic"
    Case Else
      'For Iron Condor, Butterfly, and Calendar.
      strategy = option_type
    End Select

    'Return the concatonated string as a date value.
    GetStrategy = strategy
    Exit Function
ErrorHandl:
    GetStrategy = ""
End Function

Function GetPosture(option_strategy As String)
    'If argument is null, return null.
    If option_strategy = "" Then GoTo ErrorHandl

    Select Case option_strategy
    'Match option strategy to its corresponding market posture.
    Case "Iron Condor", "Calendar", "Butterfly"
      posture = "Neutral"
    Case "Long Call", "Short Put", "Long Call Vertical", "Short Put Vertical", "Long Call Diagonal", "Long Synthetic"
      posture = "Bullish"
    Case "Short Call", "Long Put", "Long Put Diagonal", "Short Synthetic"
      posture = "Bearish"
    End Select

    'Return the posture of the trade.
    GetPosture = posture
    Exit Function
ErrorHandl:
    GetPosture = ""
End Function

Function GetStockQuote(symobol As String, trade_date As Date)
    'If argument is null, return null.
    If symobol = "" Then GoTo ErrorHandl
    
    'Get date as YYYYMMDD.
    dateAsString = Replace(Format(trade_date, "yyyy/mm/dd"), "/", "")

    'Define iex's defaul strings.
    site = "https://sandbox.iexapis.com/stable/stock/"
    Chart = "/chart/"
    byDate = "/chart/date/"
    chartByDate = "?chartByDay=true"
    token = "?token=Tpk_46cc7266d1374035bb1cb11d639476c4"

    'Create array to hold acceptable range parameters.
    Dim ranges As Variant
    Dim range As Variant

    'Populate array with range values.
    'TODO: Test to see if date falls within range, and use select it accordingly.
    ranges = Array("5d", "1m", "3m", "6m", "1y", "2y", "5y", "max")

    'Concatonate strings to formulate url.
    'TODO: Troubleshoot chartByDate booleon failure and use it instead of a range (see documentation).
    Url = site & symobol & Chart & ranges(5) & token

    'Get json object from url.
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", Url, False
    objHTTP.send
    
    'Format date to search the JSON by.
    dateHyphenated = Format(trade_date, "yyyy-mm-dd")
    regPattern = "(({""date"":"")" & dateHyphenated & "[^}]*})"

    'Use regex to isolate single stock quote.
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = regPattern: regex.Global = False
    Set matches = regex.Execute(objHTTP.responseText)
    stockQuote = matches(0)

    'TODO: hande #VALUE error due to no regex matches.

    'Return the stock quote.
    GetStockQuote = stockQuote
    Exit Function
ErrorHandl:
    GetStockQuote = ""
End Function

Function GetQuoteValue(quote_key As String, stock_quote As String)
    'If argument is null, return null.
    If stock_quote = "" Then GoTo ErrorHandl

    'Concatonate desired JSON key with regex commands to trap the value in a capturing group.
    regPattern = """" & quote_key & """:([^,]*)"

    'Use regex to return the value of the key passed in.
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = regPattern: regex.Global = False
    Set matches = regex.Execute(stock_quote)
    quoteValue = matches(0).SubMatches(0)

    GetQuoteValue = quoteValue
    Exit Function
ErrorHandl:
    GetQuoteValue = ""
End Function

Function GetPrem(trade_order As String, Optional option_type As String = "false")
    'If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    If option_type = "false" Then option_type = GetOptionType(trade_order)
    
    'Get the nth word based on option type.
    Select Case option_type
    Case "Iron Condor"
      premium = GetNthWord(trade_order, 12)
    Case "Calendar", "Diagonal"
      premium = GetNthWord(trade_order, 13)
    Case "Call", "Put"
      premium = GetNthWord(trade_order, 10)
    Case Else
      premium = GetNthWord(trade_order, 11)
    End Select

    'Remove @ symbol and add 0 to cast as numeric.
    premium = Replace(premium,"@","") + 0

    'Return the option premium.
    GetPrem = premium
    Exit Function
ErrorHandl:
    GetPrem = ""
End Function

Function GetMaxProfit(trade_order As String, option_type As String, qty As String, prem As String, Optional comm As Currency = 0#, Optional risk As Currency = 0#)
    'If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    
    'Convert strings to numbers. (Currency and Integer won't accept "" as an argument, which results in a #VALUE error.)
    Dim cPrem As Currency
    cPrem = CCur(prem)
    If comm = 0 Then comm = GetCommission(trade_order)
    
    'Get the nth word based on option type.
    Select Case option_type
    Case "Iron Condor"
      maxProfit = cPrem
    Case "Butterfly"
      'Get strikes.
      strikes = Split(GetNthWord(trade_order, 9), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      strike3 = strikes(2)

      'Get the distance between the short strike and long strike.
      spread1 = Abs(strike1 - strike2)
      spread2 = Abs(strike3 - strike2)

      'Whichever spread is smaller is the max profit.
      If spread1 <= spread2 Then
        maxProfit = spread1 - cPrem
      Else
        maxProfit = spread2 - cPrem
      End If
    Case "Calendar"
      maxProfit = risk * 2
      'TODO: GetMaxProfit = maxProfit, Exit Function
    Case "Diagonal"
      'If debit paid was negitive (actually a credit).
      If cPrem < 0 Then
        maxProfit = Abs(cPrem)
      Else
        'Get strikes.
        strikes = Split(GetNthWord(trade_order, 11), "/")
        strike1 = strikes(0)
        strike2 = strikes(1)

        'Spread Width - Premium + Extrinsic Value.
        'TODO: Use Black-scholes Model to calculate extrinsic value.
        maxProfit = Abs(strike1 - strike2) - cPrem
      End If
    Case "Combo"
        'If long, stock has unlimited profit potential.
        If InStr(trade_order, "BOT") Then
          GetMaxProfit = "Undefined"
          Exit Function
        'If short, stock can only drop to 0 (finite profits).
        Else
          maxProfit = GetNthWord(trade_order, 9) + cPrem
        End If
    Case "Call"
        'If long, call has unlimited profit potential.
        If InStr(trade_order, "BOT") Then
          GetMaxProfit = "Undefined"
          Exit Function
        'If short, call can only expire worthless.
        Else
          maxProfit = cPrem
        End If
    Case "Put"
      'If long, put has unlimited profit potential.
        If InStr(trade_order, "BOT") Then
          maxProfit = GetNthWord(trade_order, 8) - cPrem
        'If short, put can only expire worthless.
        Else
          maxProfit = cPrem
        End If
    'Case "Vertical"
    Case Else
      'If short, max profit is credit recieved.
      If InStr(trade_order, "SOLD") Then
        maxProfit = cPrem
      'If short, call can only expire worthless.
      Else
        'Get strikes.
        strikes = Split(GetNthWord(trade_order, 9), "/")
        strike1 = strikes(0)
        strike2 = strikes(1)

        maxProfit = Abs(strike1 - strike2) - cPrem
      End If
    End Select
    
    'Multiply by shares being controlled and subtract commissions.
    maxProfit = maxProfit * Abs(qty) * 100 - comm

    'Return the option maxProfit.
    GetMaxProfit = maxProfit
    Exit Function
ErrorHandl:
    GetMaxProfit = ""
End Function

Function GetRisk(trade_order As String, option_type As String, qty As String, prem As String, max_profit As String, Optional comm As Currency = 0)
    'If argument is null, return null.
    If trade_order = "" Then GoTo ErrorHandl
    
    'TODO: Replace arguments with functions.
    'qty = GetNthWord(trade_order,2)
    'prem = GetPrem(trade_order, option_type)

    'Covert strings to numbers. (Currency and Integer won't accept "" as an argument, which results in a #VALUE error.)
    Dim premium As Currency
    premium = CCur(prem)
    Dim maxProfit As Currency
    If IsNumeric(max_profit) Then maxProfit = CCur(max_profit)
    If comm = 0 Then comm = GetCommission(trade_order)
    
    'Get the nth word based on option type.
    Select Case option_type
    Case "Iron Condor"
      strikes = Split(GetNthWord(trade_order, 10), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      strike3 = strikes(2)
      strike4 = strikes(3)

      'Get the distance between the short strike and long strike.
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
      'If short, risk is spread width less credit recieved.
      If InStr(trade_order, "SOLD") Then
        'Get strikes.
        strikes = Split(GetNthWord(trade_order, 9), "/")
        strike1 = strikes(0)
        strike2 = strikes(1)

        risk = Abs(strike1 - strike2) - premium
      'If long, risk is debit paid.
      Else
        risk = premium
      End If
    Case "Diagonal"
      If premium < 0 Then
        'Get strikes.
        strikes = Split(GetNthWord(trade_order, 11), "/")
        strike1 = strikes(0)
        strike2 = strikes(1)
        spread = Abs(strike1 - strike2)

        risk = spread - Abs(premium)
      Else
        risk = premium
      End If
    Case "Combo"
      'If short, risk is unlimited.
      If InStr(trade_order, "SOLD") Then
        GetRisk = "Undefined"
        Exit Function
      'If long, risk is capped by the floor.
      Else
        strike1 = GetNthWord(trade_order, 9)
        risk = strike1 - premium
      End If
    Case "Call"
      'If short, risk is unlimited.
      If InStr(trade_order, "SOLD") Then
        GetRisk = "Undefined"
        Exit Function
      'If long, risk is debit paid.
      Else
        risk = premium
      End If
    Case "Put"
      'If short, risk is capped by the floor.
      If InStr(trade_order, "SOLD") Then
        strike1 = GetNthWord(trade_order, 9)
        risk = strike1 - premium
      'If long, risk is debit paid.
      Else
        risk = premium
      End If
    Case Else
      risk = premium
    End Select
    
    'Multiply by shares being controlled and add commissions.
    risk = risk * Abs(qty) * 100 + comm

    'Return the option risk.
    GetRisk = risk
    Exit Function
IronCondor:
    risk = risk * Abs(qty) * 100 + comm
    GetRisk = risk - maxProfit
    Exit Function
Standard:
    risk = premium
    risk = risk * Abs(qty) * 100 + comm
    GetRisk = risk
ErrorHandl:
    GetRisk = ""
End Function

Function GetPLClose(trade_order As String, option_type As String, prem As String, max_profit As String, Optional comm As Currency = 0#)
    'If argument is null, return null.
    If trade_order = "" Or prem = "" Then GoTo ErrorHandl
    
    'Covert strings to numbers.
    Dim premium As Currency
    Dim maxProfit As Currency
    premium = CCur(prem)
    If IsNumeric(max_profit) Then maxProfit = CCur(max_profit)
    If comm = 0 Then comm = GetCommission(trade_order)

    'How many shares being controlled?
    contracts = Abs(GetNthWord(trade_order,2))
    shares = contracts * 100
    debit = premium * shares

    'Get the nth word based on option type.
    Select Case option_type
    Case "Iron Condor"
      'Opening Credit - Closing Debit
      plClose = maxProfit - debit
      'Skip comissions -- Max profits accounts for it.
    Case "Combo"
      credit = Replace(GetNthWord(trade_order, 11),"@","") + 0
      plClose = debit + (credit*shares) - comm
    Case "Vertical"
      If InStr(trade_order, "SOLD") Then
        plClose = maxProfit - debit
      Else
        entryDebit = GetPrem(trade_order, option_type)
        risk = GetRisk(trade_order, option_type, CStr(contracts), CStr(entryDebit), max_profit, comm)
        credit = debit
        plClose = credit - risk
      End If
    Case Else
      credit = GetPrem(trade_order, option_type)
      'If credit was recieved...
      If (InStr(trade_order, "BOT") And credit < 0) Or InStr(trade_order, "SOLD") Then
        plClose = maxProfit - debit
      Else
        risk = GetRisk(trade_order, option_type, CStr(contracts), CStr(credit), max_profit, comm)
        credit = debit
        plClose = credit - risk
      End If
    End Select

    'Return the option plClose.
    GetPLClose = plClose
    Exit Function
ErrorHandl:
    GetPLClose = ""
End Function

Function GetPLPercent(pl_closed As String, max_profit As String, max_risk As String)
 
    'If argument is null, return null.
    If pl_closed = "" Then GoTo ErrorHandl

    'Covert strings to numbers.
    Dim plClose As Currency
    plClose = CCur(pl_closed)
    Dim maxProfit As Currency
    If IsNumeric(max_profit) Then maxProfit = CCur(max_profit) Else GoTo Undefined
    Dim risk As Currency
    If IsNumeric(max_risk) Then risk = CCur(max_risk) Else GoTo Undefined

    'If trade was profitable, show percent of max profit
    If plClose > 0 Then
      plPercent = plClose / maxProfit
    'If trade lost money, show percent of max loss
    Else
      plPercent = plClose / risk
    End If

    'Return the option plPercent.
    GetPLPercent = plPercent
    Exit Function
Undefined:
    GetPLPercent = 0
    Exit Function
ErrorHandl:
    GetPLPercent = ""
End Function

Function GetMonth(mon_abr As String)
  monthNumber = Application.Match(mon_abr, Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"), 0)
  GetMonth = Format(monthNumber, "00")
End Function

Function GetOptionSignature(trade_order As String)
    If trade_order = "" Then GoTo ErrorHandl

    option_type = GetOptionType(trade_order)

    Select Case option_type
    'Single Leg Options (AND P/L%)
    Case "Call", "Put"
      optionSignature = "." & GetNthWord(trade_order, 3) & GetNthWord(trade_order, 7) & GetMonth(GetNthWord(trade_order, 6)) & GetNthWord(trade_order, 5) & Left(GetNthWord(trade_order, 9), 1) & GetNthWord(trade_order, 8)
    Case "Iron Condor"
      signatureBase = GetNthWord(trade_order, 5) & GetNthWord(trade_order, 9) & GetMonth(GetNthWord(trade_order, 8)) & GetNthWord(trade_order, 7)
      types = Split(GetNthWord(trade_order, 11), "/")
      type1 = Left(types(0), 1)
      type2 = Left(types(1), 1)
      strikes = Split(GetNthWord(trade_order, 10), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      strike3 = strikes(2)
      strike4 = strikes(3)
      optionSignature = "." & signatureBase & type1 & strike1 & "-." & signatureBase & type1 & strike2 & "+." & signatureBase & type2 & strike3 & "-." & signatureBase & type2 & strike4
    Case "Vertical"
      signatureBase = GetNthWord(trade_order, 4) & GetNthWord(trade_order, 8) & GetMonth(GetNthWord(trade_order, 7)) & GetNthWord(trade_order, 6) & Left(GetNthWord(trade_order, 10), 1)
      strikes = Split(GetNthWord(trade_order, 9), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      optionSignature = "." & signatureBase & strike1 & "-." & signatureBase & strike2
    Case "Diagonal"
      dates = Split(GetNthWord(trade_order, 8), "/")
      date1 = dates(0) & GetMonth(GetNthWord(trade_order, 7)) & GetNthWord(trade_order, 6)
      date2 = GetNthWord(trade_order, 10) & GetMonth(GetNthWord(trade_order, 9)) & dates(1)
      tickerSymbol = GetNthWord(trade_order, 4)
      optionType = Left(GetNthWord(trade_order, 12), 1)
      strikes = Split(GetNthWord(trade_order, 11), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      optionSignature = "." & tickerSymbol & date1 & optionType & strike1 & "-." & tickerSymbol & date2 & optionType & strike2
    Case "Calendar"
      dates = Split(GetNthWord(trade_order, 8), "/")
      date1 = dates(0) & GetMonth(GetNthWord(trade_order, 7)) & GetNthWord(trade_order, 6)
      date2 = GetNthWord(trade_order, 10) & GetMonth(GetNthWord(trade_order, 9)) & dates(1)
      tickerSymbol = GetNthWord(trade_order, 4)
      signaturePostfix = Left(GetNthWord(trade_order, 12), 1) & GetNthWord(trade_order, 11)
      optionSignature = "." & tickerSymbol & date1 & signaturePostfix & "-." & tickerSymbol & date2 & signaturePostfix
    Case "Butterfly"
      signatureBase = GetNthWord(trade_order, 4) & GetNthWord(trade_order, 8) & GetMonth(GetNthWord(trade_order, 7)) & GetNthWord(trade_order, 6) & Left(GetNthWord(trade_order, 10), 1)
      strikes = Split(GetNthWord(trade_order, 9), "/")
      strike1 = strikes(0)
      strike2 = strikes(1)
      strike3 = strikes(2)
      optionSignature = "." & signatureBase & strike1 & "-." & signatureBase & strike2 & "-." & signatureBase & strike2 & "+." & signatureBase & strike3
    Case "Combo"
      signatureBase = GetNthWord(trade_order, 4) & GetNthWord(trade_order, 8) & GetMonth(GetNthWord(trade_order, 7)) & GetNthWord(trade_order, 6)
      types = Split(GetNthWord(trade_order, 10), "/")
      type1 = Left(types(0), 1)
      type2 = Left(types(1), 1)
      strike = GetNthWord(trade_order, 9)
      optionSignature = "." & signatureBase & type1 & strike & "-." & signatureBase & type2 & strike
    Case Else
    End Select
    GetOptionSignature = optionSignature
    Exit Function
ErrorHandl:
    GetOptionSignature = ""
End Function

Function GetCommission(trade_order As String)
    If trade_order = "" Then GoTo ErrorHandl

    option_strategy = GetOptionType(trade_order)
    Dim qty As Integer
    Dim noOfContracts As Integer
    Dim closedQty As Integer
    Dim BASE_FEE As Currency
    Dim FEE_PER_CONTRACT As Currency

    qty = GetNthWord(trade_order, 2)
    closedQty = 2
    BASE_FEE = 1.25
    FEE_PER_CONTRACT = 0.75

    Select Case option_strategy
    Case "Call", "Put"
      noOfContracts = 1
    Case "Vertical", "Diagonal", "Calendar", "Combo"
      noOfContracts = 2
    Case "Iron Condor", "Butterfly"
      noOfContracts = 4
    Case Else
    End Select

    GetCommission = ( BASE_FEE + (Abs(qty) * FEE_PER_CONTRACT * noOfContracts ) ) * closedQty
    Exit Function
ErrorHandl:
    GetCommission = ""
End Function

Function Clipboard(Optional StoreText As String) As String
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

  Dim x As Variant

  'Store as variant for 64-bit VBA support
    x = StoreText

  'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", x
        Case Else
          'Read from the clipboard (no variable passed through)
            Clipboard = .GetData("text")
      End Select
    End With
  End With
End Function

Function GetOptimalDTE(trade_order As String, optimal_exit_date As Date)
    If trade_order = "" Or optimal_exit_date = 0 Then GoTo ErrorHandl
    Dim expiration As Date
    expiration = DateValue(GetExpiration(trade_order))
    GetOptimalDTE = expiration - optimal_exit_date
    Exit Function
ErrorHandl:
    GetOptimalDTE = ""
End Function

Function GetActualMaxProfit(trade_order As String, optimal_prem As String)
    If trade_order = "" Or optimal_prem = "" Then GoTo ErrorHandl

    Dim option_strategy As String
    Dim qty As String
    Dim prem As String
    Dim max_profit As String
    Dim optimal_cPrem As Currency

    optimal_cPrem = CCur(optimal_prem)
    option_strategy = GetOptionType(trade_order)
    qty = Abs(GetNthWord(trade_order, 2))
    prem = GetPrem(trade_order)
    max_profit = GetMaxProfit(trade_order, option_strategy, qty, prem)
    isShort = GetNthWord(trade_order, 1) = "SOLD"

    Select Case option_strategy
    Case "Iron Condor"
      maxProfit = max_profit - (optimal_cPrem * qty * 100)
    Case "Combo"
      maxProfit = (optimal_cPrem * 100 * qty) + (prem * 100 * qty) - GetCommission(trade_order)
    Case "Vertical"
      If isShort Then
        maxProfit = max_profit - (optimal_cPrem * 100 * qty)
      Else
        risk = GetRisk(trade_order, option_strategy, qty, prem, max_profit)
        maxProfit = optimal_cPrem * 100 * qty - risk
      End If
    Case Else
      If (Not isShort And prem < 0) Or isShort Then
        maxProfit = max_profit - optimal_cPrem * 100 * qty
      Else
        risk = GetRisk(trade_order, option_strategy, qty, prem, max_profit)
        maxProfit = optimal_cPrem * 100 * qty - risk
      End If
    End Select

    GetActualMaxProfit = maxProfit
    Exit Function
ErrorHandl:
    GetActualMaxProfit = ""
End Function

Function GetPercentOfMaxProfit(max_profit As String, risk As String, actual_max_profit As String)
    If max_profit = "" Or actual_max_profit = "" Then GoTo ErrorHandl
    If max_profit = "Undefined" Or Risk = "Undefined" Then GoTo Undefined

    precentOfMaxProfit = actual_max_profit / IIf(actual_max_profit > 0, max_profit, risk)

    GetPercentOfMaxProfit = precentOfMaxProfit
    Exit Function
Undefined:
    precentOfMaxProfit = 0
    GetPercentOfMaxProfit = precentOfMaxProfit
    Exit Function   
ErrorHandl:
    GetPercentOfMaxProfit = ""
End Function

Function GetTarget(support As String, resistance As String, entry As String, percent As Double)
    If support = "" Or resistance = "" Then GoTo ErrorHandl

    operator = IIf(support<resistance, 1, -1)

    If entry <> "" Then
      target = entry + operator * ABS(resistance - support) * (percent - 1)
    Else
      target = resistance + operator * ABS(resistance - support) * (percent - 1)
    End If

    GetTarget = target
    Exit Function 
ErrorHandl:
    GetTarget = ""
End Function

Function GetTarget1(support As String, resistance As String, entry As String)
    GetTarget1 = GetTarget(support, resistance, entry, 1.618)
End Function

Function GetTarget2(support As String, resistance As String, entry As String)
    GetTarget2 = GetTarget(support, resistance, entry, 2)
End Function

Private Function GetStopLossRule(ByVal rule As StopLossRule)
  Select Case rule
    Case StopLossRule.FIXED_RESISTANCE
      stopLoss = .20
    Case StopLossRule.FIXED_SUPPORT
      stopLoss = .20
    Case StopLossRule.FIXED_PROTECTION
      stopLoss = .20
    Case StopLossRule.PERCENT_RESISTANCE
      stopLoss = .01
    Case StopLossRule.PERCENT_SUPPORT
      stopLoss = .01
    Case StopLossRule.PERCENT_MAX_LOSS
      stopLoss = .50
    Case Else
      stopLoss = 0
  End Select
  GetStopLossRule = stopLoss
End Function

Function GetTDADefaultStopLoss(pattern As String, support As Currency, resistance As Currency, stopReference As String, entryReference As Currency, returnOnRisk As String, target1 As Currency, target2 As Currency, latestSupport As String)
    If latestSupport = "" Then GoTo ErrorHandl

    'Convert strings to currency
    Dim curStopReference As Currency
    If IsNumeric(stopReference) Then curStopReference = CCur(stopReference)

    postureAdjustment = IIf(IsBullish(support, resistance), 1, -1)

    'If stopReferene doesn't exist, use resistance as the stopReference
    If stopReference <> "" Then
      'Add fixed adjustment to the stop reference
      stopLoss = curStopReference - GetStopLossRule(FIXED_SUPPORT) * postureAdjustment
    Else
      'Use 50% of the option's max loss as your stop (max loss is a negative number)
      'TODO: What in the confusing hypo's belly was I thinking when I came up with this...???
      stopLevel = IIf(entryReference <> 0, entryReference, resistance)
      optionMaxLoss = GetOptionMaxLoss(pattern, resistance, entryReference, returnOnRisk, target1, target2)
      stopLoss = stopLevel + optionMaxLoss * GetStopLossRule(PERCENT_MAX_LOSS) * postureAdjustment
    End If

    GetTDADefaultStopLoss = stopLoss
    Exit Function
ErrorHandl:
    GetTDADefaultStopLoss = ""
End Function

Function IsBullish(support As Currency, resistance As Currency)
'Determines the market posture (note this is not the posture of your option)
    IsBullish = support < resistance
End Function

Function GetOptionMaxLoss(pattern As String, resistance As Currency, entry As Currency, returnOnRisk As String, target1 As Currency, target2 As Currency)
    If pattern = "" Then GoTo ErrorHandl

    'TODO: If scaling, use max loss calculated above.

    'Convert strings to currency
    Dim curReturnOnRisk As Currency
    If IsNumeric(returnOnRisk) Then curReturnOnRisk = CCur(returnOnRisk)

    'TODO: Get from averaging recorded ROR values from Trading Journal
    typicalROR = 0.8

    ror = IIf(curReturnOnRisk <> 0, curReturnOnRisk, typicalROR)
    maxLoss = -1 / ror * GetTargetProfit(resistance, entry, target1, target2)

    GetOptionMaxLoss = maxLoss
    Exit Function
ErrorHandl:
    GetOptionMaxLoss = ""
End Function

Function GetTargetProfit(resistance As Currency, entry As Currency, target1 As Currency, target2 As Currency)
    If target1 = 0 And target2 = 0 Then GoTo ErrorHandl

    'TODO: Which Target based on Pattern Target Realiability data
    whichTarget = "T2"

    'TODO: This is the max stock movement. Multiply this number by delta to get true max gain.
    maxGain = Abs(IIf(whichTarget = "T1", target1, target2) - IIf(entry <> 0, entry, resistance))

    GetTargetProfit = maxGain
    Exit Function
ErrorHandl:
    GetTargetProfit = ""
End Function

Private Function GetTechnicalPriceRule(pattern As String) As StopRule
'TODO: Learn if the workbook has to recalculate all the cells that depend on these singletons every open
    If pattern = "" Then GoTo ErrorHandl
    Dim rule As StopRule

    'TODO: Remove after done editing workbook.
    Call InstantiateSingletons

    Select Case pattern
      Case "Symmetrical Triangle"
        Set rule = SymmetricalTriangle
      Case "Price Channel"
        Set rule = PriceChannel
      Case "Flag"
        Set rule = Flag
      Case "Pennant"
        Set rule = Pennant
      Case "Breakout"
        Set rule = Breakout
      Case "Triangle Breakout"
        Set rule = TriangleBreakout
      Case "Double Bottom"
        Set rule = DoubleBottom
      Case "Triple Bottom"
        Set rule = TripleBottom
      Case "H&S Bottom"
        Set rule = HAndSBottom
      Case "Ascending Triangle"
        Set rule = AscendingTriangle
      Case "Double Top"
        Set rule = DoubleTop
      Case "Triple Top"
        Set rule = TripleTop
      Case "Head and Shoulders"
        Set rule = HeadAndShoulders
      Case "Descending Triangle"
        Set rule = DescendingTriangle
    End Select
    Set GetTechnicalPriceRule = rule
    Exit Function
ErrorHandl:
    Set GetTechnicalPriceRule = Nothing
End Function

Private Function GetOptionPriceRule(strategy As String) As StopRule
    If strategy = "" Then GoTo ErrorHandl
    Dim rule As StopRule

    'TODO: Remove after done editing workbook.
    Call InstantiateSingletons

    Select Case strategy
      Case "Long Call Vertical"
        Set rule = LongCallVertical
      Case "Short Put Vertical"
        Set rule = ShortPutVertical
      Case "Long Synthetic"
        Set rule = LongSynthetic
      Case "Long Call Diagonal"
        Set rule = LongCallDiagonal
      Case "Long Call"
        Set rule = LongCall
      Case "Short Put"
        Set rule = ShortPut
      Case "Long Put Vertical"
        Set rule = LongPutVertical
      Case "Short Call Vertical"
        Set rule = ShortCallVertical
      Case "Short Synthetic"
        Set rule = ShortSynthetic
      Case "Long Put Diagonal"
        Set rule = LongPutDiagonal
      Case "Long Put"
        Set rule = LongPut
      Case "Short Call"
        Set rule = ShortCall
      Case "Iron Condor"
        Set rule = IronCondor
      Case "Butterfly"
        Set rule = Butterfly
      Case "Calendar"
        Set rule = Calendar
    End Select
    Set GetOptionPriceRule = rule
    Exit Function
ErrorHandl:
    Set GetOptionPriceRule = Nothing
End Function

Function GetMarketStop(pattern As String, support As Currency, resistance As Currency, stop_reference As String)
    If pattern = "" Then GoTo ErrorHandl

    Dim technicalStopRule as StopRule
    Set technicalStopRule = GetTechnicalPriceRule(pattern)
    postureAdjustment = IIf(IsBullish(support, resistance), 1, -1)
    Dim priceLevel As Currency
    Select Case technicalStopRule.PriceLevel
      Case "Support"
        priceLevel = support
      Case "Resistance"
        priceLevel = resistance
      Case "Candlestick Low"
        If Not IsNumeric(stop_reference) Then GoTo ErrorHandl
        priceLevel = CCur(stop_reference)
    End Select
    fixedOrPercent = IIf(technicalStopRule.AmountType = "Percent", priceLevel, 1)

    GetMarketStop = priceLevel - fixedOrPercent * technicalStopRule.ExitRule * postureAdjustment
    Exit Function
ErrorHandl:
    GetMarketStop = ""
End Function

Function GetOptionStop(trade_order As String, premium As String, risk As String)
    If trade_order = "" Then GoTo ErrorHandl

    Dim optionStopRule As StopRule
    Set optionStopRule = GetOptionPriceRule(GetStrategy(trade_order))
    prem = CDbl(premium) 'Consider using Abs()

    If optionStopRule.priceLevel = "Max Risk" Then
      'We don't want to consider comissions in our max stop
      qty = GetNthWord(trade_order, 2)
      rsk = CDbl(risk)
      optionStop = (rsk - GetCommission(trade_order)) / qty / 100 * optionStopRule.ExitRule + prem
    Else
      If optionStopRule.AmountType = "Percent" Then
        optionStop = prem * optionStopRule.ExitRule
      Else
        optionStop = prem + optionStopRule.ExitRule
      End If
    End If

    GetOptionStop = optionStop
    Exit Function
ErrorHandl:
    GetOptionStop = ""
End Function