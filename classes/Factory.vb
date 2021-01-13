Enum StopLossFields
  CONDITIONAL_TRIGGER
  PRICE_LEVEL
  AMOUNT_TYPE
  R1_PERCENT_PROFIT
  R1_PERCENT_DAMAGES
  R1_PERCENT_SAVES
  R1_PERCENT_DIFFERENCE
  R1_PERCENT_EXIT_RULE
  R1_FIXED_PROFIT
  R1_FIXED_DAMAGES
  R1_FIXED_SAVES
  R1_FIXED_DIFFERENCE
  R1_FIXED_EXIT_RULE
  S2_PERCENT_PROFIT
  S2_PERCENT_DAMAGES
  S2_PERCENT_SAVES
  S2_PERCENT_DIFFERENCE
  S2_PERCENT_EXIT_RULE
  S2_FIXED_PROFIT
  S2_FIXED_DAMAGES
  S2_FIXED_SAVES
  S2_FIXED_DIFFERENCE
  S2_FIXED_EXIT_RULE
  MAX_LOSS_PERCENT_PROFIT
  MAX_LOSS_PERCENT_DAMAGES
  MAX_LOSS_PERCENT_SAVES
  MAX_LOSS_PERCENT_DIFFERENCE
  MAX_LOSS_PERCENT_EXIT_RULE
End Enum

Public LongCallVertical As StopRule
Public ShortPutVertical As StopRule
Public LongSynthetic As StopRule
Public LongCallDiagonal As StopRule
Public LongCall As StopRule
Public ShortPut As StopRule
Public LongPutVertical As StopRule
Public ShortCallVertical As StopRule
Public ShortSynthetic As StopRule
Public LongPutDiagonal As StopRule
Public LongPut As StopRule
Public ShortCall As StopRule
Public IronCondor As StopRule
Public Butterfly As StopRule
Public Calendar As StopRule

Public Function CreateStopRule(strategy As String) As StopRule
    Dim stop_rule As StopRule
    Set stop_rule = New StopRule
    Dim ExitRule As StopLossFields

    ConditionalTrigger = GetStopRule(strategy, CONDITIONAL_TRIGGER)
    PriceLevel = GetStopRule(strategy, PRICE_LEVEL)
    AmountType = GetStopRule(strategy, AMOUNT_TYPE)

    whichRule = IIf(PriceLevel = "Resistance", 0, 2) + IIf(AmountType = "Percent", 1, 2) + IIf(ConditionalTrigger = "Option", 4, 0)
    Select Case whichRule
      Case 1
        ExitRule = R1_PERCENT_EXIT_RULE
      Case 2
        ExitRule = R1_FIXED_EXIT_RULE
      Case 3
        ExitRule = S2_PERCENT_EXIT_RULE
      Case 4
        ExitRule = S2_FIXED_EXIT_RULE
      Case Else
        ExitRule = MAX_LOSS_PERCENT_EXIT_RULE
    End Select

    With stop_rule
      .strategy = strategy
      .ConditionalTrigger = ConditionalTrigger
      .PriceLevel = PriceLevel
      .AmountType = AmountType
      .ExitRule = GetStopRule(strategy, ExitRule)
    End With
    Set CreateStopRule = stop_rule
End Function

Function GetStopRule(pattern As String, field As StopLossFields)
  On Error GoTo ErrorHandl
  If pattern = "" Then GoTo ErrorHandl
  colOffset = field + 2
  foundValue = Application.WorksheetFunction.VLookup(pattern, [Strategic_Stop_Loss_Rules_Table], colOffset, False)
  GetStopRule = foundValue
  Exit Function
ErrorHandl:
    GetStopRule = ""
End Function

' TODO: Add button next to table, "Update Calculations"
Public Function Instantiate()
  If LongCallVertical Is Nothing Then
    Set LongCallVertical = CreateStopRule("Long Call Vertical")
    Set ShortPutVertical = CreateStopRule("Short Put Vertical")
    Set LongSynthetic = CreateStopRule("Long Synthetic")
    Set LongCallDiagonal = CreateStopRule("Long Call Diagonal")
    Set LongCall = CreateStopRule("Long Call")
    Set ShortPut = CreateStopRule("Short Put")
    Set LongPutVertical = CreateStopRule("Long Put Vertical")
    Set ShortCallVertical = CreateStopRule("Short Call Vertical")
    Set ShortSynthetic = CreateStopRule("Short Synthetic")
    Set LongPutDiagonal = CreateStopRule("Long Put Diagonal")
    Set LongPut = CreateStopRule("Long Put")
    Set ShortCall = CreateStopRule("Short Call")
    Set IronCondor = CreateStopRule("Iron Condor")
    Set Butterfly = CreateStopRule("Butterfly")
    Set Calendar = CreateStopRule("Calendar")
  End If
End Function
