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

Public SymmetricalTriangle As StopRule
Public PriceChannel As StopRule
Public Flag As StopRule
Public Pennant As StopRule
Public Breakout As StopRule
Public TriangleBreakout As StopRule
Public DoubleBottom As StopRule
Public TripleBottom As StopRule
Public HAndSBottom As StopRule
Public AscendingTriangle As StopRule
Public DoubleTop As StopRule
Public TripleTop As StopRule
Public HeadAndShoulders As StopRule
Public DescendingTriangle As StopRule

Private Function CreateStopRule(name As String, table As range) As StopRule
    Dim stop_rule As StopRule
    Set stop_rule = New StopRule
    Dim ExitRule As StopLossFields

    ConditionalTrigger = GetStopRule(name, CONDITIONAL_TRIGGER, table)
    PriceLevel = GetStopRule(name, PRICE_LEVEL, table)
    AmountType = GetStopRule(name, AMOUNT_TYPE, table)

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
      .name = name
      .ConditionalTrigger = ConditionalTrigger
      .PriceLevel = PriceLevel
      .AmountType = AmountType
      .ExitRule = GetStopRule(name, ExitRule, table)
    End With
    Set CreateStopRule = stop_rule
End Function

Private Function GetStopRule(strategy As String, field As StopLossFields, table As range)
  On Error GoTo ErrorHandl
  If strategy = "" Then GoTo ErrorHandl
  colOffset = field + 2
  foundValue = Application.WorksheetFunction.VLookup(strategy, table, colOffset, False)
  GetStopRule = foundValue
  Exit Function
ErrorHandl:
    GetStopRule = ""
End Function

Private Function CreateTechnicalStopRule(pattern As String) As StopRule
  Set CreateTechnicalStopRule = CreateStopRule(pattern, [Technical_Stop_Loss_Rules_Table])
End Function

Private Function CreateStrategicStopRule(strategy As String) As StopRule
  Set CreateStrategicStopRule = CreateStopRule(strategy, [Strategic_Stop_Loss_Rules_Table])
End Function

' TODO: Add button next to table, "Update Calculations"
Public Function InstantiateSingletons()
  If LongCallVertical Is Nothing Then
    Set LongCallVertical = CreateStrategicStopRule("Long Call Vertical")
    Set ShortPutVertical = CreateStrategicStopRule("Short Put Vertical")
    Set LongSynthetic = CreateStrategicStopRule("Long Synthetic")
    Set LongCallDiagonal = CreateStrategicStopRule("Long Call Diagonal")
    Set LongCall = CreateStrategicStopRule("Long Call")
    Set ShortPut = CreateStrategicStopRule("Short Put")
    Set LongPutVertical = CreateStrategicStopRule("Long Put Vertical")
    Set ShortCallVertical = CreateStrategicStopRule("Short Call Vertical")
    Set ShortSynthetic = CreateStrategicStopRule("Short Synthetic")
    Set LongPutDiagonal = CreateStrategicStopRule("Long Put Diagonal")
    Set LongPut = CreateStrategicStopRule("Long Put")
    Set ShortCall = CreateStrategicStopRule("Short Call")
    Set IronCondor = CreateStrategicStopRule("Iron Condor")
    Set Butterfly = CreateStrategicStopRule("Butterfly")
    Set Calendar = CreateStrategicStopRule("Calendar")
    Set SymmetricalTriangle = CreateTechnicalStopRule("Symmetrical Triangle")
    Set PriceChannel = CreateTechnicalStopRule("Price Channel")
    Set Flag = CreateTechnicalStopRule("Flag")
    Set Pennant = CreateTechnicalStopRule("Pennant")
    Set Breakout = CreateTechnicalStopRule("Breakout")
    Set TriangleBreakout = CreateTechnicalStopRule("Triangle Breakout")
    Set DoubleBottom = CreateTechnicalStopRule("Double Bottom")
    Set TripleBottom = CreateTechnicalStopRule("Triple Bottom")
    Set HAndSBottom = CreateTechnicalStopRule("H&S Bottom")
    Set AscendingTriangle = CreateTechnicalStopRule("Ascending Triangle")
    Set DoubleTop = CreateTechnicalStopRule("Double Top")
    Set TripleTop = CreateTechnicalStopRule("Triple Top")
    Set HeadAndShoulders = CreateTechnicalStopRule("Head and Shoulders")
    Set DescendingTriangle = CreateTechnicalStopRule("Descending Triangle")
  End If
End Function