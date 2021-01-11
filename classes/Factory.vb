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

Public Function CreateStopRule(Strategy As String, ConditionalTrigger As String, PriceLevel As String, AmountType As String, ExitRule As StopLossRule) As StopRule
    Dim stop_rule As StopRule
    Set stop_rule = New StopRule
    With stop_rule
      .Strategy = Strategy
      .ConditionalTrigger = ConditionalTrigger
      .PriceLevel = PriceLevel
      .AmountType = AmountType
      .ExitRule = ExitRule
    End With
    Set CreateStopRule = stop_rule
End Function

Function GetFixedStop2Rule(pattern As String)
  On Error GoTo ErrorHandl
  If pattern = "" Then GoTo ErrorHandl
  colOffset = StopLossFields.S2_FIXED_EXIT_RULE + 2
  foundValue = Application.WorksheetFunction.VLookup(pattern, [Strategic_Stop_Loss_Rules_Table], colOffset, False)
  GetFixedStop2Rule = foundValue
  Exit Function
ErrorHandl:
    GetFixedStop2Rule = ""
End Function


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

' TODO: Add button next to table, "Update Calculations"

Public Function Instantiate(ByRef className As String) As Object
  Select className
    Case "Dog"
      If LongCallVertical Is Nothing then
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Debit", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Credit", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Market", "Option", "Fixed", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Debit", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Bid", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Market", "Ask", "Fixed", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Debit", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Market", "Option", "Fixed", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Market", "Option", "Fixed", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Debit", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Bid", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Market", "Ask", "Fixed", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Max Risk", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Debit", "Percent", GetStopLossRule(FIXED_SUPPORT))
        Set LongCallVertical = CreateStopRule("LongCallVertical", "Option", "Debit", "Percent", GetStopLossRule(FIXED_SUPPORT))


      Else
        Err.Raise 1 + vbObjectError, "Factory.Instantiate", "Instantiation failed. There already exists an instance of " & className
      End If
  End Select

  Set Instantiate = New Dog
End Function

' Singleton pattern
Public Function GetShared() As myClass

    If objSharedClass Is Nothing Then
        Set objSharedClass = New myClass
    End If

    Set GetShared = objSharedClass

End Function