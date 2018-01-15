Attribute VB_Name = "STK4_Perferences"
Public gToggleState As Boolean

Public MyRibbon As IRibbonUI

'Callback for customUI.******
Sub RibbonOnLoad(ribbon As IRibbonUI)
   Set MyRibbon = ribbon
   
   '--read previously saved value of toggle
   gToggleState = ThisWorkbook.Sheets("Prefs").Range("A2").Value
   MyRibbon.ActivateTab "tabStock1"
End Sub

'Callback for TbtnToggleHideColumn onAction
Sub Edit_ThisWorkbook(control As IRibbonControl, pressed As Boolean)
  
  '--switch state of global variable
  gToggleState = Not gToggleState
  
  '--save new state to worksheet
  ThisWorkbook.Sheets("Prefs").Range("A2").Value = gToggleState
  ThisWorkbook.IsAddin = Not gToggleState
  
End Sub

'Callback for TbtnToggleHideColumn getPressed
Sub GetPressed(control As IRibbonControl, ByRef returnedVal)
  returnedVal = gToggleState
End Sub




