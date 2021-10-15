
Dim ibmCurrentTerminal As IbmTerminal
Dim ibmCurrentScreen As IbmScreen
Dim hiddenTextEntry As String
Dim returnValue As Integer
Dim timeout As Integer
Dim waitText As String
Dim dataText As String
Dim scrapedText As String
timeout = 15000

Set ibmCurrentTerminal = ThisFrame.SelectedView.control
Set ibmCurrentScreen = ibmCurrentTerminal.screen

ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)



returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
ibmCurrentScreen.SendKeys ("ea")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)



returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
waitText = "NEXT"
returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
ibmCurrentScreen.SendKeys ("eb")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)



returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
waitText = "NEXT"
returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
ibmCurrentScreen.SendKeys ("ec")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)



returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
waitText = "NEXT"
returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
ibmCurrentScreen.SendKeys ("ed")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)

returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
waitText = "NEXT"
returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
ibmCurrentScreen.SendKeys ("ee")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)


ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
ibmCurrentScreen.SendKeys ("ea")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)

returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
waitText = "NEXT"

While ibmCurrentScreen.GetText(1, 78, 2) <> "EF"
    returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
    scrapedText = scrapedText & "," & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
    ibmCurrentScreen.SendControlKey(ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
End While

While ibmCurrentScreen.GetText(22, 20, 4) <> data
    returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 1, TextComparisonOption_MatchCase)
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
End While


returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
waitText = "NEXT"
returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)




ibmCurrentScreen.SendKeys ("bn")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)



stopText = "DATA"

returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
waitText = "NEXT"
returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 1, TextComparisonOption_MatchCase)

While ibmCurrentScreen.GetText(22, 20, 4) <> data
    returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 1, TextComparisonOption_MatchCase)
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
End While




ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
ibmCurrentScreen.SendKeys ("echl")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)

While ## CONDITION ## <> data
    returnValue = ## WAIT TEXT LOCATION ##
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
End While




ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
ibmCurrentScreen.SendKeys ("cmm")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)

While ## CONDITION ## <> STOPCONDITION
    returnValue = ## WAIT TEXT LOCATION ##
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
End While




ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
ibmCurrentScreen.SendKeys ("nl")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)

While ## CONDITION ## <> STOPCONDITION
    returnValue = ## WAIT TEXT LOCATION ##
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
End While




ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
ibmCurrentScreen.SendKeys ("py")
ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)

While ## CONDITION ## <> STOPCONDITION
    returnValue = ## WAIT TEXT LOCATION ##
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
End While