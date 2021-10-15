Sub GetData()
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
    
    ' Clear screen and go to EA
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
    ibmCurrentScreen.SendKeys ("EA")
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    waitText = "NEXT"
    
    ' Cycle through screens and grab text
    While ibmCurrentScreen.GetText(1, 78, 2) <> "EF"
        returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
        scrapedText = scrapedText & "," & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
        ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
        returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    Wend
    returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
    
    MsgBox ("e screens read")
    
    'Go to bn screen and cycle
    
    ibmCurrentScreen.SendKeys ("bn")
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 1, TextComparisonOption_MatchCase)
    scrapedText = scrapedText & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
    
    ' Stop when data is at certain screen position
    
    While ibmCurrentScreen.GetText(22, 20, 4) <> "DATA"
        returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 1, TextComparisonOption_MatchCase)
        scrapedText = scrapedText & "," & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
        ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
        returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    Wend
    
    MsgBox ("bn screens read")
    
    'go to echl and cycle 22, 13
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
    ibmCurrentScreen.SendKeys ("echl")
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    
    
    While ibmCurrentScreen.GetText(22, 13, 3) <> "END"
        returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
        scrapedText = scrapedText & "," & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
        ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
        returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    Wend
    
    MsgBox ("echl screens read")
    
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
    ibmCurrentScreen.SendKeys ("cmm")
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    scrapedText = scrapedText & "," & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
    
    While ibmCurrentScreen.GetText(22, 40, 2) <> "EA"
        returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
        scrapedText = scrapedText & "," & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
        ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
        returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    Wend
    
    MsgBox ("cmm screens read")
    
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Clear)
    ibmCurrentScreen.SendKeys ("nl")
    ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
    
    While ibmCurrentScreen.GetText(22, 28, 4) = "NEXT"
        returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
        scrapedText = scrapedText & "," & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
        ibmCurrentScreen.SendControlKey (ControlKeyCode_Transmit)
        returnValue = ibmCurrentScreen.WaitForKeyboardEnabled(timeout, 0)
    Wend
    
    MsgBox ("nl screens read")
    
    returnValue = ibmCurrentScreen.WaitForText1(timeout, waitText, 23, 2, TextComparisonOption_MatchCase)
    scrapedText = scrapedText & "," & ibmCurrentScreen.GetTextEx(1, 1, 24, 80, GetTextArea_Block, GetTextWrap_Off, GetTextAttr_Visible, GetTextFlags_CRLF)
    
    MsgBox (scrapedText)
End Sub