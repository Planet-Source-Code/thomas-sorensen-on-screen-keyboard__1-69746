Attribute VB_Name = "Module1"

Private Const WH_MOUSE_LL = 14
Private Const HC_ACTION = 0&

Type POINTAPI
    x As Long
    y As Long
End Type

Type MOUSEHOOKSTRUCT
    pt As POINTAPI
End Type

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
   
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cb As Long)

Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201

Public m_Hook As Long

Public Sub Hook()
   m_Hook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf LowLevelMouseProc, App.hInstance, 0&)
End Sub

Public Sub UnHook()
    UnhookWindowsHookEx m_Hook
End Sub

Public Function LowLevelMouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim MouseData As POINTAPI
    
    If nCode = HC_ACTION Then
 
        CopyMemory MouseData, ByVal lParam, Len(MouseData)
        FormChangeColor vbInactiveTitleBarText
        MouseOverButton MouseData.x, MouseData.y, vbGreen
            
            If wParam = WM_LBUTTONDOWN Then

      
                If IsInForm(MouseData.x, MouseData.y) = True Then
                    MousePressButton MouseData.x, MouseData.y, vbBlue ' if the left mousebutton is pressed, active the button that its over
                  '  MouseOverButton MouseData.x, MouseData.y
                    LowLevelMouseProc = 1 ' does not send the mousepress to the form
                    Exit Function
                End If
            
            End If

    End If

    LowLevelMouseProc = CallNextHookEx(m_Hook, nCode, wParam, lParam)

   
End Function
Public Function MousePressButton(ByVal x As Integer, ByVal y As Integer, ByVal ChColor As String)
'gets the x and y to all the buttons. Checks if one of them was hit, and sends the caption to the active window

Dim IsWidth As Boolean
Dim IsHeight As Boolean

Dim c As Control
For Each c In Form1.Controls
    If TypeOf c Is Label Then
        
        If x > Int(Form1.Left / 15) + c.Left And x < Int(Form1.Left / 15) + c.Left + c.Width Then
        IsWidth = True
        Else
        IsWidth = False
        End If

        If y > Int(Form1.Top / 15) + c.Top + c.Height And y < Int(Form1.Top / 15) + c.Height + c.Top + c.Height Then
        IsHeight = True
        Else
        IsHeight = False
        End If

        If IsWidth = True And IsHeight = True Then
        c.BackColor = ChColor
        
            If Len(c.Caption) = 1 Then
            'If the lenght of the caption is only one character, then send it as it is
            SendKeys (c.Caption)
            Else 'if the lenght of the caption is longer, then it is a special command
        
                Select Case c.Caption
        
                Case "Enter"
                SendKeys ("{ENTER}")
                Case "Tab"
                SendKeys ("{TAB}")
                Case "<Back"
                SendKeys ("{BACKSPACE}")
                End Select
        
            End If

        End If
    End If
Next c

End Function

Public Function MouseOverButton(ByVal x As Integer, ByVal y As Integer, ByVal ChColor As String)
'changes the color over the button

Dim IsWidth As Boolean
Dim IsHeight As Boolean

Dim c As Control

For Each c In Form1.Controls
    If TypeOf c Is Label Then
        
        If x > Int(Form1.Left / 15) + c.Left And x < Int(Form1.Left / 15) + c.Left + c.Width Then
        IsWidth = True
        Else
        IsWidth = False
        End If

        If y > Int(Form1.Top / 15) + c.Top + 20 And y < Int(Form1.Top / 15) + c.Top + c.Height + 20 Then ' the + 20 is the height of the title top
        IsHeight = True
        Else
        IsHeight = False
        End If

        If IsWidth = True And IsHeight = True Then
    
        c.BackColor = ChColor
        
         
        End If
        
    End If
Next c

End Function

Public Function FormChangeColor(ByVal ChColor As String)
'changes all buttons to one color - not the one that the mouse is over.
Dim c As Control
For Each c In Form1.Controls
    If TypeOf c Is Label Then
    
    c.BackColor = vbInactiveTitleBarText
    End If
Next c


End Function

Public Function IsInForm(ByVal x As Integer, ByVal y As Integer) As Boolean
'Checks if the mouse is over the form
form1width = Int(Form1.Width / 15)
Form1Height = Int(Form1.Height / 15) - 25
form1left = Int(Form1.Left / 15)
form1top = Int(Form1.Top / 15) + 25
Dim IsWidth As Boolean
Dim IsHeight As Boolean


If x > form1left And x < form1left + form1width Then

IsWidth = True
Else
IsWidth = False
End If

If y > form1top And y < form1top + Form1Height Then
IsHeight = True
Else
IsHeight = False
End If

If IsWidth = True And IsHeight = True Then IsInForm = True

End Function


