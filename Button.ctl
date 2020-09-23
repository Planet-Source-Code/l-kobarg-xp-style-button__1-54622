VERSION 5.00
Begin VB.UserControl Button 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   ForeColor       =   &H8000000C&
   KeyPreview      =   -1  'True
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   ToolboxBitmap   =   "Button.ctx":0000
   Begin VB.Timer TimerMouseOvrCheck 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   750
      Top             =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   150
      Top             =   120
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This XP-Style Button can be placed on any background
'MouseOver and TabStop will be highlighted
'The color of the button can be adapted to any color during runtime

'Please feel invited to visit my homepage
'http://home.t-online.de/home/l.kobarg/clk/
'There you can find a calculator using the XP-Style Button

'if you got any improvements, maybe round, or oval shapes please let me know
'l.kobarg@t-onlien.de

'Based on Leo Barsukov's cool Totally skinned Calculator********
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=38467&lngWId=1

'and Gez Lemon's transparent tutorials
'http://www.juicystudio.com/tutorial/vb/winapi.asp

'known issues:
'during programming the auto-type (auto completion) will not work properly
'if a form using the XP-Style button is open

Option Explicit

Private m_lngHeight As Long
Private m_lngWidth As Long
Private m_blnSkinFromRes As Boolean

'
' Index values for the resource file.
'
Public Enum eImages
    eNone = 0       ' No Value.
    eSkin1 = 1      ' Skin Image 1.
End Enum


'
' Win32 API-Constants.
'
Private Const RGN_OR = 2

'
' Win32 API-Declarations.
'

'*********************************************
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private ButtonLeftPressed As Boolean

'*********************************************



'For drawing the caption
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Rect drawing
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Create/Delete brush
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'For drawing lines
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Misc
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long



Dim m_CurrPoint As POINTAPI



Dim cColor As Long
'Center
Private Const DT_CENTERABS = &H65

'Default system colours
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

'Rectangle
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


'Point
Private Type POINTAPI
        x As Long
        y As Long
End Type

'Events
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseOut()

Private Height      As Long                 'Width
Private Width       As Long                 'Height

Private CurrText    As String               'Current caption
Private CurrFont    As StdFont              'Current font

'Rects structures
Private RC          As RECT
Private RC2         As RECT
Private RC3         As RECT

Private LastButton  As Byte                 'Last button pressed
Private isEnabled   As Boolean              'Enabled or not

'Default system colors
Public cFace        As Long
Private cLight      As Long
Private cHighLight  As Long
Private cShadow     As Long
Private cDarkShadow As Long
Private cText       As Long

Private lastStat    As Byte                 'Last property
Private TE          As String               'Text
Public MausOvr      As Boolean              'maus über dem Button
Private FocusFlag As Boolean                'button hat den focus
Private MausOvrDrawn As Boolean             'maus highlight bereits gemalt


Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Dim n As Integer

'Single click
Private Sub UserControl_Click()
        RaiseEvent Click
End Sub


'Double click
Private Sub UserControl_DblClick()
    
    If LastButton = 1 Then
        'Call the mousedown sub
        RaiseEvent Click
        'UserControl.Refresh
        UserControl_MouseDown 1, 1, 1, 1
    End If
    
End Sub

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = cColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    cColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property


Private Sub UserControl_GotFocus()
    FocusFlag = True
    If Not FocusFlag Then
        Redraw 0, False
    End If
End Sub


Private Sub UserControl_LostFocus()
    FocusFlag = False
    Redraw 0, False
End Sub

'Initialize
Private Sub UserControl_Initialize()
   
    LastButton = 1   'Lastbutton = right mouse button
    RC2.Left = 2
    RC2.Top = 2
    SetColors        'Get default colors
    TimerMouseOvrCheck.Enabled = True
End Sub

'Initialize properties
Private Sub UserControl_InitProperties()

    CurrText = "Caption"                'Caption
    isEnabled = True                    'Enabled
    Set CurrFont = UserControl.Font     'Font
    
End Sub




'Mousedown
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 1 Then ButtonLeftPressed = True
    LastButton = Button     'Set lastbutton
    
    If Button <> 2 Then
        Redraw 2, False     'Redraw button
    End If
    'Raise mousedown event
    RaiseEvent MouseDown(Button, Shift, x, y)
    
End Sub



'Mouseup
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ButtonLeftPressed = False
    If Button <> 2 Then
        Redraw 0, False     'Redraw
    End If
    
    'Raise mousrup event
    RaiseEvent MouseUp(Button, Shift, x, y)
    
End Sub


'Property Get: Caption
Public Property Get Caption() As String
    Caption = CurrText      'Return caption
End Property


'Property Let: Caption
Public Property Let Caption(ByVal newValue As String)
    CurrText = newValue     'Set caption
    Redraw 0, True          'Redraw
    PropertyChanged "TX"    'Last property changed is text
End Property


'Property Get: Enabled
Public Property Get Enabled() As Boolean
    Enabled = isEnabled     'Set enabled/disabled
End Property


'Property Let: Enabled
Public Property Let Enabled(ByVal newValue As Boolean)
    isEnabled = newValue            'Set enabled/disabled
    Redraw 0, True                  'Redraw
    UserControl.Enabled = isEnabled 'Set enabled/disabled
    PropertyChanged "ENAB"          'Last property changed is enabled
End Property


'Property Get: Font
Public Property Get Font() As Font
    Set Font = CurrFont             'Return font
End Property


'Property Set: Font
Public Property Set Font(ByRef newFont As Font)
    Set CurrFont = newFont          'Set font
    Set UserControl.Font = CurrFont 'Set font
    Redraw 0, True                  'Redraw
    PropertyChanged "FONT"          'Last property changed is font
End Property


'Property Get: hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd         'Return hWnd
End Property


'Resize
Private Sub UserControl_Resize()
    
    'Renew dimension variables
    Height = UserControl.ScaleHeight
    Width = UserControl.ScaleWidth
    
    'Set rect1
    RC.Bottom = Height
    RC.Right = Width
    
    'Set rect 2
    RC2.Bottom = Height
    RC2.Right = Width
    
    'Set rect 3
    RC3.Left = 4
    RC3.Top = 4
    RC3.Right = Width - 4
    RC3.Bottom = Height - 4
    
    Redraw 0, True          'Redraw
    
End Sub


'Read Properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    cColor = PropBag.ReadProperty("ForeColor", &H80000012)
    CurrText = PropBag.ReadProperty("TX", "")                       'Caption
    isEnabled = PropBag.ReadProperty("ENAB", True)                  'Enabled
    Set CurrFont = PropBag.ReadProperty("FONT", UserControl.Font)   'Font
    
    UserControl.Enabled = isEnabled     'Set enabled state
    Set UserControl.Font = CurrFont     'Set font
    
    SetColors       'Set colours
    Redraw 0, True  'Redraw
    pCreateSkin (True)
End Sub


'Write properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("ForeColor", cColor, &H80000012)
    PropBag.WriteProperty "TX", CurrText    'Caption
    PropBag.WriteProperty "ENAB", isEnabled 'Enabled state
    PropBag.WriteProperty "FONT", CurrFont  'Font

End Sub


'Redraw
Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)

  Dim i               As Long
  Dim stepXP1         As Single
  Dim XPface          As Long

    'No errors
    If Height = 0 Then Exit Sub
    
    lastStat = curStat  'Set property
    TE = CurrText       'Caption

    With UserControl
        .Cls                                        'Clear control
        'DrawRectangle 0, 0, Width, Height, cFace    'Draw button face
        
        If isEnabled = True Then            'If enabled
            SetTextColor .hdc, cText        'Set text colour
            
            'Button is Up ****************************************************
            If curStat = 0 Then             'If button is up
                
                 'Gradient step
                stepXP1 = 25 / Height
                
                'Shift color
                XPface = ShiftColor(cFace, &H30)
                
                'Draw gradient background
                For i = 2 To Height - 3
                    DrawLine 0, i, Width, i, ShiftColor(XPface, -stepXP1 * i)
                Next
                
                'Set caption
                SetTextColor UserControl.hdc, cColor
                DrawText .hdc, CurrText, Len(CurrText), RC, DT_CENTERABS
                
                'Draw outline
                DrawLine 2, 0, Width - 2, 0, &H733C00                  'upper
                DrawLine 2, Height - 1, Width - 2, Height - 1, &H733C00 'lower
                DrawLine 0, 2, 0, Height - 2, &H733C00                 'left
                DrawLine Width - 1, 2, Width - 1, Height - 2, &H733C00 'right
                
                'Draw corners
                SetPixel UserControl.hdc, 1, 1, &H7B4D10
                SetPixel UserControl.hdc, 1, Height - 2, &H7B4D10
                SetPixel UserControl.hdc, Width - 2, 1, &H7B4D10
                SetPixel UserControl.hdc, Width - 2, Height - 2, &H7B4D10
                
                'Draw shadows
                DrawLine 2, Height - 2, Width - 2, Height - 2, ShiftColor(XPface, -&H30)
                DrawLine 1, Height - 3, Width - 2, Height - 3, ShiftColor(XPface, -&H20)
                DrawLine Width - 2, 2, Width - 2, Height - 2, ShiftColor(XPface, -&H24)
                DrawLine Width - 3, 3, Width - 3, Height - 3, ShiftColor(XPface, -&H18)
                
                'Draw highlights
                DrawLine 2, 1, Width - 2, 1, ShiftColor(XPface, &H10)
                DrawLine 1, 2, Width - 2, 2, ShiftColor(XPface, &HA)
                DrawLine 1, 2, 1, Height - 2, ShiftColor(XPface, -&H5)
                DrawLine 2, 3, 2, Height - 3, ShiftColor(XPface, -&HA)
                
                'Mouse over Button ***********************************
                If MausOvr Then
                    For n = 1 To 1
                        DrawLine n + 1, n, Width - n - 1, n, &H80FF& 'upper
                        DrawLine n + 1, Height - n - 1, Width - n - 1, Height - n - 1, &H80FF& 'lower
                        DrawLine n, n + 1, n, Height - n - 1, &H80FF& 'left
                        DrawLine Width - n - 1, n + 1, Width - n - 1, Height - n - 1, &H80FF& 'right
                    Next n
                    
                    'Draw corners
                    SetPixel UserControl.hdc, 2, 2, &H80FF&          'upper left
                    SetPixel UserControl.hdc, 2, Height - 3, &H80FF& 'lower left
                    SetPixel UserControl.hdc, Width - 3, 2, &H80FF&  'upper right
                    SetPixel UserControl.hdc, Width - 3, Height - 3, &H80FF&   'lower right
                    
                    'MausOvr = False
                End If
                
                'Button got Focus ***********************************
                If FocusFlag Then
                    For n = 2 To 2
                        DrawLine n + 1, n, Width - n - 1, n, &H8000000C      'upper
                        DrawLine n + 1, Height - n - 1, Width - n - 1, Height - n - 1, &H8000000C  'lower
                        DrawLine n, n + 1, n, Height - n - 1, &H8000000C     'left
                        DrawLine Width - n - 1, n + 1, Width - n - 1, Height - n - 1, &H8000000C   'right
                    Next n
                    
                    'Draw corners
                    'SetPixel UserControl.hDC, 3, 3, &H8000000C           'upper left
                    'SetPixel UserControl.hDC, 3, Height - 4, &H8000000C  'lower left
                    'SetPixel UserControl.hDC, Width - 4, 3, &H8000000C   'upper right
                    'SetPixel UserControl.hDC, Width - 4, Height - 4, &H8000000C'lower right
                    'MausOvr = False
                End If
            
            'Button is Down ****************************************************
            ElseIf curStat = 2 Then     'Button is down
                
                'Set gradient step
                stepXP1 = 15 / Height
                
                'Shift color
                XPface = ShiftColor(cFace, &H30)
                XPface = ShiftColor(XPface, -32)
                
                'Draw gradient background
                For i = 3 To Height - 3
                    DrawLine 0, Height - i, Width, Height - i, ShiftColor(XPface, -stepXP1 * i)
                Next i
                         
                'Draw caption
                SetTextColor .hdc, cColor
                DrawText .hdc, CurrText, Len(CurrText), RC2, DT_CENTERABS
                
                'Draw outline
                DrawLine 2, 0, Width - 2, 0, &H733C00                  'upper
                DrawLine 2, Height - 1, Width - 2, Height - 1, &H733C00 'lower
                DrawLine 0, 2, 0, Height - 2, &H733C00                 'left
                DrawLine Width - 1, 2, Width - 1, Height - 2, &H733C00 'right
                
                'Draw corners
                SetPixel UserControl.hdc, 1, 1, &H7B4D10
                SetPixel UserControl.hdc, 1, Height - 2, &H7B4D10
                SetPixel UserControl.hdc, Width - 2, 1, &H7B4D10
                SetPixel UserControl.hdc, Width - 2, Height - 2, &H7B4D10
                
                'Draw shadows
                DrawLine 2, Height - 2, Width - 2, Height - 2, ShiftColor(XPface, &H10)
                DrawLine 1, Height - 3, Width - 2, Height - 3, ShiftColor(XPface, &HA)
                DrawLine Width - 2, 2, Width - 2, Height - 2, ShiftColor(XPface, &H5)
                DrawLine Width - 3, 3, Width - 3, Height - 3, XPface
                
                'Draw highlights
                DrawLine 2, 1, Width - 2, 1, ShiftColor(XPface, -&H20)
                DrawLine 1, 2, Width - 2, 2, ShiftColor(XPface, -&H18)
                DrawLine 1, 2, 1, Height - 2, ShiftColor(XPface, -&H20)
                DrawLine 2, 2, 2, Height - 2, ShiftColor(XPface, -&H16)
            
                'Mouse is over Button ***************************************************
                If MausOvr Then
                    For n = 1 To 1
                        DrawLine n + 1, n, Width - n - 1, n, &H80FF& 'upper
                        DrawLine n + 1, Height - n - 1, Width - n - 1, Height - n - 1, &H80FF& 'lower
                        DrawLine n, n + 1, n, Height - n - 1, &H80FF& 'left
                        DrawLine Width - n - 1, n + 1, Width - n - 1, Height - n - 1, &H80FF& 'right
                    Next n
                    
                    'Draw corners
                    SetPixel UserControl.hdc, 2, 2, &H80FF&          'upper left
                    SetPixel UserControl.hdc, 2, Height - 3, &H80FF& 'lower left
                    SetPixel UserControl.hdc, Width - 3, 2, &H80FF&  'upper right
                    SetPixel UserControl.hdc, Width - 3, Height - 3, &H80FF& 'lower right
                    
                    'MausOvr = False
                End If
                
                'Button got Focus
                If FocusFlag Then
                    For n = 2 To 2
                        DrawLine n + 1, n, Width - n - 1, n, &H8000000C      'oben
                        DrawLine n + 1, Height - n - 1, Width - n - 1, Height - n - 1, &H8000000C 'unten
                        DrawLine n, n + 1, n, Height - n - 1, &H8000000C         'links
                        DrawLine Width - n - 1, n + 1, Width - n - 1, Height - n - 1, &H8000000C 'rechts
                    Next n
                    
                    'Draw corners
                    'SetPixel UserControl.hDC, 3, 3, &H8000000C           'upper left
                    'SetPixel UserControl.hDC, 3, Height - 4, &H8000000C  'lower left
                    'SetPixel UserControl.hDC, Width - 4, 3, &H8000000C   'upper right
                    'SetPixel UserControl.hDC, Width - 4, Height - 4, &H8000000C'lower right
                    'MausOvr = False
                End If
            
            End If
            
        'Button is Disabled *********************************************
        Else    'Disabled state
            
            'Shift color
            XPface = ShiftColor(cFace, &H30)
            'Draw button face
            DrawRectangle 0, 0, Width, Height, ShiftColor(XPface, -&H18)
            'Caption
            SetTextColor .hdc, ShiftColor(XPface, -&H68)
            DrawText .hdc, CurrText, Len(CurrText), RC, DT_CENTERABS
            'Draw outline
            DrawLine 0, 0, Width, 0, ShiftColor(XPface, -&H54)
            DrawLine 1, Height - 1, Width, Height - 1, ShiftColor(XPface, -&H54)
            DrawLine 0, 1, 0, Height, ShiftColor(XPface, -&H54)
            DrawLine Width - 1, 1, Width - 1, Height - 1, ShiftColor(XPface, -&H54)
            'Draw corners
            'SetPixel UserControl.hDC, 1, 1, ShiftColor(XPface, -&H48)
            'SetPixel UserControl.hDC, 1, Height - 2, ShiftColor(XPface, -&H48)
            'SetPixel UserControl.hDC, Width - 2, 1, ShiftColor(XPface, -&H48)
            'SetPixel UserControl.hDC, Width - 2, Height - 2, ShiftColor(XPface, -&H48)
        End If
    End With
    
End Sub


'Draw rectangle
Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)

  Dim bRect As RECT
  Dim hBrush As Long
  Dim Ret As Long
    
    'Fill out rect
    bRect.Left = x
    bRect.Top = y
    bRect.Right = x + Width
    bRect.Bottom = y + Height
    
    'Create brush
    hBrush = CreateSolidBrush(Color)
    
    If OnlyBorder = False Then  'Just border
        Ret = FillRect(UserControl.hdc, bRect, hBrush)
    Else    'Fill whole rect
        Ret = FrameRect(UserControl.hdc, bRect, hBrush)
    End If
    
    'Delete brush
    Ret = DeleteObject(hBrush)
    
End Sub


'Draw line
Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

  Dim pt As POINTAPI

    UserControl.ForeColor = Color           'Set forecolor
    MoveToEx UserControl.hdc, X1, Y1, pt    'Move to X1/Y1
    LineTo UserControl.hdc, X2, Y2          'Draw line to X2/Y2
    
End Sub


'Set Colours
Private Sub SetColors()
    
    'Get system colours and save into variables
    cFace = RGB(200, 200, 255)
    'cFace = RGB(100, 100, 255)
    
    '####################################
    '# cFace = GetSysColor(COLOR_BTNFACE)
    '####################################
    
    cShadow = GetSysColor(COLOR_BTNSHADOW)
    cLight = GetSysColor(COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
    cText = GetSysColor(COLOR_BTNTEXT)
    
End Sub


'Shift colors
Private Function ShiftColor(ByVal Color As Long, ByVal value As Long) As Long

  Dim Red As Long, Blue As Long, Green As Long
    
    'Shift blue
    Blue = ((Color \ &H10000) Mod &H100)
    Blue = Blue + ((Blue * value) \ &HC0)
    'Shift green
    Green = ((Color \ &H100) Mod &H100) + value
    'Shift red
    Red = (Color And &HFF) + value
    
    'Check red bounds
    If Red < 0 Then
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    'Check green bounds
    If Green < 0 Then
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    'Check blue bounds
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If
    
    'Return color
    ShiftColor = RGB(Red, Green, Blue)
  
End Function

Private Sub Timer1_Timer()
  GetCursorPos m_CurrPoint
  ScreenToClient hwnd, m_CurrPoint
  MausOvrDrawn = False
    'if the mouse has left the button, reset everything....

    'Call UserControl_MouseMove(Button, Shift, X, Y)
    'Call Image1_MouseMove(Button, Shift, X, Y)
  If m_CurrPoint.x < UserControl.ScaleLeft Or _
     m_CurrPoint.y < UserControl.ScaleTop Or _
     m_CurrPoint.x > UserControl.ScaleLeft + UserControl.Width / 15 Or _
     m_CurrPoint.y > UserControl.ScaleTop + UserControl.Height / 15 Then
      
       Timer1.Enabled = False
       'Raise the mouse leave event....
       MausOvr = False
       Redraw 0, False
       RaiseEvent MouseOut
       
       TimerMouseOvrCheck.Enabled = True
  End If
End Sub

Private Sub TimerMouseOvrCheck_Timer()
    GetCursorPos m_CurrPoint
    ScreenToClient hwnd, m_CurrPoint
    'if the mouse has left the button, reset everything....
 
    'Call UserControl_MouseMove(Button, Shift, X, Y)
    'Call Image1_MouseMove(Button, Shift, X, Y)
    If Not (m_CurrPoint.x < UserControl.ScaleLeft Or _
        m_CurrPoint.y < UserControl.ScaleTop Or _
        m_CurrPoint.x > UserControl.ScaleLeft + UserControl.Width / 15 Or _
        m_CurrPoint.y > UserControl.ScaleTop + UserControl.Height / 15) Then

            TimerMouseOvrCheck.Enabled = False
            MausOvr = True
                      
            'Redraw 0, False
            If ButtonLeftPressed = True Then      'Right click
                Redraw 2, False     'Redraw Button pressed
            Else
                If Not MausOvrDrawn Then
                    Redraw 0, False     'Redraw Button up
                End If
            End If
       
            MausOvrDrawn = True
            Timer1.Enabled = True
            
            'Raise mousemove event
            'RaiseEvent MouseMove(Button, Shift, X, Y)
            
    End If
End Sub


Public Sub Refesh()
    Redraw 0, False
End Sub

'Skin Part **********************************************
'
' The optional last parameter allows you to specify the image's background color. If left blank, the
' color of the image's top left pixel is used.
'
Public Function fRegionFromBitmap(picSource As Picture, Optional lngBackColor As Long) As Long
    Dim lngReturn As Long
    Dim lngRgnTmp As Long
    Dim lngSkinRgn As Long
    Dim lngStart As Long
    Dim lngRow As Long
    Dim lngCol As Long
    '
    ' Create a rectangular region.
    ' A region is a rectangle, polygon, or ellipse (or a combination of two or more of these shapes)
    ' that can be filled, painted, inverted, framed, and used to perform hit testing (testing for
    ' the cursor location).
    '
    lngSkinRgn = CreateRectRgn(0, 0, 0, 0)
    
    With UserControl
        '
        ' Get the dimensions of the bitmap.
        '
        m_lngHeight = .Height / Screen.TwipsPerPixelY
        m_lngWidth = .Width / Screen.TwipsPerPixelX
        '
        ' If no background color is passed in, get the red, green, blue (RGB) color value of the top
        ' left pixel in the picturebox's device context (DC).
        '
        If lngBackColor < 1 Then lngBackColor = GetPixel(UserControl.hdc, 0, 0)
        '
        ' Loop through the bitmap, row by row, examining each pixel.
        ' In each row, work from left to right comparing each pixel to the background color.
        '
        For lngRow = 0 To m_lngHeight - 1
            lngCol = 0
            Do While lngCol < m_lngWidth
                '
                ' Skip all pixels in a row with the same color as the background color.
                '
                Do While lngCol < m_lngWidth And GetPixel(.hdc, lngCol, lngRow) = lngBackColor
                    lngCol = lngCol + 1
                Loop
                
                If lngCol < m_lngWidth Then
                    '
                    ' Get the start and end of the block of pixels in the row that are not the same
                    ' color as the background.
                    '
                    lngStart = lngCol
                    Do While lngCol < m_lngWidth And GetPixel(.hdc, lngCol, lngRow) <> lngBackColor
                        lngCol = lngCol + 1
                    Loop
                    If lngCol > m_lngWidth Then lngCol = m_lngWidth
                    '
                    ' Create a region equal in size to the line of pixels that don't match the
                    ' background color. Combine this region with our final region.
                    '
                    lngRgnTmp = CreateRectRgn(lngStart, lngRow, lngCol, lngRow + 1)
                    lngReturn = CombineRgn(lngSkinRgn, lngSkinRgn, lngRgnTmp, RGN_OR)
                    Call DeleteObject(lngRgnTmp)
                End If
            Loop
        Next lngRow
    End With
   
    fRegionFromBitmap = lngSkinRgn
End Function

Public Sub pCreateSkin(blnFromRes As Boolean)
    Dim lngRegion As Long
    
    'Screen.MousePointer = vbHourglass
    
   
        
        ' Based on the picture, create a region for Windows to use for our PictureBox and tell
        ' Windows not to paint anything outside this region.
        '
        lngRegion = fRegionFromBitmap(UserControl.Picture)
        Call SetWindowRgn(UserControl.hwnd, lngRegion, True)
        '.picSkin.Picture = LoadPicture("")
   
    
    'Screen.MousePointer = vbDefault
End Sub

'***********************************************************











'




