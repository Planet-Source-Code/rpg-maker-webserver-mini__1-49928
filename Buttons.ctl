VERSION 5.00
Begin VB.UserControl CmdButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   DefaultCancel   =   -1  'True
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   135
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   360
   End
End
Attribute VB_Name = "CmdButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const Version As String = "1.1"
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1 Or DT_WORDBREAK

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal HWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF = 4

Private Declare Function GetClientRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum ButtonTypes
    [Windows 16-bit] = 1    'the old-fashioned Win16 button
    [Windows 32-bit] = 2    'the classic windows button
    [Windows XP] = 3        'the new brand XP button totally owner-drawn
    [Mac] = 4               'i suppose it looks exactly as a Mac button... i took the style from a GetRight skin!!!
    [Java metal] = 5        'there are also other styles but not so different from windows one
    [Netscape 6] = 6        'this is the button displayed in web-pages, it also appears in some java apps
    [Simple Flat] = 7       'the standard flat button seen on toolbars
    [Flat Highlight] = 8    'again the flat button but this one has no border until the mouse is over it
End Enum

Public Enum ColorTypes
    [Use Windows] = 1
    [Custom] = 2
    [Force Standard] = 3
    [Use Container] = 4
End Enum

'events
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()

'variables
Private MyButtonType As ButtonTypes
Private MyColorType As ColorTypes

Private He As Long  'the height of the button
Private Wi As Long  'the width of the button

Private BackC As Long 'back color
Private ForeC As Long 'fore color
Private ForeO As Long 'fore color when mouse is over

Private elTex As String     'current text
Private WithEvents TextFont As StdFont 'current font
Attribute TextFont.VB_VarHelpID = -1

Private rc As RECT, rc2 As RECT, rc3 As RECT
Private rgnNorm As Long

Private LastButton As Byte, LastKeyDown As Byte
Private isEnabled As Boolean
Private hasFocus As Boolean, showFocusR As Boolean

Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long, cTextO As Long

Private lastStat As Byte, TE As String 'used to avoid unnecessary repaints
Private isOver As Boolean

Private Sub OverTimer_Timer()
Dim pt As POINTAPI

GetCursorPos pt
If UserControl.HWnd <> WindowFromPoint(pt.X, pt.Y) Then
    OverTimer.Enabled = False
    isOver = False
    Call Redraw(0, True)
    RaiseEvent MouseOut
End If
End Sub

Private Sub TextFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = TextFont
    Call CalcTextRects
    Call Redraw(0, True)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If MyColorType = [Use Container] Then
        Call SetColors
        Call Redraw(lastStat, True)
    End If
End Sub

Private Sub UserControl_Click()
If (LastButton = 1) And (isEnabled = True) Then
    Call Redraw(0, True) 'be sure that the normal status is drawn
    UserControl.Refresh
    RaiseEvent Click
End If
End Sub

Private Sub UserControl_DblClick()
If LastButton = 1 Then
    Call UserControl_MouseDown(1, 1, 1, 1)
End If
End Sub

Private Sub UserControl_GotFocus()
hasFocus = True
Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)

LastKeyDown = KeyCode
If KeyCode = 32 Then 'spacebar pressed
    Call UserControl_MouseDown(1, 1, 1, 1)
ElseIf (KeyCode = 39) Or (KeyCode = 40) Then 'right and down arrows
    SendKeys "{Tab}"
ElseIf (KeyCode = 37) Or (KeyCode = 38) Then 'left and up arrows
    SendKeys "+{Tab}"
End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)

If (KeyCode = 32) And (LastKeyDown = 32) Then 'spacebar pressed
    Call UserControl_MouseUp(1, 1, 1, 1)
    LastButton = 1
    Call UserControl_Click
End If
End Sub

Private Sub UserControl_LostFocus()
hasFocus = False
Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Initialize()
LastButton = 1
Call SetColors
End Sub

Private Sub UserControl_InitProperties()
    isEnabled = True
    showFocusR = True
    elTex = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
    Set TextFont = Ambient.Font
    MyButtonType = [Windows 32-bit]
    MyColorType = [Use Windows]
    BackC = GetSysColor(COLOR_BTNFACE)
    ForeC = GetSysColor(COLOR_BTNTEXT)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)

LastButton = Button
If Button <> 2 Then Call Redraw(2, False)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
If Button < 2 Then
    If X < 0 Or Y < 0 Or X > Wi Or Y > He Then
        'we are outside the button
        Call Redraw(0, False)
    Else
        'we are inside the button
        If (Button = 0) And (isOver = False) Then
            OverTimer.Enabled = True
            isOver = True
            RaiseEvent MouseOver
            Call Redraw(0, True)
            'UserControl.MouseIcon = FrmMenu.MouseIcon
            'UserControl.MousePointer = vbCustom
        ElseIf Button = 1 Then
            Call Redraw(2, False)
        End If
    End If
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
If Button <> 2 Then Call Redraw(0, False)
End Sub

'########## BUTTON PROPERTIES ##########
Public Property Get BackColor() As OLE_COLOR
BackColor = BackC
End Property
Public Property Let BackColor(ByVal theCol As OLE_COLOR)
BackC = theCol
Call SetColors
Call Redraw(lastStat, True)
PropertyChanged "BCOL"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = ForeC
End Property
Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
ForeC = theCol
If Ambient.UserMode = False Then ForeO = theCol
Call SetColors
Call Redraw(lastStat, True)
PropertyChanged "FCOL"
End Property

Public Property Get ForeOver() As OLE_COLOR
ForeOver = ForeO
End Property
Public Property Let ForeOver(ByVal theCol As OLE_COLOR)
ForeO = theCol
Call SetColors
Call Redraw(lastStat, True)
PropertyChanged "FCOLO"
End Property

Public Property Get ButtonType() As ButtonTypes
ButtonType = MyButtonType
End Property

Public Property Let ButtonType(ByVal newValue As ButtonTypes)
MyButtonType = newValue
If ButtonType = [Java metal] Then UserControl.FontBold = True
Call UserControl_Resize
PropertyChanged "BTYPE"
End Property

Public Property Get Caption() As String
Caption = elTex
End Property

Public Property Let Caption(ByVal newValue As String)
elTex = newValue
Call SetAccessKeys
Call CalcTextRects
Call Redraw(0, True)
PropertyChanged "TX"
End Property

Public Property Get Enabled() As Boolean
Enabled = isEnabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
isEnabled = newValue
Call Redraw(0, True)
UserControl.Enabled = isEnabled
PropertyChanged "ENAB"
End Property

Public Property Get Font() As Font
Set Font = TextFont
End Property

Public Property Set Font(ByRef newFont As Font)
Set TextFont = newFont
Set UserControl.Font = TextFont
Call CalcTextRects
Call Redraw(0, True)
PropertyChanged "FONT"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
FontBold = TextFont.Bold
End Property

Public Property Let FontBold(ByVal newValue As Boolean)
TextFont.Bold = newValue
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
FontItalic = TextFont.Italic
End Property

Public Property Let FontItalic(ByVal newValue As Boolean)
TextFont.Italic = newValue
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
FontUnderline = TextFont.Underline
End Property

Public Property Let FontUnderline(ByVal newValue As Boolean)
TextFont.Underline = newValue
End Property

Public Property Get FontSize() As Integer
Attribute FontSize.VB_MemberFlags = "400"
FontSize = TextFont.size
End Property

Public Property Let FontSize(ByVal newValue As Integer)
TextFont.size = newValue
End Property

Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
FontName = TextFont.Name
End Property

Public Property Let FontName(ByVal newValue As String)
TextFont.Name = newValue
End Property

'it is very common that a windows user uses custom color
'schemes to view his/her desktop, and is also very
'common that this color scheme has weird colors that
'would alter the nice look of my buttons.
'So if you want to force the button to use the windows
'standard colors you may change this property to "Force Standard"

Public Property Get ColorScheme() As ColorTypes
ColorScheme = MyColorType
End Property

Public Property Let ColorScheme(ByVal newValue As ColorTypes)
MyColorType = newValue
Call SetColors
Call Redraw(0, True)
PropertyChanged "COLTYPE"
End Property

Public Property Get ShowFocusRect() As Boolean
ShowFocusRect = showFocusR
End Property

Public Property Let ShowFocusRect(ByVal newValue As Boolean)
showFocusR = newValue
Call Redraw(lastStat, True)
PropertyChanged "FOCUSR"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal newPointer As MousePointerConstants)
    UserControl.MousePointer = newPointer
    PropertyChanged "MPTR"
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal newIcon As StdPicture)
On Local Error Resume Next
    Set UserControl.MouseIcon = newIcon
    PropertyChanged "MICON"
End Property

Public Property Get HWnd() As Long
    HWnd = UserControl.HWnd
End Property

'########## END OF PROPERTIES ##########

Private Sub UserControl_Resize()
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    
    GetClientRect UserControl.HWnd, rc3: InflateRect rc3, -4, -4
    Call CalcTextRects
    
    DeleteObject rgnNorm
    Call MakeRegion
    SetWindowRgn UserControl.HWnd, rgnNorm, True
    
    If He > 0 Then Call Redraw(0, True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    MyButtonType = .ReadProperty("BTYPE", 2)
    elTex = .ReadProperty("TX", "")
    isEnabled = .ReadProperty("ENAB", True)
    Set TextFont = .ReadProperty("FONT", UserControl.Font)
    MyColorType = .ReadProperty("COLTYPE", 1)
    showFocusR = .ReadProperty("FOCUSR", True)
    BackC = .ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
    ForeC = .ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))
    ForeO = .ReadProperty("FCOLO", GetSysColor(COLOR_BTNTEXT))
    UserControl.MousePointer = .ReadProperty("MPTR", 0)
    Set UserControl.MouseIcon = .ReadProperty("MICON", Nothing)
End With

    UserControl.Enabled = isEnabled
    Set UserControl.Font = TextFont
    Call SetColors
    Call SetAccessKeys
    Call Redraw(0, True)

End Sub

Private Sub UserControl_Terminate()
    DeleteObject rgnNorm
    Set TextFont = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    Call .WriteProperty("BTYPE", MyButtonType)
    Call .WriteProperty("TX", elTex)
    Call .WriteProperty("ENAB", isEnabled)
    Call .WriteProperty("FONT", TextFont)
    Call .WriteProperty("COLTYPE", MyColorType)
    Call .WriteProperty("FOCUSR", showFocusR)
    Call .WriteProperty("BCOL", BackC)
    Call .WriteProperty("FCOL", ForeC)
    Call .WriteProperty("FCOLO", ForeO)
    Call .WriteProperty("MPTR", UserControl.MousePointer)
    Call .WriteProperty("MICON", UserControl.MouseIcon)
End With
End Sub

Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)

'here is the CORE of the button, everything is drawn here
'it's not well commented but i think that everything is
'pretty self explanatory...

If Force = False Then 'check drawing redundancy
    If (curStat = lastStat) And (TE = elTex) Then Exit Sub
End If

If He = 0 Then Exit Sub 'we don't want errors

lastStat = curStat
TE = elTex

Dim I As Long, stepXP1 As Single, XPface As Long

With UserControl
.Cls

DrawRectangle 0, 0, Wi, He, cFace

If isEnabled = True Then
    'set font color
    If isOver Then
        SetTextColor .hdc, cTextO
    Else
        SetTextColor .hdc, cText
    End If
    If curStat = 0 Then
'#@#@#@#@#@# BUTTON NORMAL STATE #@#@#@#@#@#
        Select Case MyButtonType
            Case 1 'Windows 16-bit
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                DrawFrame cHighLight, cShadow, cHighLight, cShadow, True
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                If hasFocus Then DrawFocusR
            Case 2 'Windows 32-bit
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                If (Ambient.DisplayAsDefault = True) And (showFocusR = True) Then
                    DrawFrame cHighLight, cDarkShadow, cLight, cShadow, True
                    If hasFocus Then DrawFocusR
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                Else
                    DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False
                End If
            Case 3 'Windows XP
                stepXP1 = 25 / He
                XPface = ShiftColor(cFace, &H30, True)
                For I = 1 To He
                    DrawLine 0, I, Wi, I, ShiftColor(XPface, -stepXP1 * I, True)
                Next
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                DrawRectangle 0, 0, Wi, He, &H733C00, True
                mSetPixel 1, 1, &H7B4D10
                mSetPixel 1, He - 2, &H7B4D10
                mSetPixel Wi - 2, 1, &H7B4D10
                mSetPixel Wi - 2, He - 2, &H7B4D10
                
                If isOver Then
                    DrawRectangle 1, 2, Wi - 2, He - 4, &H31B2FF, True
                    DrawLine 2, He - 2, Wi - 2, He - 2, &H96E7&
                    DrawLine 2, 1, Wi - 2, 1, &HCEF3FF
                    DrawLine 1, 2, Wi - 1, 2, &H8CDBFF
                    DrawLine 2, 3, 2, He - 3, &H6BCBFF
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, &H6BCBFF
                ElseIf ((hasFocus Or Ambient.DisplayAsDefault) And showFocusR) Then
                    DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                    DrawLine 2, He - 2, Wi - 2, He - 2, &HEF826B
                    DrawLine 2, 1, Wi - 2, 1, &HFFE7CE
                    DrawLine 1, 2, Wi - 1, 2, &HF7D7BD
                    
                    DrawLine 2, 3, 2, He - 3, &HF0D1B5
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5
                Else 'we do not draw the bevel always because the above code would repaint over it
                    DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPface, -&H30, True)
                    DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPface, -&H20, True)
                    DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPface, -&H24, True)
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, ShiftColor(XPface, -&H18, True)
                    DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPface, &H10, True)
                    DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPface, &HA, True)
                    DrawLine 1, 2, 1, He - 2, ShiftColor(XPface, -&H5, True)
                    DrawLine 2, 3, 2, He - 3, ShiftColor(XPface, -&HA, True)
                End If
            Case 4 'Mac
                DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                mSetPixel 1, 1, cDarkShadow
                mSetPixel 1, He - 2, cDarkShadow
                mSetPixel Wi - 2, 1, cDarkShadow
                mSetPixel Wi - 2, He - 2, cDarkShadow
                mSetPixel 1, 2, cFace
                mSetPixel 2, 1, cFace
                DrawLine 3, 2, Wi - 3, 2, cHighLight
                DrawLine 2, 2, 2, He - 3, cHighLight
                mSetPixel 3, 3, cHighLight
                DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
                DrawLine 1, He - 3, Wi - 3, He - 3, cFace
                mSetPixel Wi - 4, He - 4, cFace
                DrawLine Wi - 2, 3, Wi - 2, He - 2, cShadow
                DrawLine 3, He - 2, Wi - 2, He - 2, cShadow
                mSetPixel Wi - 3, He - 3, cShadow
                mSetPixel 2, He - 2, cFace
                mSetPixel 2, He - 3, cLight
                mSetPixel Wi - 2, 2, cFace
                mSetPixel Wi - 3, 2, cLight
            Case 5 'Java
                DrawRectangle 1, 1, Wi - 1, He - 1, ShiftColor(cFace, &HC)
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                DrawRectangle 1, 1, Wi - 1, He - 1, cHighLight, True
                DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                mSetPixel 1, He - 2, ShiftColor(cShadow, &H1A)
                mSetPixel Wi - 2, 1, ShiftColor(cShadow, &H1A)
                If hasFocus And showFocusR Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
            Case 6 'Netscape
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                DrawFrame ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, False
                If hasFocus Then DrawFocusR
             Case 7, 8 'Flat buttons
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                If Not (MyButtonType = [Flat Highlight]) Then
                    DrawFrame cHighLight, cShadow, 0, 0, False, True
                ElseIf isOver Then
                    DrawFrame cHighLight, cShadow, 0, 0, False, True
                End If
                If hasFocus Then DrawFocusR
        End Select
    ElseIf curStat = 2 Then
'#@#@#@#@#@# BUTTON IS DOWN #@#@#@#@#@#
        Select Case MyButtonType
            Case 1 'Windows 16-bit
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
                DrawFrame cShadow, cHighLight, cShadow, cHighLight, True
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                If hasFocus Then DrawFocusR
            Case 2 'Windows 32-bit
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
                
                If showFocusR = True Then
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                    DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                    If hasFocus Then DrawFocusR
                Else
                    DrawFrame cDarkShadow, cHighLight, cShadow, cLight, False
                End If
            Case 3 'Windows XP
                stepXP1 = 25 / He
                XPface = ShiftColor(cFace, &H30, True)
                XPface = ShiftColor(XPface, -32, True)
                For I = 1 To He
                    DrawLine 0, He - I, Wi, He - I, ShiftColor(XPface, -stepXP1 * I, True)
                Next
                SetTextColor .hdc, cText
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
                DrawRectangle 0, 0, Wi, He, &H733C00, True
                mSetPixel 1, 1, &H7B4D10
                mSetPixel 1, He - 2, &H7B4D10
                mSetPixel Wi - 2, 1, &H7B4D10
                mSetPixel Wi - 2, He - 2, &H7B4D10
                
                DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPface, &H10, True)
                DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPface, &HA, True)
                DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPface, &H5, True)
                DrawLine Wi - 3, 3, Wi - 3, He - 3, XPface
                DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPface, -&H20, True)
                DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPface, -&H18, True)
                DrawLine 1, 2, 1, He - 2, ShiftColor(XPface, -&H20, True)
                DrawLine 2, 2, 2, He - 2, ShiftColor(XPface, -&H16, True)
            Case 4 'Mac
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                SetTextColor .hdc, cLight
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H40), True
                DrawRectangle 2, 2, Wi - 4, He - 4, ShiftColor(cShadow, -&H20), True
                mSetPixel 2, 2, ShiftColor(cShadow, -&H40)
                mSetPixel 3, 3, ShiftColor(cShadow, -&H20)
                mSetPixel 1, 1, cDarkShadow
                mSetPixel 1, He - 2, cDarkShadow
                mSetPixel Wi - 2, 1, cDarkShadow
                mSetPixel Wi - 2, He - 2, cDarkShadow
                DrawLine Wi - 3, 1, Wi - 3, He - 3, cShadow
                DrawLine 1, He - 3, Wi - 2, He - 3, cShadow
                mSetPixel Wi - 4, He - 4, cShadow
                DrawLine Wi - 2, 3, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                DrawLine 3, He - 2, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 2, He - 3, ShiftColor(cShadow, -&H20)
                mSetPixel Wi - 3, He - 2, ShiftColor(cShadow, -&H20)

                mSetPixel 2, He - 2, ShiftColor(cShadow, -&H20)
                mSetPixel 2, He - 3, ShiftColor(cShadow, -&H10)
                mSetPixel 1, He - 3, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 2, 2, ShiftColor(cShadow, -&H20)
                mSetPixel Wi - 3, 2, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 3, 1, ShiftColor(cShadow, -&H10)
            Case 5 'Java
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, &H10), False
                DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                DrawLine Wi - 1, 1, Wi - 1, He, cHighLight
                DrawLine 1, He - 1, Wi - 1, He - 1, cHighLight
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                If hasFocus And showFocusR Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
            Case 6 'Netscape
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
                DrawFrame cShadow, ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), False
                If hasFocus Then DrawFocusR
             Case 7, 8 'Flat buttons
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
                DrawFrame cShadow, cHighLight, 0, 0, False, True
                If hasFocus Then DrawFocusR
        End Select
    End If
Else
'#~#~#~#~#~# DISABLED STATUS #~#~#~#~#~#
    Select Case MyButtonType
        Case 1 'Windows 16-bit
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
            DrawFrame cHighLight, cShadow, cHighLight, cShadow, True
            DrawRectangle 0, 0, Wi, He, cDarkShadow, True
        Case 2 'Windows 32-bit
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
            DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False
        Case 3 'Windows XP
            XPface = ShiftColor(cFace, &H30, True)
            DrawRectangle 0, 0, Wi, He, ShiftColor(XPface, -&H18, True)
            SetTextColor .hdc, ShiftColor(XPface, -&H68, True)
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
            DrawRectangle 0, 0, Wi, He, ShiftColor(XPface, -&H54, True), True
            mSetPixel 1, 1, ShiftColor(XPface, -&H48, True)
            mSetPixel 1, He - 2, ShiftColor(XPface, -&H48, True)
            mSetPixel Wi - 2, 1, ShiftColor(XPface, -&H48, True)
            mSetPixel Wi - 2, He - 2, ShiftColor(XPface, -&H48, True)
        Case 4 'Mac
            DrawRectangle 1, 1, Wi - 2, He - 2, cLight
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
            DrawRectangle 0, 0, Wi, He, cDarkShadow, True
            mSetPixel 1, 1, cDarkShadow
            mSetPixel 1, He - 2, cDarkShadow
            mSetPixel Wi - 2, 1, cDarkShadow
            mSetPixel Wi - 2, He - 2, cDarkShadow
            mSetPixel 1, 2, cFace
            mSetPixel 2, 1, cFace
            DrawLine 3, 2, Wi - 3, 2, cHighLight
            DrawLine 2, 2, 2, He - 3, cHighLight
            mSetPixel 3, 3, cHighLight
            DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
            DrawLine 1, He - 3, Wi - 3, He - 3, cFace
            mSetPixel Wi - 4, He - 4, cFace
            DrawLine Wi - 2, 3, Wi - 2, He - 2, cShadow
            DrawLine 3, He - 2, Wi - 2, He - 2, cShadow
            mSetPixel Wi - 3, He - 3, cShadow
            mSetPixel 2, He - 2, cFace
            mSetPixel 2, He - 3, cLight
            mSetPixel Wi - 2, 2, cFace
            mSetPixel Wi - 3, 2, cLight
        Case 5 'Java
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
            DrawRectangle 0, 0, Wi, He, cShadow, True
        Case 6 'Netscape
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
            DrawFrame ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, False
        Case 7, 8 'Flat buttons
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
            If MyButtonType = [Simple Flat] Then
                DrawFrame cHighLight, cShadow, 0, 0, False, True
            Else
                DrawRectangle 0, 0, Wi, He, cShadow, True
            End If
    End Select
End If
End With

End Sub

Private Sub DrawRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
'this is my custom function to draw rectangles and frames
'it's faster and smoother than using the line method

Dim bRect As RECT
Dim hBrush As Long
Dim Ret As Long

bRect.Left = X
bRect.Top = Y
bRect.Right = X + Width
bRect.Bottom = Y + Height

hBrush = CreateSolidBrush(Color)

If OnlyBorder = False Then
    Ret = FillRect(UserControl.hdc, bRect, hBrush)
Else
    Ret = FrameRect(UserControl.hdc, bRect, hBrush)
End If

Ret = DeleteObject(hBrush)
End Sub

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
'a fast way to draw lines
Dim pt As POINTAPI

With UserControl
    Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, Color)))
    
    MoveToEx .hdc, X1, Y1, pt
    LineTo .hdc, X2, Y2
End With
End Sub

Private Sub DrawFrame(ByVal ColHigh As Long, ByVal ColDark As Long, ByVal ColLight As Long, ByVal ColShadow As Long, ByVal ExtraOffset As Boolean, Optional ByVal Flat As Boolean = False)
'a very fast way to draw windows-like frames
Dim pt As POINTAPI
Dim frHe As Long, frWi As Long, frXtra As Long

frHe = He - 1 + ExtraOffset: frWi = Wi - 1 + ExtraOffset: frXtra = Abs(ExtraOffset)

With UserControl
    Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColHigh)))
    '=============================
    MoveToEx .hdc, frXtra, frHe, pt
    LineTo .hdc, frXtra, frXtra
    LineTo .hdc, frWi, frXtra
    '=============================
    Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColDark)))
    '=============================
    LineTo .hdc, frWi, frHe
    LineTo .hdc, frXtra - 1, frHe
    MoveToEx .hdc, frXtra + 1, frHe - 1, pt
    If Flat Then Exit Sub
    '=============================
    Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColLight)))
    '=============================
    LineTo .hdc, frXtra + 1, frXtra + 1
    LineTo .hdc, frWi - 1, frXtra + 1
    '=============================
    Call DeleteObject(SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColShadow)))
    '=============================
    LineTo .hdc, frWi - 1, frHe - 1
    LineTo .hdc, frXtra, frHe - 1
End With
End Sub

Private Sub mSetPixel(ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    Call SetPixel(UserControl.hdc, X, Y, Color)
End Sub

Private Sub DrawFocusR()
If showFocusR Then
    SetTextColor UserControl.hdc, cText
    DrawFocusRect UserControl.hdc, rc3
End If
End Sub
Private Sub SetColors()
'this function sets the colors taken as a base to build
'all the other colors and styles.

If MyColorType = Custom Then
    cFace = BackC
    cText = ForeC
    cTextO = ForeO
    cShadow = ShiftColor(cFace, -&H40)
    cLight = ShiftColor(cFace, &H1F)
    cHighLight = ShiftColor(cFace, &H2F) 'it should be 3F but it looks too lighter
    cDarkShadow = ShiftColor(cFace, -&HC0)
ElseIf MyColorType = [Force Standard] Then
    cFace = &HC0C0C0
    cShadow = &H808080
    cLight = &HDFDFDF
    cDarkShadow = &H0
    cHighLight = &HFFFFFF
    cText = &H0
    cTextO = cText
ElseIf MyColorType = [Use Container] Then
    cFace = GetBkColor(UserControl.Parent.hdc)
    cText = GetTextColor(UserControl.Parent.hdc)
    cTextO = cText
    cShadow = ShiftColor(cFace, -&H40)
    cLight = ShiftColor(cFace, &H1F)
    cHighLight = ShiftColor(cFace, &H2F)
    cDarkShadow = ShiftColor(cFace, -&HC0)
Else
'if MyColorType is 1 or has not been set then use windows colors
    cFace = GetSysColor(COLOR_BTNFACE)
    cShadow = GetSysColor(COLOR_BTNSHADOW)
    cLight = GetSysColor(COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
    cText = GetSysColor(COLOR_BTNTEXT)
    cTextO = cText
End If
End Sub

Private Sub MakeRegion()
'this function creates the regions to "cut" the UserControl
'so it will be transparent in certain areas

Dim rgn1 As Long, rgn2 As Long
    
    DeleteObject rgnNorm
    rgnNorm = CreateRectRgn(0, 0, Wi, He)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    
Select Case MyButtonType
    Case 1 'Windows 16-bit
        rgn1 = CreateRectRgn(0, 0, 1, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 1, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 1, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    Case 3, 4 'Windows XP and Mac
        rgn1 = CreateRectRgn(0, 0, 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, 1, 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
    Case 5 'Java
        rgn1 = CreateRectRgn(0, He, 1, He - 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
End Select

DeleteObject rgn1
DeleteObject rgn2
End Sub

Private Sub SetAccessKeys()
'this is a TRUE access keys parser
'i hate seeing how other programmers just check for the
'existence of the ampersand regardless of what follows it
'the basic rule is that if an ampersand is followed by another,
'  a single ampersand is drawn and this is not the access key.
'  So we continue searching for another possible access key.

Dim ampersandPos As Long

If Len(elTex) > 1 Then
    ampersandPos = InStr(1, elTex, "&", vbTextCompare)
    If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
        If Mid(elTex, ampersandPos + 1, 1) <> "&" Then 'if text is sonething like && then no access key should be assigned, so continue searching
            UserControl.AccessKeys = LCase(Mid(elTex, ampersandPos + 1, 1))
        Else 'do only a second pass to find another ampersand character
            ampersandPos = InStr(ampersandPos + 2, elTex, "&", vbTextCompare)
            If Mid(elTex, ampersandPos + 1, 1) <> "&" Then
                UserControl.AccessKeys = LCase(Mid(elTex, ampersandPos + 1, 1))
            Else
                UserControl.AccessKeys = ""
            End If
        End If
    Else
        UserControl.AccessKeys = ""
    End If
Else
    UserControl.AccessKeys = ""
End If
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
'this function will add or remove a certain color
'quantity and return the result

Dim Red As Long, Blue As Long, Green As Long

If isXP = False Then
    Blue = ((Color \ &H10000) Mod &H100) + Value
Else
    Blue = ((Color \ &H10000) Mod &H100)
    Blue = Blue + ((Blue * Value) \ &HC0)
End If
Green = ((Color \ &H100) Mod &H100) + Value
Red = (Color And &HFF) + Value
    
    'check red
    If Red < 0 Then
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    'check green
    If Green < 0 Then
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    'check blue
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If

ShiftColor = RGB(Red, Green, Blue)
End Function

Private Sub CalcTextRects()
'this sub will calculate the rects required to draw the text
rc2.Left = 1: rc2.Right = Wi - 2: rc2.Top = 0: rc2.Bottom = He - 2
DrawText UserControl.hdc, elTex, Len(elTex), rc2, DT_CALCRECT Or DT_WORDBREAK
CopyRect rc, rc2: OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom) \ 2
CopyRect rc2, rc: OffsetRect rc2, 1, 1

End Sub

Public Sub Refresh()
    Call Redraw(lastStat, True)
End Sub
