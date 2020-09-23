VERSION 5.00
Begin VB.UserControl splitbar 
   BackColor       =   &H80000003&
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   ControlContainer=   -1  'True
   MousePointer    =   7  'Size N S
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ToolboxBitmap   =   "UserControl1.ctx":0000
End
Attribute VB_Name = "splitbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Dim ctltop As New Collection
Dim ctlbtm As New Collection
Dim userbackcolor As OLE_COLOR
Dim userborderstyle As Long
Private Const SW_MINIMIZE = 6
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
'Default Property Values:
Const m_def_MinTopHeight = 5
Const m_def_MinBottomHeight = 5
'Property Variables:
Dim m_MinTopHeight As Long
Dim m_MinBottomHeight As Long
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
'------------------------------------
' dim vars and initialize
Dim ctl As Control
Dim maxctltop As Long
Dim maxctlbtm As Long
'set mactlbtm
'get maximum top value
 For Each ctl In ctltop
    If maxctltop < ctl.Top Then maxctltop = ctl.Top
 Next
 'get maximum bottom value
  For Each ctl In ctlbtm
    If maxctlbtm = 0 Then
        maxctlbtm = ctl.Top + ctl.Height
    ElseIf maxctlbtm > ctl.Top + ctl.Height Then
        maxctlbtm = ctl.Top + ctl.Height
    End If
 Next
 '----------------------------------------------
 ' now update horizontal resize bar
UserControl.BorderStyle = 0
UserControl.BackColor = &H80000003
BringWindowToTop UserControl.hwnd
If Y + Extender.Top > maxctltop + MinTopHeight And Y + Extender.Top + Extender.Height < maxctlbtm - MinBottomHeight Then
' update splitbar
Extender.Top = Y + Extender.Top
End If
Exit Sub
bail: MsgBox Err.Description
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
update
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function update() As Variant
UserControl.BorderStyle = userborderstyle
UserControl.BackColor = userbackcolor
' update top controls
    For Each ctl In ctltop
       ' resize ctl
        ctl.Height = (Extender.Top) - ctl.Top
    Next
'update bottom controls
    For Each ctl In ctlbtm
        ' resize ctl
        ctl.Height = (ctl.Top + ctl.Height) - (Extender.Top + UserControl.ScaleHeight)
        ctl.Top = Extender.Top + UserControl.ScaleHeight
        ctl.Refresh
     Next
End Function



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5

Public Sub AddControlTop(ctl As Object)
On Error GoTo bail
     ctltop.Add ctl
Exit Sub
bail:
MsgBox Err.Description
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub AddControlBottom(ctl As Object)
On Error GoTo bail
     ctlbtm.Add ctl
Exit Sub
bail:
MsgBox Err.Description
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,10
'Public Property Get MinTopHeight() As Long
'    MinTopHeight = m_MinTopHeight
'End Property
'
'Public Property Let MinTopHeight(ByVal New_MinTopHeight As Long)
'    m_MinTopHeight = New_MinTopHeight
'    PropertyChanged "MinTopHeight"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_MinTopHeight = m_def_MinTopHeight
    m_MinTopHeight = m_def_MinTopHeight
    m_MinBottomHeight = m_def_MinBottomHeight
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000001)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
'    m_MinTopHeight = PropBag.ReadProperty("MinTopHeight", m_def_MinTopHeight)
    userbackcolor = UserControl.BackColor
    userborderstyle = UserControl.BorderStyle
    m_MinTopHeight = PropBag.ReadProperty("MinTopHeight", m_def_MinTopHeight)
    m_MinBottomHeight = PropBag.ReadProperty("MinBottomHeight", m_def_MinBottomHeight)
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000001)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
'    Call PropBag.WriteProperty("MinTopHeight", m_MinTopHeight, m_def_MinTopHeight)
    Call PropBag.WriteProperty("MinTopHeight", m_MinTopHeight, m_def_MinTopHeight)
    Call PropBag.WriteProperty("MinBottomHeight", m_MinBottomHeight, m_def_MinBottomHeight)
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,5
Public Property Get MinTopHeight() As Long
    MinTopHeight = m_MinTopHeight
End Property

Public Property Let MinTopHeight(ByVal New_MinTopHeight As Long)
    m_MinTopHeight = New_MinTopHeight
    PropertyChanged "MinTopHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,5
Public Property Get MinBottomHeight() As Long
    MinBottomHeight = m_MinBottomHeight
End Property

Public Property Let MinBottomHeight(ByVal New_MinBottomHeight As Long)
    m_MinBottomHeight = New_MinBottomHeight
    PropertyChanged "MinBottomHeight"
End Property

