VERSION 5.00
Begin VB.UserControl ypnSideMenu 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ypnSideMenu.ctx":0000
   Begin VB.PictureBox skinPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3600
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   2880
      Picture         =   "ypnSideMenu.ctx":0312
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "ypnSideMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const DT_CENTER         As Long = &H1
Private Const DT_VCENTER        As Long = &H4
Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_LEFT As Long = &H0
Private Const DT_RIGHT As Long = &H2

Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type
Enum sidemenu_skin
    skin0 = 0
    skin1 = 1
    skin2 = 2
    skin3 = 3
    skin4 = 4
    skin5 = 5
    skin6 = 6
    skin7 = 7
    skin8 = 8
    skin9 = 9
End Enum
Enum menuAlign
    左对齐 = 1
    居中 = 2
    右对齐 = 3
End Enum
Enum selIndex
    MenuID01 = 0
    MenuID02 = 1
    MenuID03 = 2
    MenuID04 = 3
    MenuID05 = 4
    MenuID06 = 5
    MenuID07 = 6
    MenuID08 = 7
    MenuID09 = 8
    MenuID10 = 9
    MenuID11 = 10
    MenuID12 = 11
    MenuID13 = 12
    MenuID14 = 13
    MenuID15 = 14
    MenuID16 = 15
    MenuID17 = 16
    MenuID18 = 17
    MenuID19 = 18
    MenuID20 = 19
    MenuID21 = 20
    MenuID22 = 21
    MenuID23 = 22
    MenuID24 = 23
    MenuID25 = 24
    MenuID26 = 25
    MenuID27 = 26
    MenuID28 = 27
    MenuID29 = 28
    MenuID30 = 29
End Enum
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private m_Caption As String
Private m_State As Long
Dim capture As Boolean
Dim ScaleWidth As Single, ScaleHeight As Single
Dim itemHeight As Single
Dim itemCount As Integer
Dim itemCaption() As String
Dim itemEnabled() As Boolean

Dim mTxtColor As OLE_COLOR, mBorderColor As OLE_COLOR
Dim alignData As Long, align As menuAlign

Dim skin As sidemenu_skin
Dim aim As Integer, oldAim As Integer '指向
Dim sel As Integer, oldsel As Integer '选定
Dim ready As Boolean
Public Event Click(itemIndex As Integer)   '声明Click事件  index从0开始

Private Function getItem(Y As Single) As Integer
    getItem = (Y - 2) \ itemHeight
    If getItem >= itemCount Then getItem = itemCount - 1
End Function
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'oldSel = Sel
        'Sel = getItem(Y)
        'Call drawMenuItem(oldSel)
        'Call drawMenuItem(Sel)
        Call switchSel(getItem(Y))
        RaiseEvent Click(sel)
    End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < ScaleWidth And Y > 0 And Y < ScaleHeight Then
        If Button = 1 Then Exit Sub
        If Not capture Then
            Call SetCapture(UserControl.hWnd)
            capture = True
        End If
        aim = getItem(Y)
        If (aim <> oldAim) Then
            Call drawMenuItem(oldAim)
            Call drawMenuItem(aim)
            oldAim = aim
        End If
    Else
        Call ReleaseCapture
        capture = False
        If aim > -1 Then
            Dim Temp As Integer
            Temp = aim
            aim = -1
            Call drawMenuItem(Temp)
        End If
        oldAim = -1
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    aim = getItem(Y)
    If capture Then
        Call ReleaseCapture
        capture = False
    End If
End Sub
Private Sub UserControl_Resize()
    Dim i As Integer
    If UserControl.ScaleWidth < 7 Then UserControl.Width = 105
    ScaleWidth = UserControl.ScaleWidth - 4
    ScaleHeight = itemHeight * itemCount
    UserControl.Height = (ScaleHeight + 4) * 15
    
    skinPicture.Width = ScaleWidth
    skinPicture.Height = itemHeight * 4
    skinPicture.Cls
    For i = 0 To 3
        skinPicture.PaintPicture Picture1.Image, 0, i * itemHeight, 3, 3, 7 * i, skin * 20, 3, 3, vbSrcCopy
        skinPicture.PaintPicture Picture1.Image, 0, i * itemHeight + 3, 3, itemHeight - 6, 7 * i, skin * 20 + 3, 3, 14, vbSrcCopy
        skinPicture.PaintPicture Picture1.Image, 0, i * itemHeight + itemHeight - 3, 3, 3, 7 * i, skin * 20 + 17, 3, 3, vbSrcCopy
        
        skinPicture.PaintPicture Picture1.Image, 3, i * itemHeight + 0, ScaleWidth - 6, 3, 7 * i + 3, skin * 20, 1, 3, vbSrcCopy
        skinPicture.PaintPicture Picture1.Image, 3, i * itemHeight + 3, ScaleWidth - 6, itemHeight - 6, 7 * i + 3, skin * 20 + 3, 1, 14, vbSrcCopy
        skinPicture.PaintPicture Picture1.Image, 3, i * itemHeight + itemHeight - 3, ScaleWidth - 6, 3, 7 * i + 3, skin * 20 + 17, 1, 3, vbSrcCopy
        
        skinPicture.PaintPicture Picture1.Image, ScaleWidth - 3, i * itemHeight + 0, 3, 3, 7 * i + 4, skin * 20, 3, 3, vbSrcCopy
        skinPicture.PaintPicture Picture1.Image, ScaleWidth - 3, i * itemHeight + 3, 3, itemHeight - 6, 7 * i + 4, skin * 20 + 3, 3, 14, vbSrcCopy
        skinPicture.PaintPicture Picture1.Image, ScaleWidth - 3, i * itemHeight + itemHeight - 3, 3, 3, 7 * i + 4, skin * 20 + 17, 3, 3, vbSrcCopy
    Next
    Call ReDrawMenu
End Sub
Private Sub ReDrawMenu()
    Dim i As Integer
    UserControl.Cls
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), mBorderColor, B
    For i = 0 To itemCount - 1
        drawMenuItem (i)
    Next
End Sub
Private Sub drawMenuItem(index As Integer)
    Dim Y As Single, rc As RECT
    If index < 0 Then Exit Sub
    Y = itemHeight * index + 2
    If Not UserControl.Enabled Or Not itemEnabled(index) Then
        UserControl.PaintPicture skinPicture.Image, 2, Y, UserControl.ScaleWidth, itemHeight, 0, itemHeight * 3, UserControl.ScaleWidth, itemHeight, vbSrcCopy
        Call SetRect(rc, 5, 5 + Y, UserControl.ScaleWidth - 3, itemHeight - 3 + Y)
    ElseIf sel = index Then
        UserControl.PaintPicture skinPicture.Image, 2, Y, UserControl.ScaleWidth, itemHeight, 0, itemHeight + itemHeight, UserControl.ScaleWidth, itemHeight, vbSrcCopy
        Call SetRect(rc, 5, 5 + Y, UserControl.ScaleWidth - 3, itemHeight - 3 + Y)
    ElseIf aim = index Then
        UserControl.PaintPicture skinPicture.Image, 2, Y, UserControl.ScaleWidth, itemHeight, 0, itemHeight, UserControl.ScaleWidth, itemHeight, vbSrcCopy
        Call SetRect(rc, 3, 3 + Y, UserControl.ScaleWidth - 3, itemHeight - 3 + Y)
    Else
        UserControl.PaintPicture skinPicture.Image, 2, Y, UserControl.ScaleWidth, itemHeight, 0, 0, UserControl.ScaleWidth, itemHeight, vbSrcCopy
        Call SetRect(rc, 3, 3 + Y, UserControl.ScaleWidth - 3, itemHeight - 3 + Y)
    End If
    
    If UserControl.Enabled Then
        UserControl.ForeColor = mTxtColor
    Else
        UserControl.ForeColor = GetSysColor(15)
    End If
    Call DrawText(UserControl.hDC, itemCaption(index), -1, rc, alignData)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)  '读取属性
    Dim i As Integer
    UserControl.Enabled = PropBag.ReadProperty("Enabled", UserControl.Enabled)
    skin = PropBag.ReadProperty("skin", 0)
    itemCount = PropBag.ReadProperty("itemCount", 1)
    itemHeight = PropBag.ReadProperty("itemHeight", 20)
    Font = PropBag.ReadProperty("Font", UserControl.Font)
    mTxtColor = PropBag.ReadProperty("txtColor", 0)
    mBorderColor = PropBag.ReadProperty("borderColor", &HDC861F)
    align = PropBag.ReadProperty("align", 2)
    ReDim itemCaption(itemCount - 1)
    ReDim itemEnabled(itemCount - 1)
    For i = 0 To itemCount - 1
        itemCaption(i) = PropBag.ReadProperty("Caption" + CStr(i), "菜单项")
        itemEnabled(i) = PropBag.ReadProperty("Enabled" + CStr(i), True)
    Next
    Call setAlign
    ready = True
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)  '写入属性
    Dim i As Integer
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled)
    Call PropBag.WriteProperty("skin", skin)
    Call PropBag.WriteProperty("itemCount", itemCount, 1)
    Call PropBag.WriteProperty("itemHeight", itemHeight, 20)
    Call PropBag.WriteProperty("Font", UserControl.Font)
    Call PropBag.WriteProperty("txtColor", mTxtColor, 0)
    Call PropBag.WriteProperty("borderColor", mBorderColor, &HDC861F)
    Call PropBag.WriteProperty("align", align, 2)
    For i = 0 To itemCount - 1
        Call PropBag.WriteProperty("Caption" + CStr(i), itemCaption(i))
        Call PropBag.WriteProperty("Enabled" + CStr(i), itemEnabled(i))
    Next
End Sub
Private Sub UserControl_Initialize()
    sel = 0
    oldsel = 0
    aim = -1
    skin = skin1
    mTxtColor = 0
    mBorderColor = &HDC861F
    itemHeight = 20
    itemCount = 4
    align = 居中
    alignData = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    ReDim itemCaption(itemCount - 1)
    ReDim itemEnabled(itemCount - 1)
    itemCaption(0) = "菜单项1"
    itemCaption(1) = "菜单项2"
    itemCaption(2) = "菜单项3"
    itemCaption(3) = "菜单项4"
    itemEnabled(0) = True
    itemEnabled(1) = True
    itemEnabled(2) = True
    itemEnabled(3) = True
End Sub
Private Sub setAlign()
    If align = 左对齐 Then
        alignData = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    ElseIf align = 居中 Then
        alignData = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Else
        alignData = DT_RIGHT Or DT_VCENTER Or DT_SINGLELINE
    End If
    Call ReDrawMenu
End Sub
Public Property Get menuHeight() As Single
    menuHeight = itemHeight
End Property
Public Property Let menuHeight(ByVal newValue As Single)  '设置可用状态
    If newValue < 20 Then newValue = 20
    itemHeight = newValue
    Call UserControl_Resize
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal newValue As Boolean) '设置可用状态
    UserControl.Enabled = newValue
    Call ReDrawMenu
End Property
Public Property Get Font() As StdFont  '返回字体
    Set Font = UserControl.Font
End Property
Public Property Let Font(ByVal newValue As StdFont)  '设置字体
    Set UserControl.Font = newValue
    If ready Then Call ReDrawMenu
End Property
Public Property Get menuItemCount() As Integer
    menuItemCount = itemCount
End Property
Public Property Let menuItemCount(ByVal newValue As Integer)
    Dim i As Integer
    If newValue > 0 And newValue <> itemCount Then
        ReDim Preserve itemCaption(newValue - 1)
        ReDim Preserve itemEnabled(newValue - 1)
        For i = itemCount To newValue - 1
            itemCaption(i) = "菜单项" & CStr(i + 1)
            itemEnabled(i) = True
        Next
        itemCount = newValue
        Call UserControl_Resize
    End If
End Property
Public Property Get activeItem() As selIndex        '当前项
    activeItem = sel
End Property
Public Property Let activeItem(ByVal newValue As selIndex)
    If newValue >= itemCount Then newValue = MenuID01
    If newValue >= 0 Then Call switchSel(CInt(newValue))
End Property
Public Property Get Caption() As String    '返回标题
    Caption = itemCaption(sel)
End Property
Public Property Let Caption(ByVal newValue As String) '设置标题
    itemCaption(sel) = newValue
    Call ReDrawMenu
End Property
Public Property Get menuSkin() As sidemenu_skin
    menuSkin = skin
End Property
Public Property Let menuSkin(ByVal newValue As sidemenu_skin)
    If skin <> newValue Then
        skin = newValue
        Call UserControl_Resize
    End If
End Property
Public Sub setItemEnabled(menuItem As Integer, newValue As Boolean)
    If menuItem >= 0 And menuItem < itemCount Then
        If itemEnabled(menuItem) <> newValue Then Call Switch(newValue)
    End If
End Sub
Public Property Get menuEnabled() As Boolean
    menuEnabled = itemEnabled(sel)
End Property
Public Property Let menuEnabled(ByVal newValue As Boolean)
    If itemEnabled(sel) <> newValue Then
        itemEnabled(sel) = newValue
        Call drawMenuItem(sel)
    End If
End Property
Public Property Get TxtColor() As OLE_COLOR '返回前景色
    TxtColor = mTxtColor
End Property
Public Property Let TxtColor(ByVal newValue As OLE_COLOR)   '设置前景色
    mTxtColor = newValue
    Call ReDrawMenu
End Property
Public Property Get BorderColor() As OLE_COLOR '返回前景色
    BorderColor = mBorderColor
End Property
Public Property Let BorderColor(ByVal newValue As OLE_COLOR)   '设置前景色
    mBorderColor = newValue
    Call ReDrawMenu
End Property
Public Property Get alignment() As menuAlign
    alignment = align
End Property
Public Property Let alignment(ByVal newValue As menuAlign)    '设置前景色
    If align <> newValue Then
        align = newValue
        Call setAlign
        If ready Then Call setAlign
    End If
End Property
Private Sub switchSel(newValue As Integer)
    If newValue >= 0 And newValue < itemCount Then
        sel = newValue
        Call drawMenuItem(oldsel)
        oldsel = sel
        Call drawMenuItem(sel)
    End If
End Sub
