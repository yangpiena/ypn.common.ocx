VERSION 5.00
Begin VB.UserControl ypnPropertySheet 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
   ToolboxBitmap   =   "ypnPropertySheet.ctx":0000
   Begin VB.VScrollBar VScroll1 
      Height          =   4095
      Left            =   3360
      Max             =   0
      SmallChange     =   50
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   15
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   1800
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2160
      Picture         =   "ypnPropertySheet.ctx":0312
      ScaleHeight     =   240
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox ico 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   2880
      Picture         =   "ypnPropertySheet.ctx":0BD4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox ico 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   2880
      Picture         =   "ypnPropertySheet.ctx":115E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   7
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "ypnPropertySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetRect Lib "user32 " (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          'Aki
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const DI_MASK = &H1                          ' ��ͼʱʹ��ͼ���MASK���� (�絥��ʹ��, �ɻ��ͼ�����ģ)
Private Const DI_IMAGE = &H2                         ' ��ͼʱʹ��ͼ���XOR���� (��ͼ��û��͸������)
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE        ' �ó��淽ʽ��ͼ
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_TOP = &H0
Private Const DT_WORDBREAK = &H10

'���ߣ�ִ��
'QQ��47815463
'��Ȩ����������⸴�ơ�ʹ�á��޸ı��ؼ��������뱣��������Ϣ��
'������£�
'15.5.23���޸�vscroll��˸BUG
'16.4.1������mouseDown�¼�
'16.4.18������delItems��������change�¼���ʹ�á�
Enum ItemTypeData
    iString = 0
    iInteger = 1
    iLong = 2
    iSingle = 3
    iBoolean = 4
    iList = 5
End Enum

Private Type itemClass
    itemName As String
    itemType As ItemTypeData '��ǰ�� ��������
    itemValue As String     'ֵ
    itemIntegerValue As Integer
    itemLongValue As Long
    itemSingleValue As Single
    itemBooleanValue As Boolean
    itemListIndex As Integer '��ǰ�б�������type=itemListʱ��Ч
    itemList() As String    '0����
    itemMax As Single       '���ֵ
    itemMin As Single       '��Сֵ��������type=index/long/singl, max<minʱ��Ч
    ItemEnabled As Boolean
    itemReadWrite As Boolean
    itemDescription As String '����
End Type

Private Type sheetClass
    sheetName As String
    sheetItemCount As Integer
    sheetItem() As itemClass
    sheetExpand As Boolean 'չ��
    sheetDescription As String '����
    sheetTop As Single '��ǰλ��
End Type

Private Type propertySheetData
    sheetIndex As Integer
    itemIndex As Integer
    itemReadWrite As Boolean
    ItemEnabled As Boolean
End Type

Private Type POINT
    X As Long
    Y As Long
End Type
Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Dim PropertySheet() As sheetClass, PSCount As Integer
Dim psList() As propertySheetData, psListCount As Integer

Dim sheetWidth As Single, sheetHeight As Single
Dim NameHeight As Single
Dim NameWidth As Single
Dim SheetHeadColor As OLE_COLOR      '��ͷ��ɫ
Dim TableColor As OLE_COLOR   '�����ɫ
Dim TableBackColor1 As OLE_COLOR      '����ɫ1
Dim TableBackColor2 As OLE_COLOR      '����ɫ2
Dim TxtColor As OLE_COLOR       '������ɫ
Dim invalidColor As Long
Dim txtHotColor As OLE_COLOR    '�ȸ���������ɫ
Dim highLightColor As OLE_COLOR     '��ǰ����ɫ
Dim DescriptionVisible As Boolean    '�Ƿ���ʾ����
Dim DescriptionHeight As Single     '�߶�
Dim DescriptionRect As RECT
Dim DescriptionText As String
Dim AutoRefresh As Boolean '�Զ�ˢ��

Dim currentSheetIndex As Integer, currentItemIndex As Integer '��ͱ����Ǵ�1��ʼ�ģ�0����
Dim hotIndex As Integer, oldHotIndex As Integer '�ȸ�����Ŀ���ϴ��ȸ�����Ŀ psList��
Dim highLightIndex As Integer, oldHighlightIndex As Integer '��ǰ��Ŀ���ϴ���Ŀ psList��
Dim scrollMax As Single, ScrollValue As Single, ScrollMouse As Single, rightMove As Boolean '��,rightMove�Ҽ��϶�
Dim editing As Boolean '�༭��
Dim currentTextType As ItemTypeData
Dim comboState As Integer 'combobox�¼�ͷ����״̬��0����ʾ,1��ͨ,2׼������,3�ȸ���,4����
Dim modified As Boolean '�иı�
Dim notChangeCombo As Boolean 'combo��ʼ�����ģ�������click

Public Event Change(sheet As Integer, Item As Integer, newValue As String)
Public Event itemClick(sheet As Integer, Item As Integer)
Public Event itemDBClick(sheet As Integer, Item As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Scroll()

Private Sub UserControl_Initialize()
    NameHeight = 16
    NameWidth = 90
    TableColor = &HFF8080
    SheetHeadColor = &HDCC887
    TableBackColor1 = &HFFF7EC
    TableBackColor2 = &HC0FFFF
    TxtColor = 0
    invalidColor = &H4D66 ' GetSysColor(COLOR_INACTIVECAPTIONTEXT)
    txtHotColor = &HFF0000
    highLightColor = &H80FF&
    DescriptionVisible = True
    DescriptionHeight = 50
    AutoRefresh = True
    
    currentSheetIndex = -1
    currentItemIndex = -1           'current�Ǵ�0��ʼ��
    highLightIndex = -1
    oldHighlightIndex = -1
    hotIndex = -1           '�ȸ�����Чֵ��-1����ǰ��Ŀû����Ч��������տ�ʼ��-1���⣬
    oldHotIndex = -1
    ScrollValue = 0
    
    ReDim PropertySheet(0)
    ReDim PropertySheet(0).sheetItem(0)
    ReDim psList(0)
    PSCount = 0
    psListCount = 0
    Call AddSheet("��ͷ", True)
    currentSheetIndex = 1
    Call AddItem("����1", iString, "�ı�")
    Call AddItem("����2", iSingle, "3.1415926")
    SetComboHeight Combo1, 400
End Sub
Public Sub SetComboHeight(oComboBox As Object, lNewHeight As Long)
    If TypeOf oComboBox.Parent Is Frame Then Exit Sub
    MoveWindow oComboBox.hwnd, oComboBox.Left, oComboBox.Top, oComboBox.Width, lNewHeight, 1
End Sub
Private Sub Timer1_Timer()
    Dim dot As POINT
    Call GetCursorPos(dot)
    ScreenToClient Picture3.hwnd, dot
    If dot.X < 0 Or dot.Y < 0 Or dot.X > UserControl.ScaleWidth Or dot.Y > UserControl.ScaleHeight Then
            Timer1.Enabled = False
            Call eraseOldHot
            comboState = 0
            Call refreshTable
            hotIndex = -1
            oldHotIndex = -1
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then  'enter
        If editing Then Call txtChange Else Text1.Visible = False
    Else
        editing = True
    End If
    If KeyAscii <> 8 Then
        If currentTextType = iInteger Or currentTextType = iLong Then
            If KeyAscii = 46 Then KeyAscii = 0
            If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 45 Then KeyAscii = 0 '           1~10=48~59,�˸�=8,-45,��46
        ElseIf currentTextType = iSingle Then
            If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 45 And (KeyAscii <> 46 Or InStr(Text1.Text, ".") > 0) Then KeyAscii = 0
        End If
    End If
    modified = True
End Sub
Public Sub UserControl_ExitFocus() '��ʱʧȥ���㲻�������¼���ˢ������ʱ��ǰsheet��1���´���
    If editing Then Call txtChange
    Text1.Visible = False
End Sub
Private Sub Combo1_Click()
    Dim Top As Single
    If notChangeCombo Then notChangeCombo = False: Exit Sub
    Top = highLightIndex * NameHeight
    With PropertySheet(currentSheetIndex).sheetItem(currentItemIndex)
        .itemListIndex = Combo1.ListIndex
        .itemValue = Combo1.List(Combo1.ListIndex)  'ֵ��0��ʼ
        If .itemType = iBoolean Then .itemBooleanValue = .itemValue = "True"
    End With
    Picture1.Line (NameWidth + 2, Top + 1)-(Picture1.Width - 2, Top + NameHeight - 1), IIf((psList(highLightIndex).itemIndex And &H1) = 1, TableBackColor1, TableBackColor2), BF
    Call refreshHighlight(Top, highLightIndex)
    Call drawComboCommand(1)
    RaiseEvent Change(currentSheetIndex, currentItemIndex, Combo1.List(Combo1.ListIndex))
    editing = False
    modified = True
End Sub
Private Sub picture3_DblClick()
    If highLightIndex = -1 Then Exit Sub
    If Not psList(highLightIndex).ItemEnabled Then Exit Sub
    RaiseEvent itemDBClick(currentSheetIndex, currentItemIndex)
    If Not psList(highLightIndex).itemReadWrite Then Exit Sub
    If editing Then Call txtChange Else Text1.Visible = False
    If currentItemIndex = 0 Then
        With PropertySheet(currentSheetIndex)
            .sheetExpand = Not .sheetExpand
            oldHotIndex = -1
            Call paintSheet
            Call refreshHighlight(highLightIndex * NameHeight, highLightIndex)
            Call refreshTable
        End With
    ElseIf currentTextType = iList Or currentTextType = iBoolean Then
        Combo1.Left = NameWidth + 1
        Combo1.Width = sheetWidth - NameWidth - 1
        Combo1.Top = highLightIndex * NameHeight - ScrollValue - 2         '�߶�20
        SendMessage Combo1.hwnd, &H14F, 1, 0
        editing = True
        notChangeCombo = False
    Else
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Text1.Visible = True
        Text1.SetFocus
    End If
End Sub
Private Sub txtChange()
    If highLightIndex = -1 Or highLightIndex > psListCount Then Exit Sub
    If Not psList(highLightIndex).itemReadWrite Then Exit Sub
    With PropertySheet(currentSheetIndex).sheetItem(currentItemIndex)
        If .itemType < iBoolean And .itemValue <> Text1.Text Then
            .itemValue = Text1.Text
            If .itemType <> iString And .itemType <> iList And .itemMax > .itemMin Then
                If CSng(.itemValue) < .itemMin Then .itemValue = CStr(.itemMin)
                If CSng(.itemValue) > .itemMax Then .itemValue = CStr(.itemMax)
            End If
            If .itemType = iInteger Then
                .itemValue = cornerInt(.itemValue)
                .itemIntegerValue = CInt(Val(.itemValue))
            ElseIf .itemType = iLong Then
                .itemValue = cornerLng(.itemValue)
                .itemLongValue = CLng(Val(.itemValue))
            ElseIf .itemType = iSingle Then
                .itemValue = cornerSng(.itemValue)
                .itemSingleValue = CSng(Val(.itemValue))
            End If
            RaiseEvent Change(currentSheetIndex, currentItemIndex, .itemValue)
'            Call paintTable(highLightIndex)
 '           Call paintItemText(highLightIndex * NameHeight, PropertySheet(currentSheetIndex).sheetItem(currentItemIndex), txtColor)
            Call refreshHighlight(highLightIndex * NameHeight, highLightIndex)
            Call refreshTable
        End If
    End With
    Text1.Visible = False
    editing = False
End Sub


Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Top As Single, Change As Boolean, tempItem As itemClass, i As Integer
    If editing Then Call txtChange
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 2 Then  '�Ҽ�
        ScrollMouse = Y
        UserControl.MousePointer = 5
        rightMove = False
        Exit Sub
    End If
    oldHighlightIndex = highLightIndex
    highLightIndex = (ScrollValue + Y) \ NameHeight
    If highLightIndex > psListCount Or highLightIndex < 0 Then 'û�и���
        highLightIndex = -1
        Text1.Visible = False
        Call paintTable(oldHighlightIndex)
        Call refreshTable
        Exit Sub
    End If
    If Not psList(highLightIndex).ItemEnabled Then
        highLightIndex = -1
        Text1.Visible = False
        Call paintTable(oldHighlightIndex)
        Call refreshTable
        Exit Sub
    End If
    Top = highLightIndex * NameHeight
    currentSheetIndex = psList(highLightIndex).sheetIndex
    currentItemIndex = psList(highLightIndex).itemIndex
    If oldHighlightIndex <> highLightIndex Then
        Text1.Visible = False
        Combo1.Visible = False
        Call paintTable(oldHighlightIndex)
        Call refreshHighlight(Top, highLightIndex)
    ElseIf PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemType = iBoolean Or PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemType = iList Then
        Call drawComboCommand(3)
    End If
    DescriptionText = PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemDescription
    If currentItemIndex > 0 Then
        If psList(highLightIndex).itemReadWrite Then
            tempItem = PropertySheet(currentSheetIndex).sheetItem(currentItemIndex)
            currentTextType = tempItem.itemType
            Combo1.Clear
            If currentTextType = iBoolean Then
                Combo1.AddItem "True"
                Combo1.AddItem "False"
                notChangeCombo = True
                Combo1.ListIndex = IIf(LCase(tempItem.itemValue) = "true", 0, 1)
                Call drawComboCommand(1)
            ElseIf currentTextType = iList Then
                For i = 1 To UBound(tempItem.itemList)
                    Combo1.AddItem tempItem.itemList(i)
                Next
                notChangeCombo = Combo1.ListIndex = -1
                Combo1.ListIndex = tempItem.itemListIndex
                Call drawComboCommand(1)
            Else
                Text1.Text = tempItem.itemValue
                Text1.BackColor = IIf((currentItemIndex And &H1) = 1, TableBackColor1, TableBackColor2)
                Text1.Left = NameWidth + 2
                Text1.Top = Top - ScrollValue + 1
                Text1.Width = Picture1.Width - NameWidth - 2
                Text1.Height = NameHeight - 1
            End If
        Else
            Text1.Visible = False
        End If
        RaiseEvent itemClick(currentSheetIndex, currentItemIndex)
        Call refreshTable
    Else
        Text1.Visible = False
    End If

End Sub
Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If rightMove Then
            VScroll1.Value = VScroll1.Value + ScrollMouse - Y
            ScrollMouse = Y
        Else '�Ҽ�չ��������
            Dim tempIndex As Integer
            If editing Then Call txtChange Else Text1.Visible = False
            tempIndex = (ScrollValue + Y) \ NameHeight
            If tempIndex <= psListCount And tempIndex >= 0 Then
                With PropertySheet(psList(tempIndex).sheetIndex)
                    .sheetExpand = Not .sheetExpand
                    Call paintSheet
                    Call refreshHighlight(highLightIndex * NameHeight, highLightIndex)
                    Call refreshTable
                End With
            End If
        End If
        UserControl.MousePointer = 0
        Exit Sub
    End If

    If highLightIndex = -1 Then Exit Sub
    If comboState = 3 And psList(highLightIndex).itemReadWrite And psList(highLightIndex).ItemEnabled Then
        Combo1.Left = NameWidth + 1
        Combo1.Width = sheetWidth - NameWidth - 1
        Combo1.Top = highLightIndex * NameHeight - ScrollValue - 2         '�߶�20
        SendMessage Combo1.hwnd, &H14F, 1, CLng(0)
        Call drawComboCommand(1)
    End If
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Top As Single
    If Button = 2 Then
        If editing Then Call txtChange Else Text1.Visible = False
        ScrollValue = ScrollValue + ScrollMouse - Y
        ScrollMouse = Y
        If ScrollValue < 0 Then ScrollValue = 0
        If ScrollValue > VScroll1.Max Then ScrollValue = VScroll1.Max
        VScroll1.Value = ScrollValue
        rightMove = True
        Exit Sub
    End If
    hotIndex = (ScrollValue + Y) \ NameHeight
    If hotIndex > psListCount Or hotIndex < 0 Or X < 0 Or X > UserControl.ScaleWidth Then
        hotIndex = -1
    ElseIf Not psList(hotIndex).ItemEnabled Then
        hotIndex = -1
    End If
    
    If oldHotIndex <> hotIndex Then
        Call eraseOldHot
        oldHotIndex = hotIndex
        
        If hotIndex <> -1 Then
            With PropertySheet(psList(hotIndex).sheetIndex)
                Timer1.Enabled = True
                Top = hotIndex * NameHeight
                If psList(hotIndex).itemIndex = 0 Then  '��ͷ
                    Call paintSheetText(Top, .sheetName, txtHotColor)
                Else
                    Call paintItemText(Top, .sheetItem(psList(hotIndex).itemIndex), txtHotColor)
                End If
            End With
        End If
        Call refreshTable
    End If
End Sub
Private Sub vscroll1_Change()
    ScrollValue = VScroll1.Value
    Call refreshTable
End Sub
Public Sub setScroll(nValue As Integer) '�ⲿ���ƹ���
    If nValue > 0 Then
        If VScroll1.Value + 20 > VScroll1.Max Then
            VScroll1.Value = VScroll1.Max
        Else
            VScroll1.Value = VScroll1.Value + 20
        End If
    Else
        If VScroll1.Value - 20 < VScroll1.Min Then
            VScroll1.Value = VScroll1.Min
        Else
            VScroll1.Value = VScroll1.Value - 20
        End If
    End If
End Sub
Private Sub drawComboCommand(newState As Integer)  '0����ʾ,1��ͨ,2�ȸ���,3����,4׼������
    comboState = newState
    Picture1.PaintPicture Picture2.Image, sheetWidth - 15, highLightIndex * NameHeight + 1, 15, NameHeight - 1, comboState * 15 - 15, 0, 15, 16, vbSrcCopy
    Picture1.PSet (sheetWidth - 15, highLightIndex * NameHeight + 1), highLightColor
    Picture1.PSet (sheetWidth - 1, highLightIndex * NameHeight + 1), highLightColor
    Picture1.PSet (sheetWidth - 15, highLightIndex * NameHeight + 15), highLightColor
    Picture1.PSet (sheetWidth - 1, highLightIndex * NameHeight + 15), highLightColor
    'Call refreshTable
End Sub
Private Sub refreshHighlight(Top As Single, tempHighlight As Integer) '�ػ�������ʾ
    If tempHighlight <> -1 And tempHighlight < psListCount Then
        With PropertySheet(psList(tempHighlight).sheetIndex)
            If psList(tempHighlight).itemIndex = 0 Then             '��ͷ
                Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), highLightColor, BF
                DrawIconEx Picture1.hDC, 2, Top + (NameHeight - 13) / 2, ico(IIf(.sheetExpand, 0, 1)).Picture, 13, 13, 0, 0, DI_NORMAL
                Call paintSheetText(Top, .sheetName, TxtColor)
            Else
                Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), highLightColor, BF
                Picture1.Line (NameWidth + 1, Top)-(NameWidth + 1, Top + NameHeight + 1), TableColor '����
                Call paintItemText(Top, .sheetItem(psList(tempHighlight).itemIndex), TxtColor)
            End If
            Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), TableColor, B
            Picture1.Refresh
        End With
    End If
End Sub
Private Sub eraseOldHot()      '�������ȸ���
    Dim Top As Single
    If oldHotIndex <> -1 And oldHotIndex <= psListCount Then
        Top = oldHotIndex * NameHeight
        If psList(oldHotIndex).itemIndex = 0 Then '��ͷ
            Call paintSheetText(Top, PropertySheet(psList(oldHotIndex).sheetIndex).sheetName, TxtColor)
        Else
            Call paintItemText(Top, PropertySheet(psList(oldHotIndex).sheetIndex).sheetItem(psList(oldHotIndex).itemIndex), TxtColor)
        End If
    End If
End Sub
Private Sub paintTable(tempIndex As Integer) '��Ĭ����ɫ����/�ػ�ĳ���   pslist(tempIndex)
    Dim Top As Single
    If tempIndex <> -1 And tempIndex <= psListCount Then
        With PropertySheet(psList(tempIndex).sheetIndex)
            Top = tempIndex * NameHeight
            If psList(tempIndex).itemIndex = 0 Then '��ͷ
                Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), SheetHeadColor, BF
                Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), TableColor, B
                DrawIconEx Picture1.hDC, 2, Top + (NameHeight - 13) / 2, ico(IIf(.sheetExpand, 0, 1)).Picture, 13, 13, 0, 0, DI_NORMAL
                Call paintSheetText(Top, .sheetName, TxtColor)
            Else
                Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), IIf((psList(tempIndex).itemIndex And &H1) = 1, TableBackColor1, TableBackColor2), BF
                Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), TableColor, B
                Picture1.Line (NameWidth + 1, Top)-(NameWidth + 1, Top + NameHeight + 1), TableColor '����
                Call paintItemText(Top, .sheetItem(psList(tempIndex).itemIndex), TxtColor)
            End If
            Picture1.Refresh
        End With
    End If
End Sub
Private Sub UserControl_Resize()
    If UserControl.ScaleWidth < 19 Then UserControl.ScaleWidth = 19
    Picture3.Width = UserControl.ScaleWidth
    Picture3.Height = UserControl.ScaleHeight
    sheetWidth = UserControl.ScaleWidth - 19
    If DescriptionVisible Then
        If UserControl.ScaleHeight < DescriptionHeight Then
            UserControl.Height = (DescriptionHeight + 15) * 15
            Picture3.Height = UserControl.ScaleHeight
            Exit Sub
        End If
        sheetHeight = UserControl.ScaleHeight - DescriptionHeight
        DescriptionRect.Left = 2
        DescriptionRect.Top = CLng(sheetHeight) + 2
        DescriptionRect.Right = CLng(UserControl.ScaleWidth) - 2
        DescriptionRect.Bottom = CLng(UserControl.ScaleHeight) - 2
    Else
        If UserControl.ScaleHeight < 20 Then UserControl.Height = 20 * 15
        sheetHeight = UserControl.ScaleHeight
    End If
    If NameWidth < UserControl.Width / 15 * 0.1 Or NameWidth > UserControl.Width / 15 * 0.9 Then NameWidth = UserControl.Width / 15 * 0.5
    Picture1.Width = sheetWidth
    Picture1.Height = sheetHeight

    VScroll1.Left = sheetWidth + 1
    VScroll1.Height = sheetHeight - 1
    Call paintSheet
End Sub
Private Sub paintSheetText(tempTop As Single, tempName As String, tempColor As OLE_COLOR)            '�ػ�����
    Dim rc As RECT
    Call SetTextColor(Picture1.hDC, tempColor)
    Call SetRect(rc, 17, tempTop + 2, sheetWidth, tempTop + NameHeight - 2)
    Call DrawText(Picture1.hDC, tempName, -1, rc, DT_LEFT Or DT_SINGLELINE)
End Sub
Private Sub paintItemText(tempTop As Single, tempItem As itemClass, tempColor As OLE_COLOR)
    Dim rc As RECT
    Call SetTextColor(Picture1.hDC, IIf(tempItem.ItemEnabled, tempColor, invalidColor))  '��Ч�ı���ɫ
    Call SetRect(rc, 3, tempTop + 2, NameWidth - 2, tempTop + NameHeight - 2)
    Call DrawText(Picture1.hDC, tempItem.itemName, -1, rc, DT_LEFT Or DT_SINGLELINE)
    
    Call SetTextColor(Picture1.hDC, IIf(tempItem.ItemEnabled And tempItem.itemReadWrite, tempColor, invalidColor))   '��Ч�ı���ɫ
    Call SetRect(rc, NameWidth + 3, tempTop + 2, Picture1.Width - 2, tempTop + NameHeight - 2)
    Call DrawText(Picture1.hDC, tempItem.itemValue, -1, rc, DT_LEFT Or DT_SINGLELINE)
End Sub
Private Sub paintSheet()        '�ػ�������
    Dim i As Integer, j As Integer
    Dim Row As Integer, Top As Single '��ǰ��������,��ǰ����λ��
    Dim tableEven As Boolean
    Row = 0
    Picture1.Cls
    Picture1.Height = PSCount * NameHeight + 1
    For i = 1 To PSCount
        With PropertySheet(i)
            Top = Row * NameHeight
            Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), SheetHeadColor, BF
            DrawIconEx Picture1.hDC, 2, Top + (NameHeight - 13) / 2, ico(IIf(.sheetExpand, 0, 1)).Picture, 13, 13, 0, 0, DI_NORMAL
            Call paintSheetText(Top, .sheetName, TxtColor)
            
            setPs Row, i
            Row = Row + 1
            If .sheetExpand Then
                tableEven = True
                Picture1.Height = Picture1.Height + .sheetItemCount * NameHeight
                For j = 1 To .sheetItemCount
                    Top = Row * NameHeight
                    Picture1.Line (0, Top)-(sheetWidth, Top + NameHeight), IIf(tableEven, TableBackColor1, TableBackColor2), BF
                    Picture1.Line (NameWidth + 1, Top)-(NameWidth + 1, (Row + 1) * NameHeight + 1), TableColor '����
                    Call paintItemText(Top, .sheetItem(j), TxtColor)
                    
                    setPs Row, i, j, .sheetItem(j).ItemEnabled, .sheetItem(j).itemReadWrite
                    tableEven = Not tableEven
                    Row = Row + 1
                Next
            End If
            Picture1.Refresh
        End With
    Next
    psListCount = Row
    scrollMax = Picture1.Height - sheetHeight - 1
    If scrollMax <= 0 Then scrollMax = 0
    If ScrollValue > scrollMax Then ScrollValue = scrollMax
    VScroll1.Max = scrollMax
    ReDim Preserve psList(psListCount)
    For i = 0 To psListCount
        Picture1.Line (0, i * NameHeight)-(Picture1.Width, i * NameHeight), TableColor
    Next
    Call refreshHighlight(highLightIndex * NameHeight, highLightIndex)
    Call refreshTable
End Sub
Private Sub refreshTable()
    Picture3.Cls
    Call BitBlt(Picture3.hDC, 0, 0, sheetWidth, sheetHeight, Picture1.hDC, 0, ScrollValue, vbSrcCopy)
    
    Call DrawText(Picture3.hDC, DescriptionText, -1, DescriptionRect, DT_LEFT Or DT_TOP Or DT_WORDBREAK)
    
    Picture3.Line (0, 0)-(sheetWidth, sheetHeight), TableColor, B
    Picture3.Line (0, sheetHeight)-(UserControl.ScaleWidth - 1, sheetHeight), TableColor, B
    Picture3.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), TableColor, B
    Picture3.Refresh
End Sub
Private Sub setPs(Index As Integer, sheetIndex As Integer, Optional itemIndex As Integer = 0, Optional ItemEnabled As Boolean = True, Optional itemReadWrite As Boolean = True) '������ʾ�б��������۵����ص���Ŀ
    If Index > UBound(psList) Then ReDim Preserve psList(Index + 100)
    psList(Index).sheetIndex = sheetIndex
    psList(Index).itemIndex = itemIndex
    psList(Index).ItemEnabled = ItemEnabled
    psList(Index).itemReadWrite = itemReadWrite
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------����
Public Sub AddSheet(sheetName As String, Optional expand As Boolean = True, Optional Description As String = "") '�ŵ����
    PSCount = PSCount + 1
    ReDim Preserve PropertySheet(PSCount)
    With PropertySheet(PSCount)
        .sheetName = sheetName
        .sheetItemCount = 0
        ReDim .sheetItem(0)
        ReDim .sheetItem(0).itemList(0)
        .sheetExpand = expand
        currentSheetIndex = PSCount
    End With
    If AutoRefresh Then Call paintSheet
End Sub
Public Sub DelCurrentSheet()
    Dim i As Integer
    For i = currentSheetIndex To PSCount - 1
        PropertySheet(i) = PropertySheet(i + 1)
    Next
    PSCount = PSCount - 1
    ReDim Preserve PropertySheet(PSCount)
    If currentSheetIndex > PSCount Then currentSheetIndex = PSCount
    If AutoRefresh Then Call paintSheet
End Sub
Public Sub delItems(tempSheetIndex As Integer)
    Dim i As Integer
    If tempSheetIndex < 0 Or tempSheetIndex > PSCount Then Exit Sub

    currentSheetIndex = tempSheetIndex
    With PropertySheet(tempSheetIndex)
        .sheetItemCount = 0
        Erase .sheetItem
        ReDim .sheetItem(0)
        ReDim .sheetItem(0).itemList(0)
    End With
    oldHotIndex = -1
    hotIndex = -1
    oldHighlightIndex = -1
    highLightIndex = -1
    Call paintSheet
End Sub
Public Sub AddItem(itemName As String, itemType As ItemTypeData, itemValue As String, Optional ReadWrite As Boolean = True, Optional itemDescription As String = "��") '��ӵ���ǰsheet���item
    If currentSheetIndex < 0 Or currentSheetIndex > PSCount Then Exit Sub
    With PropertySheet(currentSheetIndex)
        .sheetItemCount = .sheetItemCount + 1
        ReDim Preserve .sheetItem(.sheetItemCount)
        ReDim .sheetItem(.sheetItemCount).itemList(0)
        With PropertySheet(currentSheetIndex).sheetItem(PropertySheet(currentSheetIndex).sheetItemCount)
            .itemName = itemName
            .itemType = itemType
            .itemReadWrite = ReadWrite
            .ItemEnabled = True
            .itemDescription = "˵��:" & itemDescription
            ReDim .itemList(0)
            If .itemType = iList Then
                Call setListIndex(currentSheetIndex, PropertySheet(currentSheetIndex).sheetItemCount, -1)
            Else
                Call setNewValue(currentSheetIndex, PropertySheet(currentSheetIndex).sheetItemCount, itemValue)
            End If
        End With
        currentItemIndex = .sheetItemCount
    End With
    If AutoRefresh Then Call paintSheet
End Sub
Public Sub DelCurrentItem()
    Dim i As Integer
    With PropertySheet(currentSheetIndex)
        For i = currentItemIndex To .sheetItemCount - 1
            .sheetItem(i) = .sheetItem(i + 1)
        Next
        .sheetItemCount = .sheetItemCount - 1
        ReDim Preserve .sheetItem(.sheetItemCount)
        If currentItemIndex > .sheetItemCount Then currentItemIndex = .sheetItemCount
    End With
    If AutoRefresh Then Call paintSheet
End Sub
Public Sub AddListText(tempSheetIndex As Integer, tempItemIndex As Integer, Text As String)         '����б�
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        If PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemType = iList Then
            With PropertySheet(tempSheetIndex).sheetItem(tempItemIndex)
                ReDim Preserve .itemList(UBound(.itemList) + 1)
                .itemList(UBound(.itemList)) = Text
                .itemListIndex = 0
            End With
        End If
    End If
End Sub
Public Sub clearListText(tempSheetIndex As Integer, tempItemIndex As Integer)
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        If PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemType = iList Then
            With PropertySheet(tempSheetIndex).sheetItem(tempItemIndex)
                ReDim Preserve .itemList(0)
                .itemList(0) = ""
            End With
        End If
    End If
End Sub
Public Sub ModifyListText(tempSheetIndex As Integer, tempItemIndex As Integer, ListIndex As Integer, Text As String)         '�޸��б���
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        If PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemType = iList Then
            With PropertySheet(tempSheetIndex).sheetItem(tempItemIndex)
                If ListIndex >= 0 And ListIndex <= UBound(.itemList) Then
                    .itemList(ListIndex) = Text
                End If
            End With
        End If
    End If
End Sub
Public Sub setItemEnabled(tempSheetIndex As Integer, tempItemIndex As Integer, Optional Enabled As Boolean = True, Optional ReadWrite As Boolean = True) '��ѡ�񣬶�д���޸�
    Dim Row As Integer, Top As Single '��ǰ��������,��ǰ����λ��
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).ItemEnabled = Enabled
        PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemReadWrite = ReadWrite
        If AutoRefresh Then Call paintSheet
    End If
End Sub
Public Sub Clear()
    ReDim PropertySheet(0)
    ReDim PropertySheet(0).sheetItem(0)
    ReDim psList(0)
    PSCount = 0
    psListCount = 0
    Call paintSheet
End Sub
'---------------------------------------------------------------------------------����
Public Property Get theModified() As Boolean    'ÿ�λ�ȡʱ�Զ��÷�
    theModified = modified
    modified = False
End Property
Public Function getSheetCount() As Integer
    getSheetCount = PSCount
End Function
Public Function getItemCount() As Integer
    getItemCount = PropertySheet(currentSheetIndex).sheetItemCount
End Function
Public Property Get currentSheet() As Integer
    currentSheet = currentSheetIndex
End Property
Public Property Let currentSheet(ByVal vNewValue As Integer)
    If vNewValue > PSCount Or vNewValue < 1 Then Exit Property
    currentSheetIndex = vNewValue
End Property
Public Property Get currentItem() As Integer
    currentItem = currentItemIndex
End Property
Public Property Let currentItem(ByVal vNewValue As Integer)
    With PropertySheet(currentSheetIndex)
        If vNewValue < 1 Or vNewValue > .sheetItemCount Then Exit Property
        currentItemIndex = vNewValue
    End With
End Property
Public Property Get sheetName() As String
    If currentSheetIndex = -1 Then Exit Property
    sheetName = PropertySheet(currentSheetIndex).sheetName
End Property
Public Property Let sheetName(ByVal vNewValue As String)
    PropertySheet(currentSheetIndex).sheetName = vNewValue
    PropertyChanged "pscount"
    If AutoRefresh Then Call paintSheet
End Property
Public Property Get itemName() As String
    If currentSheetIndex = -1 Or currentItemIndex = -1 Then Exit Property
    itemName = PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemName
End Property
Public Property Let itemName(ByVal vNewValue As String)
    PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemName = vNewValue
    If AutoRefresh Then Call paintSheet
End Property
Public Sub setNumberRange(tempSheetIndex As Integer, tempItemIndex As Integer, tempMaX As Single, tempMin As Single)
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        With PropertySheet(tempSheetIndex).sheetItem(tempItemIndex)
        If .itemType = iInteger Or .itemType = iLong Or .itemType = iSingle Then
            .itemMax = tempMaX
            .itemMin = tempMin
        End If
        End With
    End If
End Sub
Public Property Get itemListIndex() As Integer
    If currentSheetIndex = -1 Or currentItemIndex = -1 Then Exit Property
    itemListIndex = PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemListIndex
End Property
Public Property Get itemValue() As String
    If currentSheetIndex = -1 Or currentItemIndex = -1 Then Exit Property
    itemValue = PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemValue
End Property
Public Property Let itemValue(ByVal vNewValue As String)
    If currentSheetIndex = -1 Or currentItemIndex = -1 Then Exit Property
    PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemValue = vNewValue
    If AutoRefresh Then Call paintSheet
End Property
Public Function setValue(ByVal tempSheetIndex As Integer, ByVal tempItemIndex As Integer, ByVal vNewValue As String)
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        With PropertySheet(tempSheetIndex).sheetItem(tempItemIndex)
            If .itemType = iInteger Or .itemType = iLong Or .itemType = iSingle Then
                If .itemMax > .itemMin Then
                    If CSng(vNewValue) < .itemMin Then vNewValue = CStr(.itemMin)
                    If CSng(vNewValue) > .itemMax Then vNewValue = CStr(.itemMax)
                End If
            End If
            
            Call setNewValue(tempSheetIndex, tempItemIndex, vNewValue)
        End With
        If AutoRefresh Then Call paintSheet
    End If
End Function
Public Sub setListIndex(tempSheetIndex As Integer, tempItemIndex As Integer, ByVal vNewIndex As Integer)
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        With PropertySheet(tempSheetIndex).sheetItem(tempItemIndex)
            If .itemType = iList Then
                .itemListIndex = vNewIndex
                If .itemListIndex > -1 And .itemListIndex < UBound(.itemList) Then
                    .itemValue = .itemList(.itemListIndex + 1)
                Else
                    .itemListIndex = -1
                    .itemValue = ""
                End If
            End If
        End With
    End If
    If AutoRefresh Then Call paintSheet
End Sub
Private Sub setNewValue(tempSheetIndex As Integer, tempItemIndex As Integer, ByVal vNewValue As String)
    With PropertySheet(tempSheetIndex).sheetItem(tempItemIndex)
        If .itemType = iString Then
            .itemValue = vNewValue
        ElseIf .itemType = iInteger Then
            If IsNumeric(vNewValue) Then
                .itemValue = cornerInt(vNewValue)
                .itemIntegerValue = CInt(.itemValue)
            Else
                .itemValue = "0"
                .itemIntegerValue = 0
            End If
        ElseIf .itemType = iLong Then
            If IsNumeric(vNewValue) Then
                .itemValue = cornerLng(vNewValue)
                .itemLongValue = CLng(.itemValue)
            Else
                .itemValue = "0"
                .itemLongValue = 0
            End If
        ElseIf .itemType = iSingle Then
            If IsNumeric(vNewValue) Then
                .itemValue = cornerSng(vNewValue)
                .itemSingleValue = CSng(.itemValue)
            Else
                .itemValue = "0"
                .itemSingleValue = 0
            End If
        ElseIf .itemType = iBoolean Then
            .itemValue = IIf(LCase(Trim(vNewValue)) = "true", "True", "False")
            .itemBooleanValue = .itemValue = "True"
        End If
    End With
End Sub
'ע��ȡָ�����͵�ֵʱ�����ȡ�����ͺʹ�����ʱע������Ͳ�������ȡ��0ֵ��
Public Function getValue(tempSheetIndex As Integer, tempItemIndex As Integer) As String
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        getValue = PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemValue
    Else
        getValue = "error"
    End If
End Function
Public Function getIntegerValue(tempSheetIndex As Integer, tempItemIndex As Integer) As Integer
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        If PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemType <> iInteger Then Debug.Print "ȡֵ����," & CStr(tempSheetIndex) & "/" & CStr(tempItemIndex) & "����int"
        getIntegerValue = PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemIntegerValue
    End If
End Function
Public Function getLongValue(tempSheetIndex As Integer, tempItemIndex As Integer) As Long
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        If PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemType <> iLong Then Debug.Print "ȡֵ����," & CStr(tempSheetIndex) & "/" & CStr(tempItemIndex) & "����long"
        getLongValue = PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemLongValue
    End If
End Function
Public Function getSingleValue(tempSheetIndex As Integer, tempItemIndex As Integer) As Single
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        If PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemType <> iSingle Then
            Debug.Print "ȡֵ����," & CStr(tempSheetIndex) & "/" & CStr(tempItemIndex) & "����single"
        End If
        getSingleValue = PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemSingleValue
    End If
End Function
Public Function getBooleanValue(tempSheetIndex As Integer, tempItemIndex As Integer) As Boolean
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        If PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemType <> iBoolean Then Debug.Print "ȡֵ����," & CStr(tempSheetIndex) & "/" & CStr(tempItemIndex) & "����boolean"
        getBooleanValue = PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemBooleanValue
    End If
End Function
Public Function getListIndex(tempSheetIndex As Integer, tempItemIndex As Integer) As Integer
    If CheckItem(tempSheetIndex, tempItemIndex) Then
        If PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemType <> iList Then
            Debug.Print "ȡֵ����," & CStr(tempSheetIndex) & "/" & CStr(tempItemIndex) & "����list"
        End If
        getListIndex = PropertySheet(tempSheetIndex).sheetItem(tempItemIndex).itemListIndex
    End If
End Function
Public Property Get չ��() As Boolean
    If currentSheetIndex = -1 Then Exit Property
    չ�� = PropertySheet(currentSheetIndex).sheetExpand
End Property
Public Property Let չ��(ByVal vNewValue As Boolean)
    PropertySheet(currentSheetIndex).sheetExpand = vNewValue
    Call paintSheet
End Property
Public Property Get itemDescription() As String
    If currentSheetIndex = -1 Or currentItemIndex = -1 Then Exit Property
    itemDescription = PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemDescription
End Property
Public Property Let itemDescription(ByVal vNewValue As String)
    PropertySheet(currentSheetIndex).sheetItem(currentItemIndex).itemDescription = vNewValue
    PropertyChanged "pscount"
    Call paintSheet
End Property
'----------------------------------------------------------------------��������
Public Property Get ����и�() As Single
    ����и� = NameHeight
End Property
Public Property Let ����и�(ByVal vNewValue As Single)
    If vNewValue < 12 Then vNewValue = 12
    NameHeight = vNewValue
    PropertyChanged "myNameHeight"
    Call paintSheet
End Property
Public Property Get �����п�() As Single
    �����п� = NameWidth
End Property
Public Property Let �����п�(ByVal vNewValue As Single)
    NameWidth = vNewValue
    PropertyChanged "myNameWidth"
    Call paintSheet
End Property
Public Property Get ��ͷ��ɫ() As OLE_COLOR
    ��ͷ��ɫ = SheetHeadColor
End Property
Public Property Let ��ͷ��ɫ(ByVal vNewValue As OLE_COLOR)
    SheetHeadColor = vNewValue
    Call paintSheet
End Property
Public Property Get �����ɫ() As OLE_COLOR
    �����ɫ = TableColor
End Property
Public Property Let �����ɫ(ByVal vNewValue As OLE_COLOR)
    TableColor = vNewValue
    Call paintSheet
End Property
Public Property Get ����ɫ() As OLE_COLOR
    ����ɫ = TableBackColor
End Property
Public Property Let ����ɫ(ByVal vNewValue As OLE_COLOR)
    TableBackColor = vNewValue
    Call paintSheet
End Property
Public Property Get ��ǰ����ɫ() As OLE_COLOR
    ��ǰ����ɫ = highLightColor
End Property
Public Property Let ��ǰ����ɫ(ByVal vNewValue As OLE_COLOR)
    highLightColor = vNewValue
    Call paintSheet
End Property
Public Property Get �Զ�ˢ��() As Boolean
    �Զ�ˢ�� = AutoRefresh
End Property
Public Property Let �Զ�ˢ��(ByVal vNewValue As Boolean)
    AutoRefresh = vNewValue
    If AutoRefresh Then Call paintSheet
End Property

Public Property Get ��ʾ����() As Boolean
    ��ʾ���� = DescriptionVisible
End Property
Public Property Let ��ʾ����(ByVal vNewValue As Boolean)
    If DescriptionVisible <> vNewValue Then
        DescriptionVisible = vNewValue
        Call UserControl_Resize
    End If
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("NameHeight", NameHeight)
    Call PropBag.WriteProperty("NameWidth", NameWidth)
    Call PropBag.WriteProperty("SheetHeadColor", SheetHeadColor)
    Call PropBag.WriteProperty("TableColor", TableColor)
    Call PropBag.WriteProperty("TableBackColor1", TableBackColor1)
    Call PropBag.WriteProperty("TableBackColor2", TableBackColor2)
    Call PropBag.WriteProperty("txtColor", TxtColor)
    Call PropBag.WriteProperty("txtHotColor", txtHotColor)
    Call PropBag.WriteProperty("highlightColor", highLightColor)
    Call PropBag.WriteProperty("DescriptionVisible", DescriptionVisible)
    Call PropBag.WriteProperty("DescriptionHeight", DescriptionHeight)
    Call PropBag.WriteProperty("AutoRefresh", AutoRefresh)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    NameHeight = PropBag.ReadProperty("NameHeight", 16)
    NameWidth = PropBag.ReadProperty("NameWidth", 90)
    SheetHeadColor = PropBag.ReadProperty("SheetHeadColor", &HDCC887)
    TableColor = PropBag.ReadProperty("TableColor", &HFF8080)
    TableBackColor1 = PropBag.ReadProperty("TableBackColor1", &HFFF7EC)
    TableBackColor2 = PropBag.ReadProperty("TableBackColor2", &HC0FFFF)
    TxtColor = PropBag.ReadProperty("txtColor", 0)
    txtHotColor = PropBag.ReadProperty("txtHotColor", &HFF0000)
    highLightColor = PropBag.ReadProperty("highlightColor", &H80FF&)
    DescriptionVisible = PropBag.ReadProperty("DescriptionVisible", True)
    DescriptionHeight = PropBag.ReadProperty("DescriptionHeight", 50)
    AutoRefresh = PropBag.ReadProperty("AutoRefresh", True)
    If Ambient.UserMode = True Then
        Call Clear
        modified = False
    End If
End Sub

Private Function cornerInt(vNewValue As String) As String
    If Val(vNewValue) > 32767 Then
        cornerInt = "32767"
    ElseIf Val(vNewValue) < -32768 Then
        vcornerInt = "-32768"
    Else
        cornerInt = CStr(CInt(vNewValue))
    End If
End Function
Private Function cornerLng(vNewValue As String) As String
    If Val(vNewValue) > 2147483647# Then
        cornerLng = "2147483647"
    ElseIf Val(vNewValue) < -2147483648# Then
        cornerLng = "-2147483648"
    Else
        cornerLng = CStr(CLng(vNewValue))
    End If
End Function
Private Function cornerSng(vNewValue As String) As String
    If Val(vNewValue) > 2147483647# Then
        cornerSng = "2147483647"
    ElseIf Val(vNewValue) < -2147483648# Then
        cornerSng = "-2147483648"
    Else
        cornerSng = CStr(CSng(vNewValue))
    End If
End Function
Private Function CheckItem(tempSheetIndex As Integer, tempItemIndex As Integer) As Boolean
    If tempSheetIndex > 0 And tempSheetIndex <= PSCount Then
        If tempItemIndex > 0 And tempItemIndex <= PropertySheet(tempSheetIndex).sheetItemCount Then
            CheckItem = True
        Else
            CheckItem = False
        End If
    Else
        CheckItem = False
    End If
End Function

