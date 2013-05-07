VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl LVListBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   PropertyPages   =   "LVListBox.ctx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   4740
   ToolboxBitmap   =   "LVListBox.ctx":0035
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "LVListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETCOUNTPERPAGE As Long = (LVM_FIRST + 40)
Private Const LVLB_FILE_SEPARATOR = "Ø"

Public Enum eLVLBBorderStyle
 eLVLBBNone = 0
 eLVLBBFixedSingle = 1
End Enum

Private ItemX As ListItem
Private ItemVisibles As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Event AfterEdit(Cancel As Integer, NewString As String)
Public Event BeforeEdit(Cancel As Integer)
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OnAdd(NewIndex As Long)
Public Event OnRemove(OldIndex As Long)
Public Event OnSelect(NewValue As Boolean)
Public Event Resize()

Public Property Get Appearance() As AppearanceConstants
  Appearance = UserControl.Appearance
End Property
Public Property Let Appearance(ByVal NewValue As AppearanceConstants)
  UserControl.Appearance = NewValue
  PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
 BackColor = ListView1.BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
 ListView1.BackColor = NewValue
 PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As eLVLBBorderStyle
 BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal NewValue As eLVLBBorderStyle)
 UserControl.BorderStyle = NewValue
 PropertyChanged "BorderStyle"
End Property

Public Property Get CheckBoxes() As Boolean
 CheckBoxes = ListView1.CheckBoxes
End Property
Public Property Let CheckBoxes(ByVal NewValue As Boolean)
 ListView1.CheckBoxes = NewValue
 PropertyChanged "CheckBoxes"
End Property

Public Property Get Enabled() As Boolean
 Enabled = ListView1.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
 ListView1.Enabled = NewValue
 UserControl.Enabled = NewValue
 PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = ListView1.ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
  ListView1.ForeColor = NewValue
  PropertyChanged "ForeColor"
End Property

Public Property Get Font() As StdFont
  Set Font = ListView1.Font
End Property
Public Property Set Font(ByVal NewValue As StdFont)
  Set ListView1.Font = NewValue
  PropertyChanged "Font"
End Property

Public Property Get HideSelection() As Boolean
  HideSelection = ListView1.HideSelection
End Property
Public Property Let HideSelection(ByVal NewValue As Boolean)
  ListView1.HideSelection = NewValue
  PropertyChanged "HideSelection"
End Property

Public Property Get HotTracking() As Boolean
  HotTracking = ListView1.HotTracking
End Property
Public Property Let HotTracking(ByVal NewValue As Boolean)
  ListView1.HotTracking = NewValue
  PropertyChanged "HotTracking"
End Property

Public Property Get HoverSelection() As Boolean
  HoverSelection = ListView1.HoverSelection
End Property
Public Property Let HoverSelection(ByVal NewValue As Boolean)
  ListView1.HoverSelection = NewValue
  PropertyChanged "HoverSelection"
End Property

Public Property Get ImageList() As ImageList
 Set ImageList = ListView1.SmallIcons
End Property
Public Property Set ImageList(ByVal NewValue As ImageList)
 Set ListView1.SmallIcons = NewValue
 PropertyChanged "ImageList"
End Property

Public Property Get LabelEdit() As ListLabelEditConstants
  LabelEdit = ListView1.LabelEdit
End Property
Public Property Let LabelEdit(ByVal NewValue As ListLabelEditConstants)
  ListView1.LabelEdit = NewValue
  PropertyChanged "LabelEdit"
End Property

Public Property Get LabelWrap() As Boolean
  LabelWrap = ListView1.LabelWrap
End Property
Public Property Let LabelWrap(ByVal NewValue As Boolean)
  ListView1.LabelWrap = NewValue
  PropertyChanged "LabelWrap"
End Property

Public Property Get List(Index As Long) As String
Attribute List.VB_UserMemId = 0
  If (ListView1.ListItems.Count > 0) And (Index > 0 And Index <= ListView1.ListItems.Count) Then
   List = ListView1.ListItems(Index).Text
  End If
End Property
Public Property Let List(Index As Long, NewValue As String)
 If (ListView1.ListItems.Count > 0) And (Index > 0 And Index <= ListView1.ListItems.Count) Then
   ListView1.ListItems(Index).Text = NewValue
   RaiseEvent Change
 End If
End Property

Public Property Get ListCount() As Long
 ListCount = ListView1.ListItems.Count
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
 If ListView1.ListItems.Count = 0 Then
  ListIndex = -1
 Else
  If Not (ListView1.SelectedItem Is Nothing) Then
   ListIndex = ListView1.SelectedItem.Index
  Else
   ListIndex = -1
  End If
 End If
End Property
Public Property Let ListIndex(NewValue As Long)
 If (ListView1.ListItems.Count > 0) And (NewValue > 0 And NewValue <= ListView1.ListItems.Count) Then
  Set ListView1.SelectedItem = ListView1.ListItems(NewValue)
  ListView1.SelectedItem.EnsureVisible
  RaiseEvent Change
 End If
End Property

Public Property Get MouseIcon() As StdPicture
  Set MouseIcon = ListView1.MouseIcon
End Property
Public Property Set MouseIcon(ByVal NewValue As StdPicture)
  Set ListView1.MouseIcon = NewValue
  PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MSComctlLib.MousePointerConstants
  MousePointer = ListView1.MousePointer
End Property
Public Property Let MousePointer(ByVal NewValue As MSComctlLib.MousePointerConstants)
  ListView1.MousePointer = NewValue
  PropertyChanged "MousePointer"
End Property

Public Property Get MultiSelect() As Boolean
  MultiSelect = ListView1.MultiSelect
End Property
Public Property Let MultiSelect(ByVal NewValue As Boolean)
  ListView1.MultiSelect = NewValue
  PropertyChanged "MultiSelect"
End Property

Public Property Get Picture() As StdPicture
  Set Picture = ListView1.Picture
End Property
Public Property Set Picture(ByVal NewValue As StdPicture)
  Set ListView1.Picture = NewValue
  PropertyChanged "Picture"
End Property

Public Property Get PictureAlignment() As ListPictureAlignmentConstants
  PictureAlignment = ListView1.PictureAlignment
End Property
Public Property Let PictureAlignment(ByVal NewValue As ListPictureAlignmentConstants)
  ListView1.PictureAlignment = NewValue
  PropertyChanged "PictureAlignment"
End Property

Public Property Get Selected(Index As Long) As Boolean
 If (ListView1.ListItems.Count > 0) And (Index > 0 And Index <= ListView1.ListItems.Count) Then
  If ListView1.CheckBoxes = False Then
   Selected = ListView1.ListItems(Index).Selected
  Else
   Selected = ListView1.ListItems(Index).Checked
  End If
 End If
End Property
Public Property Let Selected(Index As Long, NewValue As Boolean)
  If (ListView1.ListItems.Count > 0) And (Index > 0 And Index <= ListView1.ListItems.Count) Then
   If ListView1.CheckBoxes = False Then
    ListView1.ListItems(Index).Selected = NewValue
   Else
    ListView1.ListItems(Index).Checked = NewValue
   End If
   RaiseEvent OnSelect(NewValue)
 End If
End Property

Public Property Get Sorted() As Boolean
  Sorted = ListView1.Sorted
End Property
Public Property Let Sorted(ByVal NewValue As Boolean)
  ListView1.Sorted = NewValue
  PropertyChanged "Sorted"
End Property

Public Property Get SortOrder() As ListSortOrderConstants
  SortOrder = ListView1.SortOrder
End Property
Public Property Let SortOrder(ByVal NewValue As ListSortOrderConstants)
  ListView1.SortOrder = NewValue
  PropertyChanged "SortOrder"
End Property

Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "400"
 If Not (ListView1.SelectedItem Is Nothing) Then
  Text = ListView1.SelectedItem.Text
 End If
End Property
Public Property Let Text(ByVal NewValue As String)
 If Not (ListView1.SelectedItem Is Nothing) Then
  ListView1.SelectedItem.Text = NewValue
 End If
End Property





Public Function AddItem(Optional sText As String, Optional iSmallIcon As Integer = 0, Optional lForeColor As Long = -1, Optional bBold As Boolean = False, Optional bChecked As Boolean = False, Optional bGhosted As Boolean = False, Optional bSelected As Boolean = False, Optional sKey As String = "", Optional sTag As String = "") As Long
On Error GoTo err1

 Set ItemX = ListView1.ListItems.Add(, , sText)
  If iSmallIcon > 0 Then ItemX.SmallIcon = iSmallIcon
  If lForeColor > -1 Then ItemX.ForeColor = lForeColor
  ItemX.Bold = bBold
  ItemX.Checked = bChecked
  ItemX.Ghosted = bGhosted
  ItemX.Selected = bSelected
  If Trim(sTag) <> "" Then ItemX.Tag = sTag
  If Trim(sKey) <> "" Then ItemX.key = sKey
  AddItem = ItemX.Index
  RaiseEvent OnAdd(ItemX.Index)
  Exit Function
  
err1:
 AddItem = -1
 Exit Function
End Function

Public Sub Clear()
 LockWindowUpdate ListView1.hwnd
 ListView1.ListItems.Clear
 LockWindowUpdate 0
End Sub

Public Sub Refresh()
 ListView1.Refresh
 Call UserControl_Resize
End Sub

Public Sub RemoveItem(Index As Long)
On Error Resume Next
  ListView1.ListItems.Remove Index
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
 MsgBox "LVListBox - Listbox Control based into Listview" & vbCrLf & "Developed by Mauricio Cunha" & vbCrLf & "http://www.mcunha98.cjb.net" & vbCrLf & "mcunha98@terra.com.br", , "About..."
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
 RaiseEvent AfterEdit(Cancel, NewString)
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
 RaiseEvent BeforeEdit(Cancel)
End Sub

Private Sub ListView1_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
 RaiseEvent Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
 RaiseEvent Click
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
 RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Function FindItem(TextToFind As String, Optional WholeWordOnly As Boolean = False, Optional CaseSensitive As Boolean = False, Optional SelectOnFind As Boolean = False) As Integer
    Dim lngIndex As Long
    Dim lngIndexSub As Long
    Dim strCurrItem As String
    
    If CaseSensitive = True Then TextToFind = UCase(TextToFind)
    
    FindItem = 0
    If ListView1.ListItems.Count < 1 Then Exit Function
    If ListView1.SelectedItem.Index = -1 Then ListView1.SelectedItem.Index = 1
    For lngIndex = ListView1.SelectedItem.Index - -1 To ListView1.ListItems.Count
        If CaseSensitive = True Then
            strCurrItem = UCase(ListView1.ListItems.Item(lngIndex).Text)
        Else
            strCurrItem = ListView1.ListItems.Item(lngIndex).Text
        End If
        
        If WholeWordOnly = True Then
            If strCurrItem = TextToFind Then GoTo Finalize
        Else
            If InStr(strCurrItem, TextToFind) > 0 Then GoTo Finalize
        End If
    Next lngIndex
    Exit Function
    
Finalize:
    FindItem = lngIndex
    If SelectOnFind = True Then
     ListView1.ListItems.Item(lngIndex).EnsureVisible
     ListView1.ListItems.Item(lngIndex).Selected = SelectOnFind
    End If
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Appearance = PropBag.ReadProperty("Appearance", cc3D)
 BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
 BorderStyle = PropBag.ReadProperty("BorderStyle", eLVLBBFixedSingle)
 CheckBoxes = PropBag.ReadProperty("CheckBoxes", False)
 Enabled = PropBag.ReadProperty("Enabled", True)
 Set Font = PropBag.ReadProperty("Font", Ambient.Font)
 ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
 HideSelection = PropBag.ReadProperty("HideSelection", True)
 HotTracking = PropBag.ReadProperty("HotTracking", False)
 HoverSelection = PropBag.ReadProperty("HoverSelection", False)
 LabelEdit = PropBag.ReadProperty("LabelEdit", lvwAutomatic)
 LabelWrap = PropBag.ReadProperty("LabelWrap", False)
 Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 MousePointer = PropBag.ReadProperty("MousePointer", 0)
 MultiSelect = PropBag.ReadProperty("MultiSelect", False)
 Set Picture = PropBag.ReadProperty("Picture", Nothing)
 PictureAlignment = PropBag.ReadProperty("PictureAlignment", lvwTile)
 Sorted = PropBag.ReadProperty("Sorted", False)
 SortOrder = PropBag.ReadProperty("SortOrder", lvwAscending)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Appearance", UserControl.Appearance, cc3D)
 Call PropBag.WriteProperty("BackColor", ListView1.BackColor, vbWindowBackground)
 Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, eLVLBBFixedSingle)
 Call PropBag.WriteProperty("CheckBoxes", ListView1.CheckBoxes, False)
 Call PropBag.WriteProperty("Enabled", ListView1.Enabled, True)
 Call PropBag.WriteProperty("Font", ListView1.Font, Ambient.Font)
 Call PropBag.WriteProperty("ForeColor", ListView1.ForeColor, vbButtonText)
 Call PropBag.WriteProperty("HideSelection", ListView1.HideSelection, True)
 Call PropBag.WriteProperty("HotTracking", ListView1.HotTracking, False)
 Call PropBag.WriteProperty("HoverSelection", ListView1.HoverSelection, False)
 Call PropBag.WriteProperty("LabelEdit", ListView1.LabelEdit, lvwAutomatic)
 Call PropBag.WriteProperty("LabelWrap", ListView1.LabelWrap, False)
 Call PropBag.WriteProperty("MouseIcon", ListView1.MouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", ListView1.MousePointer, 0)
 Call PropBag.WriteProperty("MultiSelect", ListView1.MultiSelect, False)
 Call PropBag.WriteProperty("Picture", ListView1.Picture, Nothing)
 Call PropBag.WriteProperty("PictureAlignment", ListView1.PictureAlignment, lvwTile)
 Call PropBag.WriteProperty("Sorted", ListView1.Sorted, False)
 Call PropBag.WriteProperty("SortOrder", ListView1.SortOrder, lvwAscending)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
 ListView1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
 ItemVisibles = SendMessage(ListView1.hwnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)
 ListView1.ColumnHeaders(1).Width = ListView1.Width - IIf(ItemVisibles >= ListView1.ListItems.Count, 0, 200)
 RaiseEvent Resize
End Sub

Public Sub StartEdit()
  ListView1.StartLabelEdit
End Sub

Private Function FileExists(sPath As String) As Boolean
Dim lngRetVal As Long
On Error Resume Next
  lngRetVal = Len(Dir(sPath))
  If Err Or lngRetVal = 0 Then FileExists = False Else FileExists = True
End Function

Public Function SaveToFile(sFilename As String, Optional bSaveWithFormats As Boolean = False) As Boolean
Dim lFile As Long
On Error GoTo err1

lFile = FreeFile
Open sFilename For Output As lFile
  For Each ItemX In ListView1.ListItems
   If bSaveWithFormats = False Then
    Print #lFile, Replace(ItemX.Text, LVLB_FILE_SEPARATOR, "")
   Else
    Print #lFile, Replace(ItemX.Text, LVLB_FILE_SEPARATOR, "") & LVLB_FILE_SEPARATOR & ItemX.SmallIcon & LVLB_FILE_SEPARATOR & ItemX.ForeColor & LVLB_FILE_SEPARATOR & IIf(ItemX.Bold = False, 0, -1) & LVLB_FILE_SEPARATOR & IIf(ItemX.Checked = False, 0, -1) & LVLB_FILE_SEPARATOR & IIf(ItemX.Ghosted = False, 0, -1) & LVLB_FILE_SEPARATOR & IIf(ItemX.Selected = False, 0, -1) & LVLB_FILE_SEPARATOR & ItemX.Tag & LVLB_FILE_SEPARATOR & ItemX.key
   End If
  Next
Close #lFile
SaveToFile = True
Exit Function
 
 
err1:
 SaveToFile = False
 Exit Function
End Function

Public Function LoadFromFile(sFilename As String, Optional bLoadWithFormats As Boolean = False) As Boolean
Dim lFile As Long
Dim sLine As String
Dim sImage As String
On Error GoTo err1

ListView1.ListItems.Clear
lFile = FreeFile
Open sFilename For Input As lFile
 Do While Not EOF(lFile)
  Line Input #lFile, sLine
   If bLoadWithFormats = False Then
    Set ItemX = ListView1.ListItems.Add(, , sLine)
   Else
    If InStr(1, sLine, LVLB_FILE_SEPARATOR) > 0 Then
     sImage = Split(sLine, LVLB_FILE_SEPARATOR)(1)
     Set ItemX = ListView1.ListItems.Add(, , Split(sLine, LVLB_FILE_SEPARATOR)(0))
         ItemX.SmallIcon = IIf(Val(sImage) = 0, Empty, Val(sImage))
         ItemX.ForeColor = Split(sLine, LVLB_FILE_SEPARATOR)(2)
         ItemX.Bold = CBool(Split(sLine, LVLB_FILE_SEPARATOR)(3))
         ItemX.Checked = CBool(Split(sLine, LVLB_FILE_SEPARATOR)(4))
         ItemX.Ghosted = CBool(Split(sLine, LVLB_FILE_SEPARATOR)(5))
         ItemX.Selected = CBool(Split(sLine, LVLB_FILE_SEPARATOR)(6))
         ItemX.Tag = Split(sLine, LVLB_FILE_SEPARATOR)(7)
         ItemX.key = Split(sLine, LVLB_FILE_SEPARATOR)(8)
    End If
   End If
  Loop
Close #lFile
LoadFromFile = True
Call Refresh
Exit Function

err1:
 LoadFromFile = False
 Exit Function
End Function
