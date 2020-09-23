VERSION 5.00
Begin VB.UserControl AdvTextBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   600
   ScaleWidth      =   2640
   ToolboxBitmap   =   "AdvTextBox.ctx":0000
   Begin VB.TextBox txtData 
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "AdvTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' This is my first ActiveX control. I hope you enjoy this.
' This is a Simple TextBox with 5 new features added.
'
' 1. HighlightOnFocus : Sets/Gets Wheather the BackColot will be highlighted upon focus
' 2. HighlightColor : The Highlight Color
' 3. ValidCharRange : range of charactor that can be entered in the text box.
' 4. RangeSeparator : separator used for separating ranges in validcharrange
' 5. SelectOnFocus: If the control's contents are selected upon focus
' 6. TabOnenter: wheather thee control looses focusupon pressing ENTER key
'
' Please Send Comments Or Suggestions to Me : nitin_xp@yahoo.com
'---------------------------------------------------------------------------------------
Option Explicit
'Public Enums
Public Enum EnumAppearence
    Flat = 0
    [3D] = 1
End Enum
Public Enum EnumAlignment
    [Left Justify] = 0
    [Right Justify] = 1
    Center = 2
End Enum
Public Enum EnumBorderStyle
    None = 0
    [Fixed Single] = 1
End Enum
Public Enum EnumLinkMode
    None = 0
    Automatic = 1
    Manual = 2
    Notify = 3
End Enum
Public Enum EnumMousePointer
    Default = 0
    Arrow = 1
    Cross = 2
    [I - Beam] = 3
    Icon = 4
    Size = 5
    [Size NE SW] = 6
    [Size N S] = 7
    [Size NW SE] = 8
    [SIze W E] = 9
    [Up Arrow] = 10
    [HourGlass] = 11
    [No Drop] = 12
    [Arrow And HourGlass] = 13
    [Arrow And Question] = 14
    [Size All] = 15
    Custom = 99
End Enum
Public Enum EnumOLEDragMode
    Manual = 0
    Automatic = 1
End Enum
Public Enum EnumOLEDropMode
    None = 0
    Manual = 1
    Automatic = 2
End Enum
Public Enum EnumScrollBars
    None = 0
    Horizontal = 1
    Vertical = 2
    Both = 3
End Enum
'Standard Sizes
Const CONTROL_MIN_HEIGHT = 285
Const CONTROL_MIN_WIDTH = 150
Const CONTROL_STD_HEIGHT = 315
Const CONTROL_STD_WIDTH = 2415
'Default Property Values:
Const m_def_ValidCharRange = "A-Z|a-z|0-9"
Const m_def_RangeSeparator = "|"
Const m_def_HighlightOnFocus = False
Const m_def_HighlightColor = &HFFC0C0
Const m_def_TabOnEnter = True
Const m_def_SelectOnFocus = True
'Const m_def_ValidCharRange = "[A-Z]|[a-z]|[0-9]"
'Property Variables:
Dim m_ValidCharRange As String
Dim m_RangeSeparator As String
Dim m_HighlightOnFocus As Boolean
Dim m_HighlightColor As OLE_COLOR
Dim m_TabOnEnter As Boolean
Dim m_SelectOnFocus As Boolean
'Dim m_ValidCharRange As String
'User Defined Variables:
Dim TempBackColor As OLE_COLOR
Dim SearchString, SearchChar, MyPos As Variant
Dim Token As String
Dim cValid As Boolean
'Event Declarations:
Event Change() 'MappingInfo=txtData,txtData,-1,Change
Event DblClick() 'MappingInfo=txtData,txtData,-1,DblClick
Event Click() 'MappingInfo=txtData,txtData,-1,Click
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtData,txtData,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtData,txtData,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtData,txtData,-1,KeyUp
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtData,txtData,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtData,txtData,-1,MouseUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtData,txtData,-1,MouseDown
Event OLECompleteDrag(Effect As Long) 'MappingInfo=txtData,txtData,-1,OLECompleteDrag
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtData,txtData,-1,OLEDragDrop
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=txtData,txtData,-1,OLEDragOver
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=txtData,txtData,-1,OLESetData
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=txtData,txtData,-1,OLEGiveFeedback
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=txtData,txtData,-1,OLEStartDrag
Event Validate(Cancel As Boolean) 'MappingInfo=txtData,txtData,-1,Validate

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtData.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtData.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtData.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtData.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtData.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtData.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtData.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtData.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,BorderStyle
Public Property Get BorderStyle() As EnumBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = txtData.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As EnumBorderStyle)
    txtData.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Alignment
Public Property Get Alignment() As EnumAlignment
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = txtData.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As EnumAlignment)
    txtData.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Appearance
Public Property Get Appearance() As EnumAppearence
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = txtData.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As EnumAppearence)
    txtData.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,CausesValidation
Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "Returns/sets whether validation occurs on the control which lost focus."
    CausesValidation = txtData.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
    txtData.CausesValidation() = New_CausesValidation
    PropertyChanged "CausesValidation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "Returns/sets the data passed to a destination control in a DDE conversation with another application."
    LinkItem = txtData.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
    txtData.LinkItem() = New_LinkItem
    PropertyChanged "LinkItem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,LinkMode
Public Property Get LinkMode() As EnumLinkMode
Attribute LinkMode.VB_Description = "Returns/sets the type of link used for a DDE conversation and activates the connection."
    LinkMode = txtData.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As EnumLinkMode)
    txtData.LinkMode() = New_LinkMode
    PropertyChanged "LinkMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "Returns/sets the amount of time a control waits for a response to a DDE message."
    LinkTimeout = txtData.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
    txtData.LinkTimeout() = New_LinkTimeout
    PropertyChanged "LinkTimeout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "Returns/sets the source application and topic for a destination control."
    LinkTopic = txtData.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
    txtData.LinkTopic() = New_LinkTopic
    PropertyChanged "LinkTopic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtData.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtData.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtData.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtData.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = txtData.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtData.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,MousePointer
Public Property Get MousePointer() As EnumMousePointer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = txtData.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As EnumMousePointer)
    txtData.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = txtData.MultiLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = txtData.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtData.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = txtData.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    txtData.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,ScrollBars
Public Property Get ScrollBars() As EnumScrollBars
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
    ScrollBars = txtData.ScrollBars
End Property

Private Sub txtData_Change()
    RaiseEvent Change
End Sub

Private Sub txtData_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtData_Click()
    RaiseEvent Click
End Sub

Private Sub txtData_GotFocus()
    If SelectOnFocus = True Then
        txtData.SelStart = 0
        txtData.SelLength = Len(txtData.Text)
    End If
    If HighlightOnFocus = True Then
        TempBackColor = txtData.BackColor
        txtData.BackColor = HighlightColor
    End If
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If KeyCode = vbKeyReturn And TabOnEnter Then
        SendKeys "{Tab}"
        Exit Sub
    End If
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    'Allow BackSpace : 8
    If KeyAscii = 8 Then Exit Sub
    
    Dim theText As String
    theText = Trim(txtData.Text)
    
    'Do Not Allow Initial Space : 32
    If KeyAscii = 32 And Len(theText) = 0 Then
        KeyAscii = 0
    End If
        
    'Traverse Through The Valid Char Set And
    'Do Not Allow Charactors Other Than Specified
    'In Valid Char Range
    
    SearchString = ValidCharRange
    SearchChar = RangeSeparator
    MyPos = 0
    cValid = False
    
    Do
        MyPos = InStr(1, SearchString, SearchChar, 1)
        If MyPos = 0 Then
            Token = SearchString
        Else
            Token = Left(SearchString, MyPos - 1)
        End If
        SearchString = Right(SearchString, Len(SearchString) - MyPos)
        
        If Chr(KeyAscii) Like ("[" & Token & "]") Then
            cValid = True
            Exit Do
        End If
    Loop While MyPos <> 0
    
    If Not cValid Then KeyAscii = 0
End Sub

Private Sub txtData_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtData_LostFocus()
    If HighlightOnFocus Then
        txtData.BackColor = TempBackColor
    End If
End Sub

Private Sub txtData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub txtData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtData_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub txtData_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub txtData_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub txtData_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub txtData_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub txtData_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,OLEDragMode
Public Property Get OLEDragMode() As EnumOLEDragMode
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = txtData.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As EnumOLEDragMode)
    txtData.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,OLEDropMode
Public Property Get OLEDropMode() As EnumOLEDropMode
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
    OLEDropMode = txtData.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As EnumOLEDropMode)
    txtData.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub txtData_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,1,0,[A-Z]|[a-z]|[0-9]
'Public Property Get ValidCharRange() As String
'    ValidCharRange = m_ValidCharRange
'End Property
'
'Public Property Let ValidCharRange(ByVal New_ValidCharRange As String)
'    If Ambient.UserMode Then Err.Raise 382
'    m_ValidCharRange = New_ValidCharRange
'    PropertyChanged "ValidCharRange"
'End Property

Private Sub UserControl_Initialize()
    txtData.Top = 0
    txtData.Left = 0
    UserControl.Width = CONTROL_STD_WIDTH
    UserControl.Height = CONTROL_STD_HEIGHT
    
    txtData.Text = ""
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_ValidCharRange = m_def_ValidCharRange
    m_HighlightOnFocus = m_def_HighlightOnFocus
    m_HighlightColor = m_def_HighlightColor
    m_TabOnEnter = m_def_TabOnEnter
    m_SelectOnFocus = m_def_SelectOnFocus
    m_RangeSeparator = m_def_RangeSeparator
    m_ValidCharRange = m_def_ValidCharRange
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtData.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtData.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtData.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtData.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtData.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtData.Alignment = PropBag.ReadProperty("Alignment", 0)
    txtData.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtData.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
    'Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
    'Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    txtData.LinkItem = PropBag.ReadProperty("LinkItem", "")
    txtData.LinkMode = PropBag.ReadProperty("LinkMode", 0)
    txtData.LinkTimeout = PropBag.ReadProperty("LinkTimeout", 50)
    txtData.LinkTopic = PropBag.ReadProperty("LinkTopic", "")
    txtData.Locked = PropBag.ReadProperty("Locked", False)
    txtData.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtData.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtData.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    txtData.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    txtData.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    txtData.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
'    m_ValidCharRange = PropBag.ReadProperty("ValidCharRange", m_def_ValidCharRange)
    m_HighlightOnFocus = PropBag.ReadProperty("HighlightOnFocus", m_def_HighlightOnFocus)
    m_HighlightColor = PropBag.ReadProperty("HighlightColor", m_def_HighlightColor)
    m_TabOnEnter = PropBag.ReadProperty("TabOnEnter", m_def_TabOnEnter)
    m_SelectOnFocus = PropBag.ReadProperty("SelectOnFocus", m_def_SelectOnFocus)
    txtData.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtData.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtData.SelText = PropBag.ReadProperty("SelText", "")
    txtData.Text = PropBag.ReadProperty("Text", "")
    m_RangeSeparator = PropBag.ReadProperty("RangeSeparator", m_def_RangeSeparator)
    txtData.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    m_ValidCharRange = PropBag.ReadProperty("ValidCharRange", m_def_ValidCharRange)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Height < CONTROL_MIN_HEIGHT Then
        UserControl.Height = CONTROL_MIN_HEIGHT
    End If
    txtData.Height = UserControl.ScaleHeight
    
    If UserControl.Width < CONTROL_MIN_WIDTH Then
        UserControl.Width = CONTROL_MIN_WIDTH
    End If
    txtData.Width = UserControl.ScaleWidth
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", txtData.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtData.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtData.Enabled, True)
    Call PropBag.WriteProperty("Font", txtData.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", txtData.BorderStyle, 1)
    Call PropBag.WriteProperty("Alignment", txtData.Alignment, 0)
    Call PropBag.WriteProperty("Appearance", txtData.Appearance, 1)
    Call PropBag.WriteProperty("CausesValidation", txtData.CausesValidation, True)
    'Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)
    'Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("LinkItem", txtData.LinkItem, "")
    Call PropBag.WriteProperty("LinkMode", txtData.LinkMode, 0)
    Call PropBag.WriteProperty("LinkTimeout", txtData.LinkTimeout, 50)
    Call PropBag.WriteProperty("LinkTopic", txtData.LinkTopic, "")
    Call PropBag.WriteProperty("Locked", txtData.Locked, False)
    Call PropBag.WriteProperty("MaxLength", txtData.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtData.MousePointer, 0)
    Call PropBag.WriteProperty("PasswordChar", txtData.PasswordChar, "")
    Call PropBag.WriteProperty("RightToLeft", txtData.RightToLeft, False)
    Call PropBag.WriteProperty("OLEDragMode", txtData.OLEDragMode, 0)
    Call PropBag.WriteProperty("OLEDropMode", txtData.OLEDropMode, 0)
'    Call PropBag.WriteProperty("ValidCharRange", m_ValidCharRange, m_def_ValidCharRange)
    Call PropBag.WriteProperty("HighlightOnFocus", m_HighlightOnFocus, m_def_HighlightOnFocus)
    Call PropBag.WriteProperty("HighlightColor", m_HighlightColor, m_def_HighlightColor)
    Call PropBag.WriteProperty("TabOnEnter", m_TabOnEnter, m_def_TabOnEnter)
    Call PropBag.WriteProperty("SelectOnFocus", m_SelectOnFocus, m_def_SelectOnFocus)
    Call PropBag.WriteProperty("SelLength", txtData.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtData.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtData.SelText, "")
    Call PropBag.WriteProperty("Text", txtData.Text, "")
    Call PropBag.WriteProperty("RangeSeparator", m_RangeSeparator, m_def_RangeSeparator)
    Call PropBag.WriteProperty("WhatsThisHelpID", txtData.WhatsThisHelpID, 0)
    Call PropBag.WriteProperty("ValidCharRange", m_ValidCharRange, m_def_ValidCharRange)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get HighlightOnFocus() As Boolean
Attribute HighlightOnFocus.VB_Description = "Returns/Sets wheather the contro's back color will be highlited by Highlight Color."
    HighlightOnFocus = m_HighlightOnFocus
End Property

Public Property Let HighlightOnFocus(ByVal New_HighlightOnFocus As Boolean)
    m_HighlightOnFocus = New_HighlightOnFocus
    PropertyChanged "HighlightOnFocus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&HFFC0C0
Public Property Get HighlightColor() As OLE_COLOR
Attribute HighlightColor.VB_Description = "Returns/Sets the back color of the control when the focus is received."
    HighlightColor = m_HighlightColor
End Property

Public Property Let HighlightColor(ByVal New_HighlightColor As OLE_COLOR)
    m_HighlightColor = New_HighlightColor
    PropertyChanged "HighlightColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get TabOnEnter() As Boolean
Attribute TabOnEnter.VB_Description = "Returns/Sets wheather the control looses focus upon return key hit."
    TabOnEnter = m_TabOnEnter
End Property

Public Property Let TabOnEnter(ByVal New_TabOnEnter As Boolean)
    m_TabOnEnter = New_TabOnEnter
    PropertyChanged "TabOnEnter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get SelectOnFocus() As Boolean
Attribute SelectOnFocus.VB_Description = "Returns/Sets if the control's contents are selected automatically upon focus receipt."
    SelectOnFocus = m_SelectOnFocus
End Property

Public Property Let SelectOnFocus(ByVal New_SelectOnFocus As Boolean)
    m_SelectOnFocus = New_SelectOnFocus
    PropertyChanged "SelectOnFocus"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtData.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtData.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtData.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtData.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtData.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtData.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtData.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtData.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,|
Public Property Get RangeSeparator() As String
Attribute RangeSeparator.VB_Description = "Returns/Sets the separator used in ValidCharRange."
    RangeSeparator = m_RangeSeparator
End Property

Public Property Let RangeSeparator(ByVal New_RangeSeparator As String)
    If Ambient.UserMode Then Err.Raise 382
    m_RangeSeparator = New_RangeSeparator
    PropertyChanged "RangeSeparator"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
    WhatsThisHelpID = txtData.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    txtData.WhatsThisHelpID() = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,A-Z|a-z|0-9
Public Property Get ValidCharRange() As String
Attribute ValidCharRange.VB_Description = "Returns/Sets the valid range of charactors that the control will receive."
    ValidCharRange = m_ValidCharRange
End Property

Public Property Let ValidCharRange(ByVal New_ValidCharRange As String)
    If Ambient.UserMode Then Err.Raise 382
    m_ValidCharRange = New_ValidCharRange
    PropertyChanged "ValidCharRange"
End Property

