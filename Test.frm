VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucTreeView 1.3 - Test"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   630
   StartUpPosition =   2  'CenterScreen
   Begin Test.ucTreeView ucTreeView1 
      Height          =   3960
      Left            =   150
      TabIndex        =   1
      Top             =   540
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   6985
   End
   Begin VB.PictureBox fraOLEDragInsertStyle 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   5325
      ScaleHeight     =   555
      ScaleWidth      =   1875
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4620
      Width           =   1875
      Begin VB.OptionButton optOLEDragInsertStyle 
         Appearance      =   0  'Flat
         Caption         =   "DropHilite"
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   255
         Value           =   -1  'True
         Width           =   1470
      End
      Begin VB.OptionButton optOLEDragInsertStyle 
         Appearance      =   0  'Flat
         Caption         =   "InsertMark (default)"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1830
      End
   End
   Begin VB.PictureBox fraDropEffect 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   1110
      ScaleHeight     =   555
      ScaleWidth      =   1530
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4620
      Width           =   1530
      Begin VB.OptionButton optDropEffect 
         Appearance      =   0  'Flat
         Caption         =   "Copy"
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   255
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optDropEffect 
         Appearance      =   0  'Flat
         Caption         =   "Move"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7365
      OLEDragMode     =   1  'Automatic
      TabIndex        =   5
      Top             =   2565
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   7365
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      TabIndex        =   4
      Top             =   1575
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      Height          =   885
      HideSelection   =   0   'False
      Left            =   7365
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      TabIndex        =   3
      Top             =   540
      Width           =   1920
   End
   Begin Test.ucTreeView ucTreeView2 
      Height          =   3960
      Left            =   3750
      TabIndex        =   2
      Top             =   540
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   6985
   End
   Begin VB.Label lblNodeFullPath 
      Caption         =   "NodeFullPath:"
      Height          =   240
      Left            =   300
      TabIndex        =   19
      Top             =   6255
      Width           =   990
   End
   Begin VB.Label lblNodeFullPathVal 
      Height          =   585
      Left            =   1365
      TabIndex        =   20
      Top             =   6255
      Width           =   7920
   End
   Begin VB.Label lblInsertAfterVal 
      Height          =   240
      Left            =   1365
      TabIndex        =   18
      Top             =   5955
      Width           =   1500
   End
   Begin VB.Label lblInsertAfter 
      Caption         =   "InsertAfter:"
      Height          =   240
      Left            =   300
      TabIndex        =   17
      Top             =   5955
      Width           =   990
   End
   Begin VB.Label lblNodeDrop 
      Caption         =   "hNodeDrop:"
      Height          =   240
      Left            =   300
      TabIndex        =   15
      Top             =   5670
      Width           =   990
   End
   Begin VB.Label lblDragDropTitle 
      BackColor       =   &H80000010&
      Caption         =   " OLE drag & drop test: Move/copy to second TreeView"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   9135
   End
   Begin VB.Label lblNodeDropVal 
      Height          =   240
      Left            =   1365
      TabIndex        =   16
      Top             =   5670
      Width           =   1500
   End
   Begin VB.Label lblOLEDropInfo 
      BackColor       =   &H80000010&
      Caption         =   " OLE drop info:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   150
      TabIndex        =   14
      Top             =   5310
      UseMnemonic     =   0   'False
      Width           =   9135
   End
   Begin VB.Label lblOLEDragInsertStyle 
      Caption         =   "OLEDragInsertStyle:"
      Height          =   270
      Left            =   3765
      TabIndex        =   10
      Top             =   4620
      Width           =   1515
   End
   Begin VB.Label lblDropEffect 
      Caption         =   "Drop effect:"
      Height          =   270
      Left            =   165
      TabIndex        =   6
      Top             =   4620
      Width           =   1005
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTestTop 
      Caption         =   "&Test"
      Begin VB.Menu mnuTest 
         Caption         =   "&Reset all"
         Index           =   0
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuPSCTop 
      Caption         =   "&PSC"
      Begin VB.Menu mnuPSC 
         Caption         =   "Go to ucTreeView 1.3 PSC page..."
         Index           =   0
      End
      Begin VB.Menu mnuPSC 
         Caption         =   "Go to ucTreeView 1.2 PSC page..."
         Index           =   1
      End
   End
   Begin VB.Menu mnuContextTop 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuContext 
         Caption         =   "Delete"
         Index           =   0
      End
      Begin VB.Menu mnuContext 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Expand children"
         Index           =   2
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Collapse children"
         Index           =   3
      End
      Begin VB.Menu mnuContext 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Cancel"
         Index           =   5
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Some API
Private Const GWL_STYLE As Long = (-16)
Private Const BS_FLAT   As Long = &H8000&
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const SW_SHOW   As Long = 5
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const URL_UCTREEVIEW_1_3 As String = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57047&lngWId=1"
Private Const URL_UCTREEVIEW_1_2 As String = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56958&lngWId=1"





'========================================================================================
' Initialize all...
'========================================================================================

Private m_lKey As Long ' Auto-increased key for ucTreeView2

Private Sub Form_Load()

  Dim oCtl As Control

    Call pvInitializeTreeView1
    Call pvInitializeTreeView2
    Call pvFillTreeView1
    Call pvFillTreeView2
    
    Call pvFillStandartControls
    
    For Each oCtl In fTest.Controls
        If (TypeOf oCtl Is CommandButton) Then
            Call pvFlattenButton(oCtl.hWnd)
        End If
    Next oCtl
End Sub

Private Sub pvInitializeTreeView1()
    
  Dim lResIcon As Long
  
    With ucTreeView1
        
        Call .Initialize
        Call .InitializeImageList
        
        For lResIcon = 101 To 105
            Call .AddIcon(VB.LoadResPicture(lResIcon, vbResIcon))
        Next lResIcon
        
        .ItemHeight = 18
        .HasButtons = True
        .HasLines = True
        .HasRootLines = True
    
        .OLEDragMode = [drgAutomatic]
        .OLEDropMode = [drpNone] 'Default
    End With
End Sub

Private Sub pvInitializeTreeView2()
    
  Dim lResIcon As Long
    
    With ucTreeView2
        
        Call .Initialize
        Call .InitializeImageList
        
        For lResIcon = 101 To 108
            Call .AddIcon(LoadResPicture(lResIcon, vbResIcon))
        Next lResIcon
        
        .ItemHeight = 18
        .HasButtons = True
        .HasLines = True
        .HasRootLines = True
        .CheckBoxes = True
        .LabelEdit = True
        
        .BackColor = &HC8FFC8
        .ForeColor = &H0
        .LineColor = &H8000&
        
        .OLEDragMode = [drgNone] 'Default
        .OLEDropMode = [drpManual]
        
        .OLEDragInsertStyle = [disDropHilite]
        .OLEDragAutoExpand = True
    End With
End Sub

Private Sub pvFillTreeView1()
  
  Dim lKey        As Long
  
  Dim lBook       As Long
  Dim lChapter    As Long
  Dim lPage       As Long
  Dim lNote       As Long
  
  Dim hBook       As Long
  Dim hChapter    As Long
  Dim hPage       As Long
    
    With ucTreeView1
        
        Call .SetRedrawMode(Enable:=False)
        Call .Clear
        
        For lBook = 1 To 1
            lKey = lKey + 1
            hBook = .AddNode(, , lKey, "Book #" & lBook, 0, 1)
            
            For lChapter = 1 To 5
                lKey = lKey + 1
                hChapter = .AddNode(hBook, , lKey, "Chapter #" & lChapter, 0, 1)
                
                For lPage = 1 To 10
                    lKey = lKey + 1
                    hPage = .AddNode(hChapter, , lKey, "Page #" & lPage, 2, 2)
                    
                    For lNote = 1 To 2
                        lKey = lKey + 1
                        Call .AddNode(hPage, , lKey, "Note #" & lNote, 3, 3)
        
        Next lNote, lPage, lChapter, lBook
        
        .SelectedNode = .NodeRoot
        Call .Expand(.SelectedNode)
        Call .SetRedrawMode(Enable:=True)
    End With
End Sub

Private Sub pvFillTreeView2()
  
  Dim lFolder    As Long
  Dim lSubFolder As Long
  
  Dim hFolder    As Long
    
    With ucTreeView2
        
        Call .SetRedrawMode(Enable:=False)
        Call .Clear
        
        For lFolder = 1 To 2
            m_lKey = m_lKey + 1
            hFolder = .AddNode(, , m_lKey, "Folder #" & lFolder, 5, 6)
            
            For lSubFolder = 1 To 10
                m_lKey = m_lKey + 1
                
                Call .AddNode(hFolder, , m_lKey, "Folder #" & lFolder & "." & lSubFolder, 5, 6)
        
        Next lSubFolder, lFolder
        
        .SelectedNode = .NodeRoot
        Call .Expand(0, ExpandChildren:=True)
        Call .EnsureVisible(.SelectedNode)
        Call .SetRedrawMode(Enable:=True)
    End With
End Sub

Private Sub pvFillStandartControls()

    Text1.Text = "TextBox text"
    
    Call List1.Clear
    Call List1.AddItem("ListBox item 0")
    Call List1.AddItem("ListBox item 1")
    Call List1.AddItem("ListBox item 2")

    Call Combo1.Clear
    Call Combo1.AddItem("ComboBox item 0")
    Call Combo1.AddItem("ComboBox item 1")
    Call Combo1.AddItem("ComboBox item 2")
    Combo1.Text = Combo1.List(0)
End Sub

'========================================================================================
' Menus
'========================================================================================

Private Sub mnuFile_Click(Index As Integer)
    
    Call Unload(Me)
End Sub

Private Sub mnuTest_Click(Index As Integer)

    Call pvFillTreeView1
    Call pvFillTreeView2
    Call pvFillStandartControls
End Sub

Private Sub mnuPSC_Click(Index As Integer)

    Select Case Index
        Case 0
            Call pvNavigate(URL_UCTREEVIEW_1_3)
        Case 1
            Call pvNavigate(URL_UCTREEVIEW_1_2)
    End Select
End Sub

Private Sub mnuContext_Click(Index As Integer)
    
    With ucTreeView2
        Select Case Index
            Case 0
                Call .DeleteNode(.SelectedNode)
            Case 2
                Call .Expand(.SelectedNode, True)
            Case 3
                Call .Collapse(.SelectedNode, True)
        End Select
    End With
End Sub

Private Sub ucTreeView2_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    
  Dim hNode As Long
    
    With ucTreeView2
        If (Button = vbRightButton) Then
            hNode = .HitTest(x, y, False)
            If (hNode) Then
                .SelectedNode = hNode
                Call Me.PopupMenu(mnuContextTop, , .Left + x, .Top + y, mnuContext(0))
            End If
        End If
    End With
End Sub

Private Sub ucTreeView2_NodeCheck(ByVal hNode As Long)
    
    With ucTreeView2
        Call .CheckChildren(hNode, .NodeChecked(hNode))
    End With
End Sub



'========================================================================================
' OLE Drag & Drop test
'========================================================================================

Private Sub optOLEDragInsertStyle_Click(Index As Integer)
    ucTreeView2.OLEDragInsertStyle = Index
End Sub

'//

Private Sub ucTreeView1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    
    '-- ucTreeView1 has started dragging. Set allowed effects:
    
    AllowedEffects = IIf(optDropEffect(0), vbDropEffectMove, vbDropEffectCopy)
End Sub

Private Sub ucTreeView2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
        
  Dim hNodeDrop   As Long
  Dim InsertAfter As Boolean
  
    '-- 'Something is over' ucTreeView2:
    
    If (Effect <> vbOLEDropNone) Then
        
        '-- Get Drop info:
        Call ucTreeView2.OLEGetDropInfo(hNodeDrop, InsertAfter)
        
        lblNodeDropVal.Caption = hNodeDrop
        lblInsertAfterVal.Caption = IIf(InsertAfter, "True", "False")
        lblNodeFullPathVal.Caption = ucTreeView2.NodeFullPath(hNodeDrop)
        
      Else
        lblNodeDropVal.Caption = vbNullString
        lblInsertAfterVal.Caption = vbNullString
        lblNodeFullPathVal.Caption = vbNullString
    End If
End Sub

Private Sub ucTreeView2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim hNodeDrag     As Long
  Dim hNodeDrop     As Long
  Dim hNodeInsert   As Long
  Dim hWndTreeView  As Long
  Dim bInstertAfter As Boolean
  Dim eRelation     As tvRelationConstants
  Dim sText()       As String
  Dim lIdx          As Long
    
    '-- Something has been Dropped on ucTreevView2....
    
    With ucTreeView2
        
        '-- Is there our format (ucTreeView)
        If (.OLEIsMyFormat(Data)) Then
            
            '-- Yes, extract Drag info (TreeView handle and node-drag)
            Call .OLEGetDragInfo(Data, hWndTreeView, hNodeDrag)
            
            '-- Extract Drop info (node-over and insert-after flag)
            '   Insert-after flag will be always True when OLEDragInsertStyle=[disDropHilite]
            '   It's your choice decide how to insert data...
            Call .OLEGetDropInfo(hNodeDrop, bInstertAfter)
            
            '-- Check insertion relation
            If (.OLEDragInsertStyle = [disInsertMark]) Then
                If (bInstertAfter = False) Then
                    eRelation = [rPrevious]
                  Else
                    eRelation = [rNext]
                End If
              Else
                eRelation = [rLast]
            End If
            '-- Let's insert our node/s
            Call pvCopyNode(hNodeDrag, hNodeDrop, eRelation)
            
            '-- Are we moving nodes ?
            If (optDropEffect(0)) Then
                Call ucTreeView1.DeleteNode(hNodeDrag)
            End If
          
        '-- Not there, check format
        Else
            
            Select Case True
                
                '-- Text format ?
                Case Data.GetFormat(vbCFText)
                    
                    '-- We can have multiple items (multi-select listbox)
                    sText() = Split(Data.GetData(vbCFText), vbCrLf)
                    
                    '-- Extract Drop info (node-over and insert-after flag)
                    Call .OLEGetDropInfo(hNodeDrop, bInstertAfter)
                    
                    '-- Let's insert our first string
                    If (.OLEDragInsertStyle = [disInsertMark]) Then
                        m_lKey = m_lKey + 1
                        hNodeInsert = .AddNode(hNodeDrop, IIf(bInstertAfter, [rNext], [rPrevious]), m_lKey, sText(0), 7, 7)
                      Else
                        m_lKey = m_lKey + 1
                        hNodeInsert = .AddNode(hNodeDrop, [rFirst], m_lKey, sText(0), 7, 7)
                    End If
                    '-- More strings ? (insert as [rNext] of previous inserted
                    For lIdx = 1 To UBound(sText())
                        m_lKey = m_lKey + 1
                        hNodeInsert = .AddNode(hNodeInsert, [rNext], m_lKey, sText(lIdx), 7, 7)
                    Next lIdx
                    
                    '-- Ensure visible last inserted
                    Call .EnsureVisible(hNodeInsert)

                '-- No more formats checked here
                '   This should be checked on OLEDragOver() event.
                '   There, you can change your cursor, cancel, etc.
                Case Else
                    
                    Call MsgBox("Data format not supported, sorry.", vbInformation)
            End Select
        End If
    End With
    
    lblNodeDropVal.Caption = vbNullString
    lblInsertAfterVal.Caption = vbNullString
    lblNodeFullPathVal.Caption = vbNullString
End Sub

Private Sub pvCopyNode(ByVal hNodeFrom As Long, ByVal hNodeTo As Long, ByVal eRelation As tvRelationConstants)
'
' Important!
'
'   This sub. as well as next one, are using ucTreeView2 object to manage hNodes.
'   In fact, it doesn't matter which object we are using: hNodes are handles and our controls are
'   working as wrappers. But be careful: this only is true for hNode navigation! Forget to delete
'   nodes or perform any other operation that involves background operations with current ucTreeView
'   internal collection/array, as well as change any property related to appearance, style, etc.

    With ucTreeView2
        
        m_lKey = m_lKey + 1
        hNodeTo = .AddNode(hNodeTo, eRelation, m_lKey, .NodeText(hNodeFrom), .NodeImage(hNodeFrom), .NodeSelectedImage(hNodeFrom))
        
        Call .EnsureVisible(hNodeTo)
        Call pvCopyChildren(hNodeFrom, hNodeTo)
    End With
End Sub

Private Sub pvCopyChildren(ByVal hNodeFrom As Long, ByVal hNodeTo As Long)
  
  Dim hNextFrom  As Long
  Dim hNextTo    As Long
        
    With ucTreeView2
    
        hNextFrom = .NodeChild(hNodeFrom)
        hNextTo = hNodeTo
        
        Do While hNextFrom
            
            m_lKey = m_lKey + 1
            hNodeTo = .AddNode(hNextTo, , m_lKey, .NodeText(hNextFrom), .NodeImage(hNextFrom), .NodeSelectedImage(hNextFrom))
        
            Call pvCopyChildren(hNextFrom, hNodeTo)
            hNextFrom = .NodeNextSibling(hNextFrom)
        Loop
    End With
End Sub










'========================================================================================
' Forget this...
'========================================================================================

Private Sub pvFlattenButton(ByVal hButton As Long)
    
  Dim lS As Long

    lS = GetWindowLong(hButton, GWL_STYLE)
    Call SetWindowLong(hButton, GWL_STYLE, lS Or BS_FLAT)
End Sub

Private Sub pvNavigate(ByVal sURL As String)
    
    '-- Open URL
    Call ShellExecute(Me.hWnd, "open", sURL, vbNullString, vbNullString, SW_SHOW)
End Sub
