VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListC 
   Caption         =   "Customer List for Searching - by: pagut2000@yahoo.com"
   ClientHeight    =   8235
   ClientLeft      =   825
   ClientTop       =   1785
   ClientWidth     =   10665
   Icon            =   "frmListC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10665
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Option1"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListC.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListC.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListC.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListC.frx":0ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListC.frx":102A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListC.frx":1186
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7860
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CustID"
         Object.Width           =   2117
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Company Name"
         Object.Width           =   6174
         ImageIndex      =   6
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Address"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "City"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Postal Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Telp"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Contact Name"
         Object.Width           =   2540
         ImageIndex      =   2
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Contact's Phone"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search on:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmListC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' By : Paul PS, 11 Juli 2004
' References ADO 2.5 Library
' Database name : GATUK.MDB - MS ACCESS 2000
' email : pagut2000@yahoo.com
' Facilities:
' 1. Search Customer by Customer ID
' 2. Search Customer by the others criteria
' 3. Sort the item by cliking columheader of ListView Control
'    (Ascending or Descending)
' Please enjoy n part of your project database
' Good luck
'------------------------------------

Option Explicit

Private strCon As ADODB.Connection
Private rs As ADODB.Recordset
Private strString As String
Private sDaftar As ListItem


Private Sub Form_Load()
    
    Set strCon = New ADODB.Connection
    
    strString = "PROVIDER=Microsoft.Jet.OLEDB.4.0; " & _
         "Persist Security Info = False; " & _
        "Data Source=" & App.Path & "\GATUK.MDB"
    
    strCon.Open strString
    
    
    Set rs = New ADODB.Recordset
    
    With rs
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open "customers", strCon
    End With
   
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
        rs.MoveFirst
    End If
    
    FillCust
    
    Option1(0).Caption = "by Customer ID"
    Option1(1).Caption = "by Others item"
    
    rs.Close
    strCon.Close
    
End Sub

Private Sub Form_Resize()
    
    If Me.Width <= 10000 Then
        Me.Width = 10000
        Exit Sub
    End If
    cmdSearch.Left = Me.Width - 1200
    Label1(0).Width = Me.Width
    
    If Me.Height <= 3000 Then
        Me.Height = 3000
        Exit Sub
    End If
    ListView1.Move 0, 800, Me.Width - 100, Me.Height - 1600
    
End Sub

Sub FillCust()
    Dim x As Long
   
    ListView1.ListItems.Clear
    While Not rs.EOF
        Set sDaftar = ListView1.ListItems.Add(, , rs(0), , 0)
        For x = 1 To 7
            If IsNull(rs.Fields(x).Value) = False Then sDaftar.SubItems(x) = rs.Fields(x).Value
        Next x
        rs.MoveNext
    Wend
End Sub

Private Sub cmdSearch_Click()
   Dim intSelectedOption As Integer
   Dim strFindMe As String
   
   If Option1(0).Value = True Then
      strFindMe = InputBox("Search on : " & Option1(0).Caption, "Search")
      intSelectedOption = lvwText
   End If
   If Option1(1).Value = True Then
      strFindMe = InputBox("Search on: " & Option1(1).Caption, "Search")
      intSelectedOption = lvwSubItem
   End If

    If strFindMe = "" Then
        Exit Sub
    End If

   Dim itmFound As ListItem
   
   Set itmFound = ListView1.FindItem(strFindMe, intSelectedOption, , lvwPartial)
   
   If itmFound Is Nothing Then
      MsgBox "Sorry, no match found. Thank U" & vbCrLf, vbInformation + vbOKOnly, "No Found"
      Exit Sub
   Else
       itmFound.EnsureVisible
       itmFound.Selected = True
       ListView1.SetFocus
   End If
End Sub

'Sort the Item by clicking columnHeader
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView1.Sorted And _
        ColumnHeader.Index - 1 = ListView1.SortKey Then
        ListView1.SortOrder = 1 - ListView1.SortOrder
    Else
        ListView1.SortOrder = lvwAscending
        ListView1.SortKey = ColumnHeader.Index - 1
    End If
    ListView1.Sorted = True
End Sub

Private Sub ListView1_LostFocus()
    Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(i).Selected = False
    Next i
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value = True Or _
        Option1(1).Value = True Then cmdSearch.Enabled = True
        
End Sub
