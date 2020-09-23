VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Load Save DB Picture"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Build Query"
      Height          =   330
      Left            =   1800
      TabIndex        =   25
      Top             =   4020
      Width           =   1260
   End
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   2010
      TabIndex        =   22
      Top             =   5640
      Width           =   1230
      Begin VB.CommandButton Command6 
         Caption         =   "Close"
         Height          =   375
         Left            =   75
         TabIndex        =   23
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Save Picture"
      Height          =   1200
      Left            =   5010
      TabIndex        =   18
      Top             =   4545
      Width           =   1980
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1005
         TabIndex        =   24
         Top             =   270
         Width           =   870
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save into Database"
         Enabled         =   0   'False
         Height          =   375
         Left            =   105
         TabIndex        =   21
         Top             =   705
         Width           =   1770
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Browse"
         Enabled         =   0   'False
         Height          =   375
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Load Picture"
      Height          =   750
      Left            =   105
      TabIndex        =   17
      Top             =   5640
      Width           =   1860
      Begin VB.CommandButton Command2 
         Caption         =   "Load from Database"
         Enabled         =   0   'False
         Height          =   390
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   1620
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5010
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   510
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3480
      Width           =   4830
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2775
      TabIndex        =   13
      Top             =   2925
      Width           =   2160
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2775
      TabIndex        =   11
      Top             =   2595
      Width           =   2160
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   4365
      Width           =   4830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect to Database"
      Height          =   390
      Left            =   1335
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1350
      TabIndex        =   3
      Top             =   1560
      Width           =   3600
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1350
      TabIndex        =   2
      Top             =   1080
      Width           =   3585
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1350
      TabIndex        =   1
      Top             =   615
      Width           =   3600
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1350
      TabIndex        =   0
      Top             =   135
      Width           =   3600
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   5325
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      Height          =   1920
      Left            =   5280
      Top             =   2565
      Width           =   1710
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "The auto build query:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   4155
      Width           =   1500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Extra Condition (where Clause of the Query)"
      Height          =   195
      Left            =   105
      TabIndex        =   14
      Top             =   3240
      Width           =   3075
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Field Name (which stores picture)"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2970
      Width           =   2340
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Table Name (where picture is stored)"
      Height          =   195
      Left            =   105
      TabIndex        =   9
      Top             =   2670
      Width           =   2595
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   7005
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   570
      TabIndex        =   7
      Top             =   1635
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   195
      Left            =   465
      TabIndex        =   6
      Top             =   1110
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Database Name"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server Name"
      Height          =   195
      Left            =   330
      TabIndex        =   4
      Top             =   195
      Width           =   930
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Cn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim FileName As String

Private Sub Command1_Click()
    On Error GoTo ErrLine
    Dim ConStr As String
    ConStr = "Provider=SQLOLEDB.1; data source=" & Text1.Text & ";initial catalog=" & Text2.Text & ";uid=" & Text3.Text & ";pwd=" & Text4.Text
    Cn.Open ConStr
    MsgBox "Connection established with database.", vbInformation, App.ProductName
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text6.Enabled = True
    Text7.Enabled = True
    Text8.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = False
    Command7.Enabled = True
    Exit Sub
ErrLine:
    MsgBox "Connection could not be established with the database, please check your provided information.", vbCritical, App.ProductName
End Sub

Private Sub Command2_Click()
    Command3_Click
    LoadStudentPicture
End Sub

Private Sub Command3_Click()
    Text5.Text = "Select " & Text7.Text & " from " & Text6.Text & " where " & Text8.Text
End Sub

Private Sub Command4_Click()
    CommonDialog1.ShowOpen
    FileName = CommonDialog1.FileName
    Image1.Picture = LoadPicture(FileName)
    If FileName <> "" Then Command5.Enabled = True
End Sub

Private Sub Command5_Click()
    UPdateStudentPicture
    Command5.Enabled = False
End Sub

Private Sub Command6_Click()
    If MsgBox("Do you want to close the project?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Command7_Click()
    Image1.Picture = LoadPicture("")
    FileName = ""
End Sub

Private Sub Form_Load()
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
End Sub

Private Function LoadStudentPicture() As Boolean

    LoadStudentPicture = False
    
    On Error GoTo Label2
    
    Dim strStream As New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    
    Dim Qry As String
    Qry = Text5.Text
    Set Rs = Cn.Execute(Qry)
    If Not Rs.EOF Then
    
        If Not IsNull(Rs.Fields("pic").Value) Then
            If Err.Number <> 0 Then
                MsgBox "Problem in the Picture of the student"
                GoTo Label2
            End If
            strStream.Write Rs("pic")
            strStream.SaveToFile "C:\Temp.bmp", adSaveCreateOverWrite
            Image1.Picture = LoadPicture("C:\Temp.bmp")
            Kill ("C:\Temp.bmp")
            strStream.Close
    
            GoTo Label2
        End If
    End If
    Exit Function
Label1:
    Image1.Refresh
    Image1.Picture = Nothing
    LoadStudentPicture = True
Label2:
End Function

Private Function UPdateStudentPicture() As Boolean
    UPdateStudentPicture = False
    On Error GoTo Err_Line
    
    If Image1.Picture = 0 Or FileName = "" Then
        MsgBox "No picture to save.", vbInformation, App.ProductName
        Exit Function
    End If
    
    Set Rs = New ADODB.Recordset
    Dim strStream As New ADODB.Stream
    
    strStream.Type = adTypeBinary
    strStream.Open
    
    strStream.LoadFromFile FileName
    Rs.Open "Select " & Text7.Text & " from " & Text6.Text & " where " & Text8.Text, Cn, adOpenKeyset, adLockOptimistic
    If Not Rs.EOF Then
        Rs("pic") = strStream.Read
        Rs.Update
    End If
    UPdateStudentPicture = True
    MsgBox "The picture is successfully stored into the database.", vbInformation, App.ProductName
    FileName = ""
    Exit Function
Err_Line:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf & "Please check the auto build query is correct or not.", vbCritical, App.ProductName
End Function

