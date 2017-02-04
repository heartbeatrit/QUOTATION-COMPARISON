VERSION 5.00
Object = "{AE106030-6295-4032-B94A-58E066679281}#1.1#0"; "XButtons.ocx"
Begin VB.Form FrmPrintRep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5850
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Page orientation"
      Height          =   1095
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
      Begin VB.OptionButton optPortrait 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Portrait"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optLandScape 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Landscap"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.ComboBox cmbDesign 
      Height          =   360
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   885
      Width           =   3975
   End
   Begin XButtons.XButton cmdPrint 
      Height          =   705
      Left            =   1560
      TabIndex        =   8
      Top             =   3120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1244
      ForeColor       =   0
      Caption         =   "&Print"
      Picture         =   "FrmPrintRep.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ContainerID     =   5179040
   End
   Begin XButtons.XButton cmdCancel 
      Height          =   705
      Left            =   2805
      TabIndex        =   9
      Top             =   3120
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1244
      ForeColor       =   0
      Caption         =   "&Cancel"
      Picture         =   "FrmPrintRep.frx":055A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ContainerID     =   5179040
   End
   Begin VB.Label lblDesign 
      BackStyle       =   0  'Transparent
      Caption         =   "Design"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   885
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Printing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Tag             =   "NN"
      Top             =   75
      Width           =   2535
   End
   Begin VB.Image ImageBar 
      Height          =   375
      Left            =   -240
      Picture         =   "FrmPrintRep.frx":0AB4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18975
   End
End
Attribute VB_Name = "FrmPrintRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Public mPrinter As Printer, pageno As Integer
    Public serialno As Integer, firstpage As Boolean, lastpage As Boolean, multiplepage As Boolean
    Public mSrlCheck As String
    Dim gYearcode As String
    Dim mxTop As Integer, mBodyTop As Integer, mTopMargin As Integer
    Dim mLeftMargin As Integer, mBodyBottom As Integer
    Dim rsDocDesign1 As New ADODB.Recordset, rsDocDesign As New ADODB.Recordset
    Dim mBase1 As String, mHeader As Currency, mColHeader As Currency
    Dim mFooter As Currency, mEnd As Currency
    Dim mtable As String, mTableStk As String, mAddon As String
    Dim msqlstr As String, mDesignName As String
    Dim mSubType As String, mType As String, mSrl As String, mPrefix As String, mYearCode As String
    
Private Sub cmdDiscard_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub cmdPrint_Click()
    If cmbPrinter.Text = "" Then
        cmbPrinter.SetFocus
    End If
    
    If Me.Tag <> "" And cmbDesign.ListIndex = -1 Then
        MsgBox " Design Not Selected", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    For Each mPrinter In Printers
        If mPrinter.DeviceName = cmbPrinter.Text Then
            Set Printer = mPrinter
        End If
    Next
    If optPortrait.Value = True Then
        Printer.Orientation = 1
    Else
        Printer.Orientation = 2
    End If
    
    If Me.Tag <> "" Then
        mDesignName = cmbDesign.Text
        PrintVoucher
    End If
    
    Me.Tag = "S"
    Me.Hide
End Sub

Private Sub Form_Activate()
    If Me.Tag <> "" Then
        'Me.Tag = "MP00101042010MFG000023"     'SubType+Yearcode+Type+Srl
        mSubType = Left(Me.Tag, 2)
        mYearCode = Mid(Me.Tag, 3, 11)
        mType = Mid(Me.Tag, 14, 3)
        mSrl = Right(Me.Tag, 6)
        
        lblDesign.Visible = True
        cmbDesign.Visible = True
        'ComboAdd cmbDesign, "Select distinct Code from " + gDataBaseGenStr + "TfatFormats where SubType='" + mSubType + "' order by Code"
    End If
End Sub

Private Sub Form_Load()
    'FormSetup Me, False
    For Each mPrinter In Printers
        cmbPrinter.AddItem mPrinter.DeviceName
    Next
    'cmdPrint.ImageSetup "Select", "Print"
    'cmdCancel.ImageSetup "Cancel", "Cancel"
    mSubType = ""
End Sub

Private Function PrintVoucher()
    Printer.EndDoc
    firstpage = False
    multiplepage = False
    lastpage = False
    pageno = 0
End Function
Public Function ComboAdd(cmbData As ComboBox, msqlstr As String, Optional mItemDataField As String = "")
    cmbData.Clear
    Dim rscnn As New ADODB.Recordset, i As Integer
    rscnn.CursorLocation = adUseClient
    rscnn.Open msqlstr, FrmReport.cnnDataBase, adOpenStatic, adLockReadOnly, adCmdText
    If rscnn.RecordCount > 0 Then
        i = 0
        rscnn.MoveFirst
        Do Until rscnn.EOF
            cmbData.AddItem rscnn.Fields(0), i
            If mItemDataField <> "" Then
                cmbData.ItemData(i) = Val(rscnn.Fields(mItemDataField))
            End If
            i = i + 1
            rscnn.MoveNext
        Loop
    End If
    rscnn.Close
End Function
