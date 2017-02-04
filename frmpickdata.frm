VERSION 5.00
Object = "{AE106030-6295-4032-B94A-58E066679281}#1.1#0"; "XButtons.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmpickdata 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4965
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmpickdata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4320
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc rsRep 
      Height          =   330
      Left            =   7320
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridData 
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6376
      _Version        =   393216
      RowHeightMin    =   350
      BackColorFixed  =   14737632
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin XButtons.XButton cmdCancel 
      Height          =   465
      Left            =   8640
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   820
      ForeColor       =   0
      Caption         =   "&Cancel"
      Picture         =   "frmpickdata.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ContainerID     =   2819906
   End
   Begin XButtons.XButton cmdAccept 
      Height          =   465
      Left            =   7530
      TabIndex        =   1
      Top             =   4320
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   820
      ForeColor       =   0
      Caption         =   "&Accept"
      Picture         =   "frmpickdata.frx":0566
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ContainerID     =   2819906
   End
   Begin XButtons.XButton cmdSelectAll 
      Height          =   465
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   820
      ForeColor       =   0
      Caption         =   "&Select All"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ContainerID     =   2819906
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblheading 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Data"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   4815
      Left            =   75
      Top             =   75
      Width           =   9975
   End
End
Attribute VB_Name = "frmpickdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public mFormOpen As String
    Public mlinkdocno As String
    Public mPartycode As String
    Public mRepCode As String
    Public mRepStr As String
    Public mGridBase As Boolean
    Public mMultiSelect As Boolean
    Public mQty As Currency, mColSel As Integer, mBatch As Boolean, mqty2 As Currency, MDATE As String
    Dim mRowSel As Integer, i As Integer

Private Sub cmdaccept_Click()
    mQty = 0
    If mMultiSelect = True Then
        mlinkdocno = ""
        For i = 1 To GridData.Rows - 1
            GridData.Row = i
            GridData.Col = 1
            If GridData.CellBackColor = vbGreen Then
                 mlinkdocno = mlinkdocno + "'" + GridData.TextMatrix(i, 4) + "',"
            End If
        Next
        If mlinkdocno <> "" Then
            mlinkdocno = Mid(mlinkdocno, 1, Len(mlinkdocno) - 1)
        End If
    End If
    mMultiSelect = True
    Me.Tag = mlinkdocno
    Me.Hide
End Sub
Private Sub cmdCancel_Click()
    Unload frmpickdata
End Sub

Private Sub cmdSelectAll_Click()
    Dim x As Integer
    With GridData
        For i = .FixedRows To .Rows - 1
            .Row = i
            For x = 1 To .Cols - 1
                .Col = x
                If cmdSelectAll.Caption = "&Select All" Then
                    .CellBackColor = vbGreen
                Else
                    .CellBackColor = vbWhite
                End If
            Next
        Next
    End With
    If cmdSelectAll.Caption = "&Select All" Then
        cmdSelectAll.Caption = "&Unselect All"
    Else
        cmdSelectAll.Caption = "&Select All"
    End If
End Sub

Private Sub Form_Activate()
    pointerglass
    'formcenter frmpickdata
        mRepCode = "ABC"
      '  On Error Resume Next
        Dim rsRep As New ADODB.Recordset
        rsRep.CursorLocation = adUseClient
        rsRep.Open mRepStr, FrmReport.cnnDataBase, adOpenStatic, adLockReadOnly, adCmdText
        If rsRep.EOF = False Then
            Set GridData.DataSource = rsRep
            gridheading
        End If
        rsRep.Close
    GridData.ColWidth(0) = 0
    pointerdefault
    GridData.ColWidth(1) = 1200
    GridData.ColWidth(2) = 1500
    GridData.ColWidth(3) = 1200
    GridData.ColWidth(4) = 0
End Sub

Private Sub griddata_Click()
    'Set GridData.Sort = MSHierarchicalFlexGridLib.flexSortGenericAscending
     GridSorting GridData
End Sub

Private Sub griddata_DblClick()
    If mMultiSelect = True Then
        With GridData
            .Row = .RowSel
            For i = 1 To GridData.Cols - 1
                .Col = i
                If .CellBackColor = vbGreen Then
                    .CellBackColor = vbWhite
                Else
                    .CellBackColor = vbGreen
                End If
            Next
        End With
    Else
        cmdaccept_Click
    End If
End Sub

Private Function gridheading()
    With GridData
        If mRepCode = "PROCESSLIST" Then
            .ColWidth(0) = 0
            .ColWidth(1) = 2000
        ElseIf mRepCode = "IndentFetch" Then
            mRepStr = "Select I.YearCode,I.Type,I.Prefix,I.Sno,I.Srl,I.Sno,Convert(varchar(10),I.DocDate,103) as [Doc Date]" _
                & "I.Code as [Product COde],M.Name as [Product Name],I.xGrade as Grade,I.xOD as OD,I.xThk as Thickness,I.xLength as Length" _
                & " from IndentStk I "
            .ColWidth(0) = 0
            .ColWidth(1) = 500
            .ColWidth(2) = 0
            .ColWidth(3) = 0
            .ColWidth(4) = 900
            .ColWidth(5) = 1200
        End If
    End With
End Function

Private Sub txtSearch_Change()
    mColSel = GridData.ColSel
    mRowSel = GridData.RowSel
    For i = GridData.FixedRows To GridData.Rows - 1
        If InStr(UCase(GridData.TextMatrix(i, mColSel)), UCase(txtSearch.Text)) <> 0 Then
            GridData.Row = i
            GridData.Col = mColSel
            GridData.TopRow = i
            GridData.SetFocus
            Exit For
        End If
    Next
    txtSearch.SetFocus
End Sub

Public Function GridSorting(Grid As MSHFlexGrid)
    If Grid.Row = 1 And Grid.RowSel = Grid.Rows - 1 Then
        Grid.Sort = MSHierarchicalFlexGridLib.SortSettings.flexSortGenericAscending
    End If
End Function


Public Function pointerglass()
    Screen.MousePointer = vbHourglass
End Function

Public Function pointerdefault()
    Screen.MousePointer = vbDefault
End Function
Public Function formcenter(mFormOpen As Form)
    Dim mfheight As Integer
    Dim mfwidth As Integer
    mfheight = FrmReport.Height
    mfwidth = FrmReport.Width
    mFormOpen.Top = ((mfheight - mFormOpen.Height) / 2)
    mFormOpen.Left = ((mfwidth - mFormOpen.Width) / 2)
End Function

