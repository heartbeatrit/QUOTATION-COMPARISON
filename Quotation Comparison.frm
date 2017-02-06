VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AE106030-6295-4032-B94A-58E066679281}#1.1#0"; "XButtons.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmReport 
   BackColor       =   &H00158C7D&
   Caption         =   "Rate Comparison"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnSelParty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Select Quotes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker FromDate 
      Height          =   375
      Left            =   11400
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   53084161
      CurrentDate     =   41176
   End
   Begin MSComCtl2.DTPicker ToDate 
      Height          =   375
      Left            =   13320
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   53084161
      CurrentDate     =   41176
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridQuote 
      Height          =   6375
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11245
      _Version        =   393216
      BackColor       =   16777215
      RowHeightMin    =   350
      BackColorFixed  =   6485218
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLines       =   2
      MergeCells      =   1
      AllowUserResizing=   3
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin XButtons.XButton cmdPrint 
      Height          =   615
      Left            =   13440
      TabIndex        =   7
      Top             =   7200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      ForeColor       =   0
      Caption         =   "&Print"
      Picture         =   "QUOTAT~1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ContainerID     =   787394
   End
   Begin VB.Label lblDemo 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo Version Expires on 31/01/2017"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   12960
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   10320
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   1200
      Y1              =   8640
      Y2              =   9240
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -120
      X2              =   19080
      Y1              =   8480
      Y2              =   8480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00158C7D&
      Caption         =   "Rate Comparison"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00158C7D&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   19095
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public cnnDataBase As New ADODB.Connection
    Public cnnGeneral As New ADODB.Connection
    Public gUser As String, gSal As String
    Public gBranch As String, gYearcode As String, gBranchName As String
    Public gHWSerial As String
    Public gDataBase As String, gDataBaseGen As String, gDataBaseGenStr As String
    Public gActiveDate As String
    Public gCurrName As String
    Public gTokenNumber As Integer
    Public gSQLServer As String
    Public rsData As New ADODB.Recordset
    Public msqlstr As String
    Public BtnFlag As String
    Public mFromdate As String, mToDate As String
    Public Cnn As New ADODB.Connection, mSqlServer As String
    Public rsmain As New ADODB.Recordset, mPWD As String
    Public mSrlCheck As String
    Dim mSrl As String, mDocDate As String, mPlant As String
    Dim mDocdatePrint As String
    Dim AppFlag As String
Private Sub Enable()
    'BtnRefresh.Enabled = True
    'BtnSelParty.Enabled = True
    'BtnSelProduct.Enabled = True
    'BtnUpdate.Enabled = True
    'BtnUnAuthorise.Enabled = True
End Sub
Private Sub disable()
    'BtnRefresh.Enabled = False
    'BtnSelParty.Enabled = False
    'BtnSelProduct.Enabled = False
    'BtnUpdate.Enabled = False
    'BtnUnAuthorise.Enabled = False
End Sub

Private Sub BtnDiscard_Click()
    PanClose
End Sub
Private Sub PanClose()
'    PanSelectQuote.Visible = False
'   BtnRefresh.Enabled = True
'    BtnSelParty.Enabled = True
'    BtnSelProduct.Enabled = True
'    BtnUpdate.Enabled = True
End Sub

Private Sub BtnItemAccept_Click()
'    PnlListItem.Visible = False
'    Enable
End Sub
Public Sub ClickPartyBtn()
    
End Sub
Private Sub BtnSelParty_Click()
    Dim SqlStr As String
    mFromdate = Format(FromDate.Value, "DD-mmm-YYYY")
    mToDate = Format(ToDate.Value, "DD-mmm-YYYY")
    GridQuote.Clear
    GridQuote.Rows = 2
    With frmpickdata
        .mMultiSelect = True
        '.mRepStr = "Select Distinct M.Name as 'Party Name',Q.Srl,Q.YearCode+Q.Type+Q.Prefix+Q.Srl as mSerial,Convert(varchar(10),Q.Docdate,103) as QtnDate, M.code as Code from Quote Q Inner join Quotestk on Q.Branch = Quotestk.Branch And Q.YearCode = QuoteStk.YEarcode And  Q.type = Quotestk.type And Q.Prefix = Quotestk.Prefix And Q.Srl = Quotestk.Srl Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' AND PENDING <> 0 And Q.Docdate Between '" + mFromdate + "' And '" + mToDate + "'"
        .mRepStr = "Select Distinct Indentstk.Srl as IndentNo,Indentstk.Code,convert(varchar(10),Indentstk.Docdate,103) as IndentDate,Indentstk.Yearcode+Indentstk.type+Indentstk.prefix+Indentstk.sno+Indentstk.Srl as Ind,AddonQP.F003 as ComparisionNo " _
         & " from Quote Q  Inner join  QuoteStk on Q.Branch = QuoteStk.Branch And Q.YearCode = QuoteStk.Yearcode And  Q.type = QuoteStk.type And Q.Prefix = QuoteStk.Prefix And Q.Srl = QuoteStk.Srl " _
         & " Inner Join Indentstk On Indentstk.branch = Q.Branch And Quotestk.Indnumber = Indentstk.YEarcode+Indentstk.type+IndentStk.prefix+Indentstk.Sno+Indentstk.Srl " _
         & " Inner join Master m on m.Code = Q.Code  " _
         & " Left Outer Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
         & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
         & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' AND Quotestk.PENDING <> 0 And Q.Docdate Between '" + mFromdate + "' And '" + mToDate + "' And Indentstk.Pending > 0 And IndentStk.CloseFlag=0 And Indentstk.Cancflag=0 "
        
        .Show vbModal
        Dim i As Integer
        i = 1
        
        GridQuote.Cols = 4
        If .Tag <> "" Then
    
            If FindValue("Select Isnull(Name,'') from SysObjects Where name = 'TmpQuote'") <> "" Then
                cnnDataBase.Execute "Drop table TmpQuote"
            End If
            If FindValue("Select Isnull(Name,'') from SysObjects Where name = 'TmpQuoteStk'") <> "" Then
                cnnDataBase.Execute "Drop table TmpQuoteStk"
            End If
            If FindValue("Select Isnull(Name,'') from SysObjects Where name = 'TmpStock'") <> "" Then
                cnnDataBase.Execute "Drop table TmpStock"
            End If
            
            Dim mIndentRec As New ADODB.Recordset
            cnnDataBase.Execute "Select * into TmpQuote from Quote Where 1 = 2"
            cnnDataBase.Execute "Select * into TmpQuoteStk from QuoteStk Where 1 = 2"
            
            SqlStr = "Select Distinct YearCode+Type+Prefix+Srl as Link from QuoteStk Where Branch = 'HO' and Indnumber in (" + .Tag + ")"
            mIndentRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
            If mIndentRec.EOF <> True Then
                Do While mIndentRec.EOF <> True
                    cnnDataBase.Execute "Insert into TmpQuote Select *  from Quote Q Where Q.YearCode+Q.Type+Q.Prefix+Q.Srl in ('" + mIndentRec!link + "')"
                    cnnDataBase.Execute "Insert into TmpQuoteStk Select *  from QuoteStk Q Where Q.YearCode+Q.Type+Q.Prefix+Q.Srl in ('" + mIndentRec!link + "')"
                    mIndentRec.MoveNext
                Loop
            End If
            mIndentRec.Close
            
            SqlStr = "Select * into TmpStock From Stock Where Branch = 'HO' And Code in (Select Distinct Code From TmpQuoteStk) "
            cnnDataBase.Execute SqlStr
            
            Dim mRs As New ADODB.Recordset
            SqlStr = "Select Distinct IndentStk.Srl as Indnumber,Cast(TmpQuoteStk.Narr as varchar(254)) as INarr, (Select Name From Itemmaster Where code = TmpQuoteStk.Code) as Descr," _
             & "  case when TmpQuoteStk.xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='OD'),'') + ' : ' + TmpQuoteStk.XOD +" _
                            & "   (Case when TmpQuoteStk.xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='THK'),'') +' : ' + TmpQuoteStk.XTHK else '' end) + " _
                            & "   (case when TmpQuoteStk.xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='LENGTH'),'') +' : ' + TmpQuoteStk.XLength else '' end)+ " _
                            & "   (case when TmpQuoteStk.xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='Grade'),'') +' : ' + TmpQuoteStk.XGrade  else '' end)+" _
                            & "     (case when TmpQuoteStk.xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='SPEC'),'') +' : '+ TmpQuoteStk.XSPEC  else '' end)+" _
                            & "     (case when TmpQuoteStk.xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='NOI'),'') +' : ' + TmpQuoteStk.XNOI   else '' end)+" _
                            & "     (case when TmpQuoteStk.xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='WFI'),'') +' : ' + TmpQuoteStk.XWFI else '' end) else '' End as Alloc, " _
             & " TmpQuoteStk.Qty,TmpQuoteStk.Unit,TmpQuotestk.Qty2,TmpQuotestk.Unit2 from TmpQuote Q Inner join  TmpQuoteStk on Q.Branch = TmpQuoteStk.Branch And Q.YearCode = TmpQuoteStk.Yearcode And  Q.type = TmpQuoteStk.type And Q.Prefix = TmpQuoteStk.Prefix And Q.Srl = TmpQuoteStk.Srl " _
             & " Inner join Master m on m.Code = Q.Code  Inner Join indentstk On INdentStk.Branch = 'HO' And  TmpQuotestk.Indnumber = Indentstk.YEarcode+Indentstk.type+IndentStk.prefix+Indentstk.Sno+Indentstk.Srl  " _
             & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' Order By TmpQuoteStk.Qty"
            
            GridQuote.TextMatrix(0, 0) = "SR. NO."
            GridQuote.TextMatrix(0, 1) = "DESCR"
            GridQuote.TextMatrix(0, 2) = "QTY "
            
            GridQuote.Col = 1
            GridQuote.Row = 0
            GridQuote.CellFontBold = True
            GridQuote.Col = 2
            GridQuote.Row = 0
            GridQuote.CellFontBold = True
            
            mRs.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
            If mRs.EOF <> True Then
                Do While mRs.EOF <> True
                    GridQuote.TextMatrix(0, 1) = "DESCRIPTION   |" + "  Indent No. : " + mRs!Indnumber
                    GridQuote.TextMatrix(i, 0) = i
                    GridQuote.TextMatrix(i, 1) = mRs!Descr + "       " + mRs!Alloc + "  " + mRs!INarr
                    GridQuote.TextMatrix(i, 2) = CStr(Round(mRs!Qty, 2)) + "  " + mRs!unit
                    If InStr(1, UCase(GridQuote.TextMatrix(i, 1)), "LAFA") <> 0 Then
                        GridQuote.TextMatrix(i, 2) = GridQuote.TextMatrix(i, 2) + "  " + CStr(Round(mRs!Qty2, 2)) + "  " + mRs!unit2
                    End If
                    GridQuote.ColWidth(1) = 4500
                    GridQuote.RowHeight(i) = 800
                    GridQuote.WordWrap = True
                    GridQuote.Col = 1
                    GridQuote.Row = i
                    GridQuote.CellFontBold = False
                    
                    GridQuote.Col = 2
                    GridQuote.Row = i
                    GridQuote.CellFontBold = False
                    GridQuote.WordWrap = True
                    i = i + 1
                    GridQuote.Rows = GridQuote.Rows + 1
                    mRs.MoveNext
                Loop
            End If
            mRs.Close
            Dim x As Integer, mRow As Integer
            
            GridQuote.ColWidth(2) = 650
            
           'Stock
                'GridQuote.Rows = GridQuote.Rows + 1
                GridQuote.ColWidth(0) = 500
    
                'GridQuote.Cols = GridQuote.Cols + 1
                x = 3
                GridQuote.Cols = GridQuote.Cols + 1
                GridQuote.TextMatrix(0, x) = "STOCK"
                Dim mStock As String
                For i = 1 To GridQuote.Rows - 1
                    If GridQuote.TextMatrix(i, 1) <> "" Then
                        If GridQuote.TextMatrix(i, 1) = "TAXABLE  " Then
                            Exit For
                        Else
                            'Dim mLastRate As String, mTouchvalue As String
                            Dim mReplace As String, mCode As String
                            
                            mReplace = FindValue("Select Cast(Narr as varchar(254)) From TmpQuoteStk Where (Select Name From Itemmaster Where code = TmpQuoteStk.Code) + '       '+" _
                            & "  case when xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='OD'),'') + ' : ' + TmpQuotestk.XOD +" _
                            & "   (Case when xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='THK'),'') +' : ' + TmpQuotestk.XTHK else '' end) + " _
                            & "   (case when xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='LENGTH'),'') +' : ' + TmpQuotestk.XLength else '' end)+ " _
                            & "   (case when xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='Grade'),'') +' : ' + TmpQuotestk.XGrade  else '' end)+" _
                            & "     (case when xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='SPEC'),'') +' : '+ TmpQuotestk.XSPEC  else '' end)+" _
                            & "     (case when xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='NOI'),'') +' : ' + TmpQuotestk.XNOI   else '' end)+" _
                            & "     (case when xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='WFI'),'') +' : ' + TmpQuotestk.XWFI else '' end) else '' End + '  ' + Cast(Narr as varchar(254)) = '" + GridQuote.TextMatrix(i, 1) + "'")
                         
                            mReplace = "  " + mReplace
                            mCode = ""
                            mCode = Replace(GridQuote.TextMatrix(i, 1), mReplace, "")
                            
                            mStock = FindValue("Select isnull(Sum(Qty),0) From TmpStock Inner JOin Itemmaster On TmpStock.Code = Itemmaster.Code Where Branch = 'HO' And Itemmaster.Name +'       '+ " _
                            & "  case when xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='OD'),'') + ' : ' + TmpStock.XOD +" _
                            & "   (Case when xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='THK'),'') +' : ' + TmpStock.XTHK else '' end) + " _
                            & "   (case when xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='LENGTH'),'') +' : ' + TmpStock.XLength else '' end)+ " _
                            & "   (case when xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='Grade'),'') +' : ' + TmpStock.XGrade  else '' end)+" _
                            & "     (case when xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='SPEC'),'') +' : '+ TmpStock.XSPEC  else '' end)+" _
                            & "     (case when xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='NOI'),'') +' : ' + TmpStock.XNOI   else '' end)+" _
                            & "     (case when xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='WFI'),'') +' : ' + TmpStock.XWFI else '' end) else '' End  = '" + mCode + "' And Docdate<='" + Format(Now, "DD-MMM-YYYY") + "' And Left(TmpStock.Authorise,1) = 'A' and TmpStock.NotinStock=0 and TmpStock.Store='100009'")
                            GridQuote.TextMatrix(i, x) = Round(Val(mStock), 2)
                            GridQuote.ColWidth(x) = 1000
                        End If
                    End If
                Next
                'GridQuote.Rows = GridQuote.Rows + 1
                GridQuote.ColWidth(0) = 500
    
                GridQuote.Cols = GridQuote.Cols + 1
                x = 5
                'GridQuote.Cols = GridQuote.Cols + 1
                GridQuote.TextMatrix(0, x) = "Last Purch. Rate"
                GridQuote.TextMatrix(0, x - 1) = "Last Purch. Date"
                
                For i = 1 To GridQuote.Rows - 1
                    If GridQuote.TextMatrix(i, 1) <> "" Then
                        If GridQuote.TextMatrix(i, 1) = "TAXABLE  " Then
                            Exit For
                        Else
                            Dim mLastRate As String, mTouchvalue As String
                            'Dim mReplace As String, mCode As String
                            mReplace = FindValue("Select Cast(Narr as varchar(254)) From TmpQuoteStk Where (Select Name From Itemmaster Where code = TmpQuoteStk.Code) + '       '+" _
                            & "  case when xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='OD'),'') + ' : ' + TmpQuotestk.XOD +" _
                            & "   (Case when xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='THK'),'') +' : ' + TmpQuotestk.XTHK else '' end) + " _
                            & "   (case when xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='LENGTH'),'') +' : ' + TmpQuotestk.XLength else '' end)+ " _
                            & "   (case when xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='Grade'),'') +' : ' + TmpQuotestk.XGrade  else '' end)+" _
                            & "     (case when xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='SPEC'),'') +' : '+ TmpQuotestk.XSPEC  else '' end)+" _
                            & "     (case when xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='NOI'),'') +' : ' + TmpQuotestk.XNOI   else '' end)+" _
                            & "     (case when xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuotestk.code and Allocation ='WFI'),'') +' : ' + TmpQuotestk.XWFI else '' end) else '' End  + '  ' + Cast(Narr as varchar(254)) = '" + GridQuote.TextMatrix(i, 1) + "'")
                         
                            mReplace = "  " + mReplace
                            mCode = ""
                            mCode = Replace(GridQuote.TextMatrix(i, 1), mReplace, "")
                            
                            mLastRate = FindValue("Select Top 1 isnull(PurcRate,0) From ItemRates Inner JOin Itemmaster On ItemRates.Code = Itemmaster.Code Where Branch = 'HO' And Itemmaster.Name +' '+" _
                            & "  case when xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = ItemRates.code and Allocation ='OD'),'') + ' : ' + ItemRates.XOD +" _
                            & "   (Case when xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = ItemRates.code and Allocation ='THK'),'') +' : ' + ItemRates.XTHK else '' end) + " _
                            & "   (case when xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = ItemRates.code and Allocation ='LENGTH'),'') +' : ' + ItemRates.XLength else '' end)+ " _
                            & "   (case when xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = ItemRates.code and Allocation ='Grade'),'') +' : ' + ItemRates.XGrade  else '' end)+" _
                            & "     (case when xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = ItemRates.code and Allocation ='SPEC'),'') +' : '+ ItemRates.XSPEC  else '' end)+" _
                            & "     (case when xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = ItemRates.code and Allocation ='NOI'),'') +' : ' + ItemRates.XNOI   else '' end)+" _
                            & "     (case when xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = ItemRates.code and Allocation ='WFI'),'') +' : ' + ItemRates.XWFI else '' end) else '' End  = '" + mCode + "' Order By ItemRates.Touchvalue Desc ")
                            
                            mTouchvalue = FindValue("Select Top 1 Docdate From TmpStock Inner Join Itemmaster On TmpStock.Code = Itemmaster.Code Where Branch = 'HO' And TmpStock.Subtype = 'IC' And Itemmaster.Name +' '+" _
                            & "  case when xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='OD'),'') + ' : ' + TmpStock.XOD +" _
                            & "   (Case when xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='THK'),'') +' : ' + TmpStock.XTHK else '' end) + " _
                            & "   (case when xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='LENGTH'),'') +' : ' + TmpStock.XLength else '' end)+ " _
                            & "   (case when xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='Grade'),'') +' : ' + TmpStock.XGrade  else '' end)+" _
                            & "     (case when xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='SPEC'),'') +' : '+ TmpStock.XSPEC  else '' end)+" _
                            & "     (case when xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='NOI'),'') +' : ' + TmpStock.XNOI   else '' end)+" _
                            & "     (case when xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='WFI'),'') +' : ' + TmpStock.XWFI else '' end) else '' End  = '" + mCode + "' Order By TmpStock.Touchvalue Desc ")
                            
                            If mTouchvalue <> "" Then
                                mTouchvalue = Format(mTouchvalue, "DD/MM/YYYY")
                            End If
                            
                            GridQuote.TextMatrix(i, x - 1) = mTouchvalue
                            GridQuote.TextMatrix(i, x) = IIf(mLastRate = "", 0, Round(Val(mLastRate), 2))
                            GridQuote.ColWidth(x) = 1000
                            GridQuote.ColWidth(x - 1) = 1000
                        End If
                    End If
                Next
                
                'Last party
                GridQuote.Cols = GridQuote.Cols + 1
                x = GridQuote.Cols - 1
                GridQuote.TextMatrix(0, x) = "Last Purch. Party"
                For i = 1 To GridQuote.Rows - 1
                    If GridQuote.TextMatrix(i, 1) <> "" Then
                        If GridQuote.TextMatrix(i, 1) = "TAXABLE  " Then
                            Exit For
                        Else
                            Dim mLastParty As String
                           mReplace = FindValue("Select Cast(Narr as varchar(254)) From TmpQuoteStk Where (Select Name From Itemmaster Where code = TmpQuoteStk.Code) + '       '+" _
                            & "  case when xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='OD'),'') + ' : ' + TmpQuoteStk.XOD +" _
                            & "   (Case when xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='THK'),'') +' : ' + TmpQuoteStk.XTHK else '' end) + " _
                            & "   (case when xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='LENGTH'),'') +' : ' + TmpQuoteStk.XLength else '' end)+ " _
                            & "   (case when xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='Grade'),'') +' : ' + TmpQuoteStk.XGrade  else '' end)+" _
                            & "     (case when xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='SPEC'),'') +' : '+ TmpQuoteStk.XSPEC  else '' end)+" _
                            & "     (case when xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='NOI'),'') +' : ' + TmpQuoteStk.XNOI   else '' end)+" _
                            & "     (case when xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='WFI'),'') +' : ' + TmpQuoteStk.XWFI else '' end) else '' End + '  ' + Cast(Narr as varchar(254)) = '" + GridQuote.TextMatrix(i, 1) + "'")
            
                            mReplace = "  " + mReplace
                            mCode = ""
                            mCode = Replace(GridQuote.TextMatrix(i, 1), mReplace, "")
                            
                            
                            mLastParty = FindValue("Select Top 1 isnull(TmpStock.Party,0) From TmpStock  Inner Join Itemmaster On TmpStock.Code = Itemmaster.Code Where Branch = 'HO' And Subtype  in ('RP','IM') And Itemmaster.Name +' '+ " _
                            & "  case when xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='OD'),'') + ' : ' + TmpStock.XOD +" _
                            & "   (Case when xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='THK'),'') +' : ' + TmpStock.XTHK else '' end) + " _
                            & "   (case when xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='LENGTH'),'') +' : ' + TmpStock.XLength else '' end)+ " _
                            & "   (case when xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='Grade'),'') +' : ' + TmpStock.XGrade  else '' end)+" _
                            & "     (case when xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='SPEC'),'') +' : '+ TmpStock.XSPEC  else '' end)+" _
                            & "     (case when xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='NOI'),'') +' : ' + TmpStock.XNOI   else '' end)+" _
                            & "     (case when xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpStock.code and Allocation ='WFI'),'') +' : ' + TmpStock.XWFI else '' end) else '' End = '" + mCode + "' Order By TmpStock.Touchvalue Desc ")

                            GridQuote.TextMatrix(i, x) = CStr(FindValue("Select isnull(Name,'') From Master Where code = '" + mLastParty + "'"))
                            GridQuote.ColWidth(x) = 1000
                        End If
                    End If
                Next
    
            Dim n As Integer
            
            n = 7
            
            Dim mRecordSet As New ADODB.Recordset
            Dim mQuoteStk As New ADODB.Recordset
            
            
            SqlStr = "Select Distinct YearCode+Type+Prefix+Srl as Link from QuoteStk Where Branch = 'HO' and Indnumber in (" + .Tag + ")"
            mIndentRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
            If mIndentRec.EOF <> True Then
                Do While mIndentRec.EOF <> True
                        SqlStr = "Select Distinct Q.YearCode+Q.Type+Q.Prefix+Q.Srl as Link From TmpQuote Q Where Branch = 'HO' And Q.YearCode+Q.Type+Q.Prefix+Q.Srl in ('" + mIndentRec!link + "') "
                        mRecordSet.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mRecordSet.EOF <> True Then
                            Do While mRecordSet.EOF <> True
                                SqlStr = "Select Distinct M.Name as PartyName,Q.Srl as QSrl from TmpQuote Q Inner join TmpQuoteStk on Q.Branch = TmpQuoteStk.Branch And Q.YearCode = TmpQuoteStk.Yearcode And  Q.type = TmpQuoteStk.type And Q.Prefix = TmpQuoteStk.Prefix And Q.Srl = TmpQuoteStk.Srl Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' AND PENDING <> 0 And Q.YearCode+Q.Type+Q.Prefix+Q.Srl in ('" + mRecordSet!link + "') Order By 2"
                                mRs.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                                i = 1
                                If mRs.EOF <> True Then
                                    Do While mRs.EOF <> True
                                        If n >= GridQuote.Cols Then
                                            GridQuote.Cols = GridQuote.Cols + 1
                                        End If
                                        'GridQuote.TextMatrix(0, n) = mRs!PartyName + " # " + mRs!QSrl
                                        GridQuote.TextMatrix(0, n) = mRs!PartyName + "QtnNo :" + mRs!QSrl
                                        GridQuote.Col = n
                                        GridQuote.Row = 0
                                        GridQuote.CellFontBold = True
                                        GridQuote.WordWrap = True
                                        GridQuote.RowHeight(0) = 700
                                        GridQuote.ColWidth(n) = 2500
                                        
                                   For i = 1 To GridQuote.Rows - 1
                                        If GridQuote.TextMatrix(i, 1) <> "" Then
                                            SqlStr = "Select TmpQuoteStk.Type,TmpQuotestk.Rate,TmpQuoteStk.Qty*TmpQuoteStk.Rate as Val,TmpQuoteStk.Disc from TmpQuote Q Inner join TmpQuoteStk on Q.Branch = TmpQuoteStk.Branch And Q.YearCode = TmpQuoteStk.Yearcode And  Q.type = TmpQuoteStk.type And Q.Prefix = TmpQuoteStk.Prefix And Q.Srl = TmpQuoteStk.Srl Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' AND PENDING <> 0 And " _
                                                & " Q.YearCode+Q.Type+Q.Prefix+Q.Srl in ('" + mRecordSet!link + "') And (Select Name From Itemmaster Where code = TmpQuoteStk.Code)+'       '+ " _
                                                & "  case when xod<>'' then  ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='OD'),'') + ' : ' + TmpQuoteStk.XOD +" _
                                                & "   (Case when xthk <> '' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='THK'),'') +' : ' + TmpQuoteStk.XTHK else '' end) + " _
                                                & "   (case when xlength<>'' then  ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='LENGTH'),'') +' : ' + TmpQuoteStk.XLength else '' end)+ " _
                                                & "   (case when xgrade<>'' then   ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='Grade'),'') +' : ' + TmpQuoteStk.XGrade  else '' end)+" _
                                                & "     (case when xspec<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='SPEC'),'') +' : '+ TmpQuoteStk.XSPEC  else '' end)+" _
                                                & "     (case when xnoi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='NOI'),'') +' : ' + TmpQuoteStk.XNOI   else '' end)+" _
                                                & "     (case when xwfi<>'' then ' ' +ISNULL((Select Top 1 Descr From Allocated Where Code = TmpQuoteStk.code and Allocation ='WFI'),'') +' : ' + TmpQuoteStk.XWFI else '' end) else '' End +'  ' + cast(TmpQuoteStk.Narr as varchar(254))  = '" + GridQuote.TextMatrix(i, 1) + "' "
                                            mQuoteStk.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                                            If mQuoteStk.EOF <> True Then
                                                Do While mQuoteStk.EOF <> True
                                                    'GridQuote.TextMatrix(i, 0) = mRecordSet!link
                                                    GridQuote.TextMatrix(i, n) = Round(mQuoteStk!Rate, 2)
                                                    GridQuote.Col = n
                                                    GridQuote.Row = 1
                                                    GridQuote.CellFontBold = False
                                                    GridQuote.WordWrap = True
                                                    If n + 1 >= GridQuote.Cols Then
                                                        GridQuote.Cols = GridQuote.Cols + 1
                                                    End If
                                                    
                                                    GridQuote.Cols = GridQuote.Cols + 1
                                                    n = n + 1
                                                    GridQuote.TextMatrix(0, n) = "VALUE "
                                                    GridQuote.TextMatrix(i, n) = Round(mQuoteStk!Val, 2)
                                                    
                                                    GridQuote.Cols = GridQuote.Cols + 1
                                                    n = n + 1
                                                    ''''disc
                                                    GridQuote.TextMatrix(0, n) = "Disc(%) "
                                                    GridQuote.TextMatrix(i, n) = mQuoteStk!Disc
                                                    
                                                    GridQuote.Cols = GridQuote.Cols + 1
                                                    n = n + 1
                                                    ''''disc
                                                    GridQuote.TextMatrix(0, n) = "Disc Amt "
                                                    GridQuote.TextMatrix(i, n) = IIf(Val(mQuoteStk!Disc) = 0, 0, Round(mQuoteStk!Val * mQuoteStk!Disc / 100, 2))
                                                    
                                                    GridQuote.Cols = GridQuote.Cols + 1
                                                    n = n + 1
                                                    ''''disc
                                                    GridQuote.TextMatrix(0, n) = "Net Value "
                                                    GridQuote.TextMatrix(i, n) = Val(mQuoteStk!Val) - Round(mQuoteStk!Val * mQuoteStk!Disc / 100, 2)
                                                    
                                                    
                                                    GridQuote.Col = n + 1
                                                    GridQuote.Row = 1
                                                    GridQuote.CellFontBold = False
                                                    GridQuote.WordWrap = True
                                                    mQuoteStk.MoveNext
                                                Loop
                                            End If
                                            mQuoteStk.Close
                                        End If
                                   Next
                                        'n = n + 2
                                        n = n + 1
                                        'GridQuote.Cols = GridQuote.Cols + 1
                                        mRs.MoveNext
                                    Loop
                                End If
                                mRs.Close
                                mRecordSet.MoveNext
                            Loop
                        End If
                        mRecordSet.Close
                    mIndentRec.MoveNext
                 Loop
            End If
            
            GridQuote.Cols = GridQuote.Cols - 1
            n = 7
            Dim mChargesRec As New ADODB.Recordset
            'GridQuote.Rows = GridQuote.Rows + 1
            Dim mRowFix As Integer, mLastSrl As String, mLastSrlNEw As String
            
            'GridQuote.Rows = GridQuote.Rows + 1
            
            'Excise
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                GridQuote.TextMatrix(mRow, 1) = "QTN. NO. "
                For x = 1 To GridQuote.Cols - 1
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Srl as QtnNo from TmpQuote Q Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name+'QtnNo :'+ Q.srl in ('" + GridQuote.TextMatrix(0, x - 4) + "') "
                              
                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!QtnNo
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                GridQuote.Rows = GridQuote.Rows + 1
                
                
                For x = 0 To 0
                    For i = 7 To GridQuote.Cols - 1
                        If InStr(1, GridQuote.TextMatrix(x, i), "QtnNo :") <> 0 Then
                            GridQuote.TextMatrix(x, i) = Trim(Left(GridQuote.TextMatrix(x, i), InStr(1, GridQuote.TextMatrix(x, i), "QtnNo :") - 1))
                        End If
                    Next
                Next
                Dim xyz As Integer
                xyz = 0
                'MAKE
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                For x = 1 To GridQuote.Cols - 1
                    If xyz = 0 Then
                     If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                        xyz = x
                     Else
                        xyz = 0
                     End If
                    End If
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Isnull((select Distinct Head from Addons Where Branch = 'HO' and subtype = 'QP' and Fld = 'F055') ,'') Head,AddonQP.F055 from TmpQuote Q Inner join Master m on m.Code = Q.Code  " _
                                & " Inner Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
                                & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
                                & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "') and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "')"

                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                GridQuote.TextMatrix(mRow, 1) = UCase(mChargesRec!Head)
                                GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!F055
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
'            GridQuote.Rows = GridQuote.Rows + 1
'
'            mRow = GridQuote.Rows - 1
'            xyz = 0
'            'GridQuote.Rows = GridQuote.Rows + 1
'            For i = 1 To GridQuote.Rows - 1
'                For x = 1 To GridQuote.Cols - 1
'                    If xyz = 0 Then
'                    If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
'                        xyz = x
'                     Else
'                        xyz = 0
'                     End If
'                    End If
'                     If GridQuote.TextMatrix(0, x) = "VALUE " Then
'                        SqlStr = "Select TmpQuoteStk.Srl,Sum((TmpQuoteStk.Qty*Rate)*TmpQuoteStk.Disc/100) as DiscAmt from TmpQuote Q Inner join TmpQuoteStk on Q.Branch = TmpQuoteStk.Branch And Q.YearCode = TmpQuoteStk.Yearcode And  Q.type = TmpQuoteStk.type And Q.Prefix = TmpQuoteStk.Prefix And Q.Srl = TmpQuoteStk.Srl Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' AND PENDING <> 0 And " _
'                            & " M.Name in ('" + GridQuote.TextMatrix(0, x - 1) + "')  and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "') Group BY TmpQuoteStk.Srl"
'                        mQuoteStk.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
'                        If mQuoteStk.EOF <> True Then
'                            Do While mQuoteStk.EOF <> True
'                                'GridQuote.TextMatrix(mRow, 0) = mQuoteStk!Srl
'                                GridQuote.TextMatrix(mRow, 1) = "ITEM DISCOUNT"
'                                GridQuote.TextMatrix(mRow, x) = mQuoteStk!DiscAmt * -1
'                                GridQuote.Col = x
'                                GridQuote.Row = mRow
'                                GridQuote.CellFontBold = False
'                                GridQuote.WordWrap = True
'                             mQuoteStk.MoveNext
'                            Loop
'                        End If
'                        mQuoteStk.Close
'                    End If
'                Next
'            Next
        
            GridQuote.Rows = GridQuote.Rows + 1
            mRowFix = GridQuote.Rows - 1
            mRow = GridQuote.Rows - 1
            'Charges
            xyz = 0
            'GridQuote.Rows = GridQuote.Rows + 1
                For x = 1 To GridQuote.Cols - 1
                     
                     If xyz = 0 Then
                        If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                               xyz = x
                        Else
                               xyz = 0
                        End If
                     End If
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Distinct Fld,Head from Charges where Branch = 'HO' and Doctype in (select Top 1 Type From TmpQuoteStk)  and Fld <>'V001'  order by Fld  "
                        mChargesRec.CursorLocation = adUseServer
                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                Dim mVal As String
                                If InStr(1, mChargesRec!Fld, "00") <> 0 Then
                                    mVal = Replace(mChargesRec!Fld, "00", "AL")
                                Else
                                    mVal = Replace(mChargesRec!Fld, "0", "AL")
                                End If
                                
                                SqlStr = "Select Q.Srl,Q." + mVal + " as Val from TmpQuote Q Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "') and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "') Order By Q.srl"
                                
                                mQuoteStk.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                                'mQuoteStk.Close
                                If mQuoteStk.EOF <> True Then
                                    Do While mQuoteStk.EOF <> True
                                        If mLastSrl = "" Or mLastSrl = mQuoteStk!Srl Then
                                            'GridQuote.TextMatrix(mRow, 0) = mQuoteStk!Srl
                                            GridQuote.TextMatrix(mRow, 1) = UCase(mChargesRec!Head)
                                            GridQuote.TextMatrix(mRow, x) = Round(mQuoteStk!Val, 2)
                                            GridQuote.Col = x
                                            GridQuote.Row = mRow
                                        Else
                                            mRow = mRowFix
                                            'GridQuote.TextMatrix(mRow, 0) = mQuoteStk!Srl
                                            GridQuote.TextMatrix(mRow, 1) = UCase(mChargesRec!Head)
                                            GridQuote.TextMatrix(mRow, x) = Round(mQuoteStk!Val, 2)
                                            GridQuote.Col = x
                                            GridQuote.Row = mRow
                                            mLastSrlNEw = CStr(mLastSrlNEw) + CStr(mQuoteStk!Srl)
                                        End If
                                        GridQuote.CellFontBold = False
                                        GridQuote.WordWrap = True
                                        mLastSrl = mQuoteStk!Srl
                                        
                                        mQuoteStk.MoveNext
                                    Loop
                                End If
                                mQuoteStk.Close
                                mRow = mRow + 1
                                If Len(mLastSrlNEw) <= 6 Then
                                    Dim mRecord As String
                                    mRecord = Val(FindValue("Select Count(Fld) as Nos from Charges where Branch = 'HO' and Doctype in (select Top 1 Type From TmpQuoteStk)  and Fld <>'V001'"))
                                    GridQuote.Rows = GridQuote.Rows + mRecord
                                    mLastSrlNEw = "ABCDEFGHIJKLMN"
                                End If
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                'Net Amount
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                GridQuote.Rows = GridQuote.Rows + 1
                xyz = 0
                GridQuote.TextMatrix(mRow, 1) = "NET AMOUNT  "
                For x = 1 To GridQuote.Cols - 1
                    If xyz = 0 Then
                        If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                               xyz = x
                        Else
                               xyz = 0
                        End If
                     End If
                    If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Sum(Taxable)-Sum(ExcAmt) as Taxable from TmpQuote Q Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "') and Q.srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "') Group BY TaxCode "
                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                'GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!TaxCode
                                GridQuote.TextMatrix(mRow, x) = Round(mChargesRec!Taxable, 2)
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                'Excise
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                xyz = 0
                GridQuote.TextMatrix(mRow, 1) = "EXCISE "
                For x = 1 To GridQuote.Cols - 1
                     If xyz = 0 Then
                        If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                               xyz = x
                        Else
                               xyz = 0
                        End If
                     End If
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select ExcAmt/Amt1*100 as Per,ExcAmt as ExcAmt,isnull((select Top 1 ExciseAs from stockexcise where SubType = 'QP' and Branch = 'HO' And YearCode  = Q.YearCode and Type = Q.Type " _
                                & " and Prefix = Q.Prefix And Srl = Q.Srl),0) as ExAs from TmpQuote Q Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "')  and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "')"
                              
                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                If mChargesRec!exAs = 0 Then
                                    GridQuote.TextMatrix(mRow, x - 1) = ""
                                    GridQuote.TextMatrix(mRow, x) = Round(mChargesRec!ExcAmt, 2)
                                Else
                                    If mChargesRec!exAs = 2 And mChargesRec!ExcAmt = 0 Then
                                        GridQuote.TextMatrix(mRow, x - 1) = "EXCISE INCLUDED"
                                        GridQuote.TextMatrix(mRow, x) = Round(mChargesRec!ExcAmt, 2)
                                    Else
                                        GridQuote.TextMatrix(mRow, x - 1) = IIf(mChargesRec!ExcAmt = 0, "", "12.5")
                                        GridQuote.TextMatrix(mRow, x) = Round(mChargesRec!ExcAmt, 2)
                                    End If
                                End If
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                GridQuote.Rows = GridQuote.Rows + 1
                
                'taxable
                xyz = 0
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                GridQuote.Rows = GridQuote.Rows + 1
                
                GridQuote.TextMatrix(mRow, 1) = "TAXABLE "
                For x = 1 To GridQuote.Cols - 1
                     If xyz = 0 Then
                        If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                               xyz = x
                        Else
                               xyz = 0
                        End If
                     End If
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        
                        SqlStr = "Select Sum(Taxable) as Taxable from TmpQuote Q Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "')   and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "') Group BY TaxCode "
                              
                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                'GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!TaxCode
                                GridQuote.TextMatrix(mRow, x) = Round(mChargesRec!Taxable, 2)
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                'VAT
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                GridQuote.Rows = GridQuote.Rows + 1
                xyz = 0
                GridQuote.TextMatrix(mRow, 1) = "VAT / CST / SERVICE TAX  ****  CREDIT / NON CREDIT "
                For x = 1 To GridQuote.Cols - 1
                     If xyz = 0 Then
                        If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                               xyz = x
                        Else
                               xyz = 0
                        End If
                     End If
                     
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        
                        SqlStr = "Select Q.TaxCode,sum(Q.TaxAmt+Q.Addtax) as TaxAmt from TmpQuote Q Inner join Master m on m.Code = Q.Code  Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "')  and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "') Group BY TaxCode "
                              
                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                If mChargesRec!TaxCode = "15%11" Then
                                    GridQuote.TextMatrix(mRow, x - 1) = "15%"
                                ElseIf mChargesRec!TaxCode = "22%11" Then
                                    GridQuote.TextMatrix(mRow, x - 1) = "22%"
                                ElseIf mChargesRec!TaxCode = "5%11" Then
                                    GridQuote.TextMatrix(mRow, x - 1) = "5%"
                                ElseIf mChargesRec!TaxCode = "2%" Then
                                    GridQuote.TextMatrix(mRow, x - 1) = "C.S.T 2%"
                                Else
                                    GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!TaxCode
                                End If
                                GridQuote.TextMatrix(mRow, x) = Round(mChargesRec!TaxAmt, 2)
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                
                'Total
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                GridQuote.Col = 1
                GridQuote.Row = mRow
                xyz = 0
                GridQuote.CellFontBold = True
                GridQuote.TextMatrix(mRow, 1) = "TOTAL "
                
                For x = 1 To GridQuote.Cols - 1
                    If xyz = 0 Then
                    If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                        xyz = x
                     Else
                        xyz = 0
                     End If
                    End If
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Q.Amt from TmpQuote Q Inner join Master m on m.Code = Q.Code  " _
                                & " Left Outer Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
                                & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
                                & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "')  and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "')"


                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                GridQuote.Col = x
                                GridQuote.Row = mRow
                                GridQuote.CellFontBold = True
                                GridQuote.TextMatrix(mRow, x) = Round(mChargesRec!Amt, 2)
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                GridQuote.Rows = GridQuote.Rows + 1
                'Payment Terms
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                xyz = 0
                GridQuote.TextMatrix(mRow, 1) = "PAYMENT TERMS "
                For x = 1 To GridQuote.Cols - 1
                     If xyz = 0 Then
                        If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                            xyz = x
                         Else
                            xyz = 0
                         End If
                    End If
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select AddonQP.F004 from TmpQuote Q Inner join Master m on m.Code = Q.Code  " _
                                & " Left Outer Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
                                & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
                                & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "') and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "')"


                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                               ' GridQuote.TextMatrix(mRow, 1) = "Payment Terms "
                                GridQuote.TextMatrix(mRow, x - 1) = IIf(IsNull(mChargesRec!F004) = True, "", mChargesRec!F004)
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                GridQuote.Rows = GridQuote.Rows + 1
                'Delivery Date
                xyz = 0
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                For x = 1 To GridQuote.Cols - 1
                    If xyz = 0 Then
                     If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                        xyz = x
                     Else
                        xyz = 0
                     End If
                    End If
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Isnull((select Distinct Head from Addons Where Branch = 'HO' and subtype = 'QP' and Fld = 'f005') ,'') Head,AddonQP.F005 from TmpQuote Q Inner join Master m on m.Code = Q.Code  " _
                                & " Inner Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
                                & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
                                & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "')  and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "')"

                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                GridQuote.TextMatrix(mRow, 1) = UCase(mChargesRec!Head)
                                GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!F005
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                GridQuote.Rows = GridQuote.Rows + 1
                'F054
                xyz = 0
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                For x = 1 To GridQuote.Cols - 1
                    If xyz = 0 Then
                     If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                        xyz = x
                     Else
                        xyz = 0
                     End If
                    End If
                     If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Isnull((select Distinct Head from Addons Where Branch = 'HO' and subtype = 'QP' and Fld = 'f054') ,'') Head,AddonQP.F054 from TmpQuote Q Inner join Master m on m.Code = Q.Code  " _
                                & " Inner Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
                                & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
                                & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "')  and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "')"

                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                GridQuote.TextMatrix(mRow, 1) = UCase(mChargesRec!Head)
                                GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!F054
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                GridQuote.Rows = GridQuote.Rows + 1
                'F006
                mRowFix = GridQuote.Rows - 1
                xyz = 0
                mRow = GridQuote.Rows - 1
                For x = 1 To GridQuote.Cols - 1
                    If xyz = 0 Then
                     If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                        xyz = x
                     Else
                        xyz = 0
                     End If
                    End If
                    If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Isnull((select Distinct Head from Addons Where Branch = 'HO' and subtype = 'QP' and Fld = 'f006') ,'') Head,AddonQP.F006 from TmpQuote Q Inner join Master m on m.Code = Q.Code  " _
                                & " Inner Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
                                & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
                                & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "')  and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "')"

                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                GridQuote.TextMatrix(mRow, 1) = UCase(mChargesRec!Head)
                                GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!F006
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                'GridQuote.Rows = GridQuote.Rows + 1
                GridQuote.Rows = GridQuote.Rows + 1
                'Delivery Date
                xyz = 0
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                GridQuote.TextMatrix(mRow, 1) = "NET AMOUNT "
'                For x = 1 To GridQuote.Cols - 1
'                     If GridQuote.TextMatrix(0, x) = "Value " Then
'                        SqlStr = "Select Q.DeliveryDate from TmpQuote Q Inner join Master m on m.Code = Q.Code  " _
'                                & " Inner Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
'                                & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
'                                & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
'                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 1) + "') "
'
'                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
'                        If mChargesRec.EOF <> True Then
'                            Do While mChargesRec.EOF <> True
'                                GridQuote.TextMatrix(mRow, x - 1) = mChargesRec!DeliveryDate
'                                mChargesRec.MoveNext
'                            Loop
'                        End If
'                        mChargesRec.Close
'                    End If
'                Next
                
                
    
    
                x = GridQuote.Rows
                GridQuote.Rows = GridQuote.Rows + 1
                GridQuote.RowHeight(x) = 450
                
                GridQuote.TextMatrix(x, 1) = "EXECUTIVE (CONSU, PUR.,R.O.)"
                x = GridQuote.Rows
                GridQuote.Rows = GridQuote.Rows + 1
                GridQuote.RowHeight(x) = 450
                GridQuote.TextMatrix(x, 1) = "DY. MANAGER (CONSU, PUR.,R.O.)"
                x = GridQuote.Rows
                GridQuote.Rows = GridQuote.Rows + 1
                GridQuote.RowHeight(x) = 450
                GridQuote.TextMatrix(x, 1) = "DY. MANAGER (COMM.,R.O.)"
                'x = GridQuote.Rows
                'GridQuote.Rows = GridQuote.Rows + 1
                'GridQuote.TextMatrix(x, 1) = "SR. MANAGER"
                x = GridQuote.Rows
                GridQuote.Rows = GridQuote.Rows + 1
                GridQuote.RowHeight(x) = 450
                GridQuote.TextMatrix(x, 1) = "DIRECTOR (COMM.)"
                x = GridQuote.Rows
                GridQuote.Rows = GridQuote.Rows + 1
                GridQuote.RowHeight(x) = 450
                GridQuote.TextMatrix(x, 1) = "CEO / MD "
                GridQuote.RowHeight(0) = 900
                
                GridQuote.Rows = GridQuote.Rows + 1
                'Remarks
                Dim mRemarks As String
                mRowFix = GridQuote.Rows - 1
                mRow = GridQuote.Rows - 1
                For x = 1 To GridQuote.Cols - 1
                    If xyz = 0 Then
                     If GridQuote.TextMatrix(x, 1) = "QTN. NO. " Then
                        xyz = x
                     Else
                        xyz = 0
                     End If
                    End If
                    If GridQuote.TextMatrix(0, x) = "Net Value " Then
                        SqlStr = "Select Isnull((select Distinct Head from Addons Where Branch = 'HO' and subtype = 'QP' and Fld = 'f002') ,'') Head,AddonQP.F002 from TmpQuote Q Inner join Master m on m.Code = Q.Code  " _
                                & " Inner Join AddonQP On Q.Branch = AddonQP.Branch And Q.YearCode = AddonQP.YearCode And " _
                                & " Q.Type = AddonQP.Type And Q.Prefix = AddonQP.Prefix And Q.Srl = AddonQP.Srl " _
                                & " Where Q.Branch = 'HO' And Q.MainType='PR' And Q.code=m.code and Left(Q.Authorise,1) = 'A' And " _
                                    & " M.Name in ('" + GridQuote.TextMatrix(0, x - 4) + "')  and Q.Srl in ('" + GridQuote.TextMatrix(xyz, x - 1) + "')"

                        mChargesRec.Open SqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
                        If mChargesRec.EOF <> True Then
                            Do While mChargesRec.EOF <> True
                                If mRemarks = "" Then
                                    mRemarks = mChargesRec!F002
                                    GridQuote.TextMatrix(mRow, 1) = UCase(mChargesRec!Head)
                                    GridQuote.TextMatrix(mRow, 2) = mRemarks
                                End If
                                mChargesRec.MoveNext
                            Loop
                        End If
                        mChargesRec.Close
                    End If
                Next
                
                For x = 2 To GridQuote.Cols - 1
                    If mRemarks <> "" Then
                        GridQuote.TextMatrix(mRow, x) = mRemarks
                    End If
                Next
                For x = 2 To GridQuote.Col - 1
                    GridQuote.Col = x
                    GridQuote.Row = mRow
                    GridQuote.MergeCol(x) = True
                    GridQuote.MergeRow(mRow) = True
                Next
        End If
    End With
    Unload frmpickdata
End Sub
'Private Sub PopupSelect(Command As String, msqlstr As String)
'   'PnlListItem.Visible = True
'   'LblHeader.Caption = Command
'   'LVItem.ListItems.Clear
'   'connectionOPen
'    Dim rsData As New ADODB.Recordset
'        rsData.CursorLocation = adUseClient
'        rsData.Open msqlstr, cnnDataBase, adOpenStatic, adLockReadOnly, adCmdText
'
'        With rsData
'            If .RecordCount > 0 Then
'               ' LVItem.ListItems.Clear
'                .MoveFirst
'                Do While rsData.EOF = False
'                    Set Objlist = LVItem.ListItems.Add(, , "") ' Adds Items to List
'                    Objlist.SubItems(1) = !Name
'                    Objlist.SubItems(2) = !Srl
'                    Objlist.SubItems(3) = !Mserial
'                    Objlist.Tag = !Mserial
'                    .MoveNext
'                Loop
'            Else
'                'LVItem.ListItems.Clear
'            End If
'        End With
'    rsData.Close
   'ConnectionClose
'End Sub
 

Private Sub cmdPrint_Click()
    Dim mRsNew As New ADODB.Recordset
    mSrlCheck = ""
    mSrl = ""
    mPlant = ""
    mDocdatePrint = ""
    mRsNew.Open "Select Distinct Branch,YearCode,Type,Prefix,Srl From TmpQuote", cnnDataBase, adOpenDynamic, adLockReadOnly
    
    If mRsNew.EOF <> True Then
        Do While mRsNew.EOF <> True
            If mSrlCheck = "" Then
                mSrlCheck = FindValue("Select Max(isnull(Cast(F003 as MOney),'')) From AddonQP Where Branch='" + mRsNew!Branch + "'")
                mSrlCheck = GetSnoIncrease(mSrlCheck, 6)
            End If
            mDocDate = Format(Now, "DD/MM/YYYY")
            If FindValue("Select isnull(F003,'') From AddonQP Where Branch='" + mRsNew!Branch + "' And YearCode = '" + mRsNew!YearCode + "' " _
                    & " And  Type ='" + mRsNew!Type + "'  And  Prefix ='" + mRsNew!Prefix + "'  And  Srl ='" + mRsNew!Srl + "'") = "" Then
                    
                cnnDataBase.Execute "Update AddonQP Set F003 = '" + mSrlCheck + "' Where Branch='" + mRsNew!Branch + "' And YearCode = '" + mRsNew!YearCode + "' " _
                    & " And  Type ='" + mRsNew!Type + "'  And  Prefix ='" + mRsNew!Prefix + "'  And  Srl ='" + mRsNew!Srl + "'"
                
            Else
                mSrl = FindValue("Select isnull(F003,'') From AddonQP Where Branch='" + mRsNew!Branch + "' And YearCode = '" + mRsNew!YearCode + "' " _
                    & " And  Type ='" + mRsNew!Type + "'  And  Prefix ='" + mRsNew!Prefix + "'  And  Srl ='" + mRsNew!Srl + "'")
            End If
            
            'If mDocdatePrint = "" Or mDocdatePrint = " " Then
                    mDocdatePrint = FindValue("Select isnull(F008,'') From AddonQP Where Branch='" + mRsNew!Branch + "' And YearCode = '" + mRsNew!YearCode + "' " _
                            & " And  Type ='" + mRsNew!Type + "'  And  Prefix ='" + mRsNew!Prefix + "'  And  Srl ='" + mRsNew!Srl + "'")
            'End If
            If mDocdatePrint = "" Or mDocdatePrint = " " Then
                cnnDataBase.Execute "Update AddonQP Set F008 = '" + mDocDate + "' Where Branch='" + mRsNew!Branch + "' And YearCode = '" + mRsNew!YearCode + "' " _
                    & " And  Type ='" + mRsNew!Type + "'  And  Prefix ='" + mRsNew!Prefix + "'  And  Srl ='" + mRsNew!Srl + "'"
                mDocdatePrint = mDocDate
            Else
                'mDocdatePrint = FindValue("Select isnull(F008,'') From AddonQP Where Branch='" + mRsNew!Branch + "' And YearCode = '" + mRsNew!YearCode + "' " _
                '    & " And  Type ='" + mRsNew!Type + "'  And  Prefix ='" + mRsNew!Prefix + "'  And  Srl ='" + mRsNew!Srl + "'")
            End If
            
            If mPlant = "" Or mPlant = " " Then
                mPlant = FindValue("Select isnull(F007,'') From AddonQP Where Branch='" + mRsNew!Branch + "' And YearCode = '" + mRsNew!YearCode + "' " _
                    & " And  Type ='" + mRsNew!Type + "'  And  Prefix ='" + mRsNew!Prefix + "'  And  Srl ='" + mRsNew!Srl + "'")
            End If
            
            'Else
            '    cnnDataBase.Execute "Update AddonQP Set F003 = '" + mSrlCheck + "' Where Branch='" + mRsNew!Branch + "' And YearCode = '" + mRsNew!YearCode + "' " _
            '        & " And  Type ='" + mRsNew!Type + "'  And  Prefix ='" + mRsNew!Prefix + "'  And  Srl ='" + mRsNew!Srl + "'"
            'End If
            mRsNew.MoveNext
        Loop
    End If
    mRsNew.Close
    
    FrmPrintRep.Show vbModal
    If FrmPrintRep.Tag <> "" Then GridPrint GridQuote
    Unload FrmPrintRep
End Sub

Private Sub Form_Load()
    Cnn.Open "Provider=Microsoft.jet.oledb.4.0;" & "Data Source=" + App.Path + "\MergeGeneralData.mdb;ms access;Pwd=Parmeshwar"
    rsmain.CursorLocation = adUseClient
   ' FrmReport.Height = Screen.Height - 100
   ' FrmReport.Width = Screen.Width - 100
    
    rsmain.Open "select * from sqlserver", Cnn, adOpenStatic, adLockReadOnly, adCmdText
        If rsmain.EOF = False Then
            gDataBase = rsmain.Fields("databasename")
            gDataBaseGen = rsmain.Fields("nTfatset")
            gDataBaseGenStr = rsmain.Fields("nTfatset") + ".dbo."
            mSqlServer = rsmain.Fields("servername")
            gSQLServer = mSqlServer
            'gBranch = rsmain.Fields("branch")
            gHWSerial = rsmain!HWSerial
         mPWD = IIf(IsNull(rsmain!Password) = True, "", rsmain!Password)
        End If
    rsmain.Close
    Cnn.Close
    initlist
    connectionOPen
    'GvQoutes.BackColor = vbYellow
    'FromDate.SetFocus
    AppFlag = 0
    FromDate.Value = Now
    ToDate.Value = Now
    mFromdate = Format(FromDate.Value, "DD-mmm-YYYY")
    mToDate = Format(ToDate.Value, "DD-mmm-YYYY")
    
    
    SaveSetting "RateComparison", "Tfat", "Date", "31-Jan-2017"
    Dim MDATE As String
    MDATE = GetSetting("RateComparison", "Tfat", "Date")

    Dim mDateCheck As String
    mDateCheck = DateAdd("D", 15, MDATE)
    
    lblDemo.Caption = "Demo Version Expires on " + MDATE
    lblDemo.Visible = True
    If DateValue(Format(mDateCheck, "DD-MMM-YYYY")) < DateValue(Format(Now, "DD-MMM-YYYY")) Then
        Unload Me
    End If
    'BtnRefresh_Click
End Sub
Private Sub connectionOPen()
Dim mconnstr As String
    mconnstr = "Provider=SQLOLEDB.1;User Id=sa;Password='" & mPWD & "';server=" + mSqlServer + ";Initial Catalog=" + gDataBase
    cnnDataBase.ConnectionString = mconnstr
    cnnDataBase.Open
    cnnDataBase.CommandTimeout = 500
End Sub
Private Sub ConnectionClose()
        cnnDataBase.Close
End Sub
Public Sub initlist()        ' Used to Inialize the properties of ListView
'    With LVItem
'      '  .View = lvwReport
'        .Appearance = ccFlat
'        .FullRowSelect = True
'        '.GvQouteslines = False
'        .HotTracking = True
'        .HoverSelection = True
'        .HideSelection = False
'        .ColumnHeaders.Add 1, "d1", "", 400
'        .ColumnHeaders.Add 2, "d2", "Name", 4500
'        .ColumnHeaders.Add 3, "d3", "Code", 1000
'        .ColumnHeaders.Add 4, "d4", "Prifix"
'    End With
End Sub


Public Function FindValue(ByVal xSqlStr As String)
    FindValue = ""
    Dim mRecordSet As New ADODB.Recordset
    If (xSqlStr) <> "" Then
        mRecordSet.Open xSqlStr, cnnDataBase, adOpenDynamic, adLockReadOnly
        If mRecordSet.EOF <> True Then
            FindValue = mRecordSet.Fields(0)
        End If
    End If
    If IsNull(FindValue) = True Then FindValue = ""
End Function
Public Function GridPrint(Grid As MSHFlexGrid)
    Dim mcurrentx As Currency, mcurrenty As Currency, mCellWidth As Integer
    Dim mTColWidth As Integer, mRowHeigh As Integer, mGridWidth As Integer
    Dim mColWidth As Integer, mScaleHeight As Integer, mScaleWidth As Integer
    Dim mRow As Integer, mxRow As Integer, myRow As Integer, mCol As Integer
    Dim LeftMargin As Integer, mtext As String, mTextLast As String, mPageNo As Integer, RepLeftMargin As Integer
    Dim mRepCode As String, mFooter As String, reptopmargin As Integer
    
    mRepCode = Grid.Tag
    mPageNo = 0
    RepLeftMargin = 250
    mScaleHeight = Printer.ScaleHeight
    mScaleWidth = Printer.ScaleWidth
    LeftMargin = RepLeftMargin
    mcurrentx = LeftMargin
    mCol = 0
    mTColWidth = 0
    Printer.Orientation = 2
    
   ' mFooter = FindData(gDataBaseGenStr + "TfatSearch", "SubCodeOf+' --> '+Code", "Code='" + mRepCode + "' And Ltrim(Rtrim(SubCodeOf))<>''")
    Do While mCol <> Grid.Cols
        mTColWidth = mTColWidth + IIf(Grid.ColWidth(mCol) = -1, 750, Grid.ColWidth(mCol))
        mCol = mCol + 1
    Loop
    Dim printerwidthnotset As Boolean
    printerwidthnotset = False
    
    If mTColWidth > mScaleWidth Then
       MsgBox "Print Width Not Proper. Please Resize Some Column.", vbOKOnly + vbInformation
       printerwidthnotset = True
       Exit Function
    End If
    
    Call GridPrintHeadingNoDatabase(mPageNo, Grid, mTColWidth, mScaleWidth)
    
    'Report Heading - Fixed row Print
    mcurrentx = RepLeftMargin
    mRow = 0
    mcurrenty = reptopmargin
    mcurrenty = 1500
    
    Printer.Line (LeftMargin, mcurrenty)-(mTColWidth + 150, mcurrenty)
    
    mRow = 0
    mcurrentx = LeftMargin
    
    Do While mRow <> Grid.FixedRows
        mCol = 0
        Printer.Line (mcurrentx, mcurrenty)-(mcurrentx, mcurrenty + Grid.RowHeight(mRow))
        'Printer.Line (LeftMargin, mcurrenty)-(mTColWidth, mcurrenty)
        Do While mCol <> Grid.Cols
            If IIf(Grid.ColWidth(mCol) = -1, 750, Grid.ColWidth(mCol)) > 150 Then
                Grid.Row = mRow
                Grid.Col = mCol
                mtext = Grid.TextMatrix(mRow, mCol)
                mColWidth = IIf(Grid.ColWidth(mCol) = -1, 750, Grid.ColWidth(mCol))
                Printer.Font = Grid.CellFontName
                Printer.FontSize = 8
                Printer.FontItalic = Grid.CellFontItalic
                Printer.FontBold = True
                Printer.CurrentY = mcurrenty + ((Grid.RowHeight(mRow) - Printer.TextHeight(Trim(mtext))) / 2)
                
                If mCol = 1 Then
                    mtext = "ITEM NAME "
                End If
                
                If Printer.TextWidth(mtext) > IIf(Grid.ColWidth(mCol) = -1, 750, Grid.ColWidth(mCol)) Then
                    Printer.CurrentX = mcurrentx + 100
                Else
                    Printer.CurrentX = mcurrentx + 200
                End If
                If Printer.TextWidth(mtext) > IIf(Grid.ColWidth(mCol) = -1, 750, Grid.ColWidth(mCol)) Then
                     Call GridPrintRowHeight(mRow, mCol, mtext, IIf(Grid.ColWidth(mCol) = -1, 750, Grid.ColWidth(mCol)), Printer.CurrentX, Printer.CurrentY)
                Else
                     Printer.Print mtext
                End If
                
                mcurrentx = mcurrentx + IIf(Grid.ColWidth(mCol) = -1, 750, Grid.ColWidth(mCol))
                Printer.Line (mcurrentx, mcurrenty)-(mcurrentx, mcurrenty + Grid.RowHeight(mRow))
                'Printer.Line (LeftMargin, mcurrenty)-(mTColWidth, mcurrenty)
            End If
            mCol = mCol + 1
        Loop
        mcurrentx = LeftMargin
        mcurrenty = mcurrenty + Grid.RowHeight(mRow)
        'Printer.Line (mcurrentx, mcurrenty)-(mcurrentx, mcurrenty + Grid.RowHeight(mRow))
        Printer.Line (LeftMargin, mcurrenty)-(mTColWidth + 150, mcurrenty)
        
        mRow = mRow + 1
    Loop
    
    'Report Body Print
    mRow = Grid.FixedRows
    Dim i As Integer, n As Integer, mMergeRowPrint As Boolean, mOldText As String
    For i = Grid.FixedRows To Grid.Rows - 1
        mCol = 0
        Printer.Line (mcurrentx, mcurrenty)-(mcurrentx, mcurrenty + Grid.RowHeight(mRow))
        mMergeRowPrint = False
        mOldText = ""
        For n = 0 To Grid.Cols - 1
            If IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n)) > 50 Then
                ''If grid.ColWidth(mcol) <> 0 Then
                mtext = Grid.TextMatrix(i, n)
                mColWidth = IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n))
                Grid.Col = n
                Grid.Row = i
                Printer.Font = Grid.CellFontName
                Printer.FontSize = Grid.CellFontSize
                Printer.FontBold = Grid.CellFontBold
                Printer.FontItalic = Grid.CellFontItalic
                Printer.CurrentX = mcurrentx + 50
                If Grid.RowHeight(i) > 350 Then
                   Printer.CurrentY = mcurrenty
                Else
                   Printer.CurrentY = mcurrenty + ((Grid.RowHeight(i) - Printer.TextHeight(Trim(mtext))) / 2)
                End If
                'If n = 12 Then
                If n <> 1 And n <> 6 And n <> 2 Then
                If Grid.ColAlignment(n) >= 7 And Grid.ColAlignment(n) <= 9 Then
                   Printer.CurrentX = mcurrentx + mColWidth - Printer.TextWidth(mtext) - 50
                End If
                End If
                'End If
                If Grid.MergeCol(n) = True And n <> 2 And n <> 6 Then
                    If Grid.TextMatrix(i, n - 1) <> mtext Then
                        Printer.Print mtext
                    'ElseIf Grid.TextMatrix(i, n - 1) <> mtext Then
                    '    Printer.Print mtext
                    Else
                        mOldText = mtext
                        If i < 14 Then Printer.Print mtext
                    End If
                ElseIf Grid.MergeRow(i) = True Then
                    'Changed on 10-02-2011
                    'If mMergeRowPrint = False Then
                    '    Printer.Print mtext
                    '    mMergeRowPrint = True
                    'End If
                    If mtext = mOldText Then
                        mtext = ""
                    Else
                        If n = 1 Then
                            Printer.Print mtext
                        Else
                            Printer.Print "                      " + mtext
                        End If
                        mOldText = mtext
                    End If
                Else
                    If Printer.TextWidth(mtext) > IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n)) Then
                       ' If FindData(gDataBaseGenStr + "TfatSearch", "FitToHeight", "Code='" + mRepCode + "' And Sno='" + CStr(n) + "'") = "True" Then
                            Printer.CurrentY = mcurrenty
                            Call GridPrintRowHeight(mRow, mCol, mtext, IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n)), Printer.CurrentX, Printer.CurrentY)
                       ' Else
                           ' Printer.Print mtext
                       ' End If
                    Else
                        Printer.Print mtext
                    End If
                End If
                
                
                'Horizontal Line
              '  Printer.Zoom = 60
                Printer.Line (mcurrentx, mcurrenty + Grid.RowHeight(i))-(mcurrentx + IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n)), mcurrenty + Grid.RowHeight(i))
                
                'Vertical Line
                'Printer.Line (mCurrentx + Grid.ColWidth(n), mCurrenty)-(mCurrentx + Grid.ColWidth(n), mCurrenty + Grid.RowHeight(i))
                'Changed on 12-05-2010
                'Printer.Line (mCurrentx + Grid.ColWidth(n), mCurrenty)-(mCurrentx + Grid.ColWidth(n), mCurrenty + Grid.RowHeight(i))
                If Grid.MergeRow(i) = True And n <> 1 Then
                    If n = Grid.Cols - 1 Then
                        Printer.Line (mcurrentx + IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n)), mcurrenty)-(mcurrentx + IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n)), mcurrenty + Grid.RowHeight(i))
                    End If
                Else
                    Printer.Line (mcurrentx + IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n)), mcurrenty)-(mcurrentx + IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n)), mcurrenty + Grid.RowHeight(i))
                End If
                
                mcurrentx = mcurrentx + IIf(Grid.ColWidth(n) = -1, 750, Grid.ColWidth(n))
                'mTextLast = Grid.TextMatrix(mRow, mCol)
            End If
            ''''jjmCol = mCol + 1
        Next
        mMergeRowPrint = False
        mcurrentx = LeftMargin
        mcurrenty = mcurrenty + Grid.RowHeight(i)
        
        
        '''jjjmRow = mRow + 1
        If mcurrenty >= (mScaleHeight - 1000) And i <> Grid.Rows Then
                       
            Printer.Line (LeftMargin, mcurrenty)-(mTColWidth, mcurrenty)
            Printer.CurrentY = mScaleHeight - 250
            Printer.CurrentX = RepLeftMargin
            Printer.FontBold = True
            Printer.FontSize = 6
            Printer.FontName = "Verdana"
           ' Printer.Print "T.fat                    " + mFooter
            
            Printer.CurrentY = mScaleHeight - 250
            Printer.CurrentX = RepLeftMargin + 800
            Printer.FontSize = 6
            Printer.FontName = "Verdana"
            'Printer.Line (mx, my)-(mx + mwidth, my + mheight), , B
            
            Printer.NewPage
            ''Call gridprintheading(rsrep, mpageno, grid)
            Call GridPrintHeadingNoDatabase(mPageNo, Grid, mTColWidth, mScaleWidth)
            mcurrentx = LeftMargin
            mcurrenty = 1300
            Printer.Line (LeftMargin, mcurrenty)-(mTColWidth, mcurrenty)
            mcurrentx = LeftMargin
            mxRow = 0
            Dim mxCol As Integer
            mxCol = 0
            Do While mxRow <> Grid.FixedRows
                mxCol = 0
                Printer.Line (mcurrentx, mcurrenty)-(mcurrentx, mcurrenty + Grid.RowHeight(mRow))
                Do While mxCol <> Grid.Cols
                    If IIf(Grid.ColWidth(mxCol) = -1, 750, Grid.ColWidth(mxCol)) > 150 Then
                        Grid.Col = mxCol
                        Grid.Row = mxRow
                        mtext = Grid.TextMatrix(mxRow, mxCol)
                        mColWidth = IIf(Grid.ColWidth(mxCol) = -1, 750, Grid.ColWidth(mxCol))
                        Printer.Font = Grid.CellFontName
                        Printer.FontSize = Grid.CellFontSize
                        Printer.FontItalic = Grid.CellFontItalic
                        Printer.FontBold = True
                        Printer.CurrentY = mcurrenty + ((Grid.RowHeight(mxRow) - Printer.TextHeight(Trim(mtext))) / 2)
                        Printer.CurrentX = mcurrentx + ((mColWidth - Printer.TextWidth(mtext)) / 2)
                        Printer.Print mtext
                        mcurrentx = mcurrentx + IIf(Grid.ColWidth(mxCol) = -1, 750, Grid.ColWidth(mxCol))
                        Printer.Line (mcurrentx, mcurrenty)-(mcurrentx, mcurrenty + Grid.RowHeight(mxRow))
                    End If
                    mxCol = mxCol + 1
                Loop
              '  Printer.Zoom = 60
                mcurrentx = LeftMargin
                mcurrenty = mcurrenty + Grid.RowHeight(mxRow)
                Printer.Line (LeftMargin, mcurrenty)-(mTColWidth, mcurrenty)
                mxRow = mxRow + 1
            Loop
        End If
    Next
        Printer.CurrentY = mcurrenty + 100
        Printer.CurrentX = RepLeftMargin + 500
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Tahoma"
        Printer.Print "PO NO : ______________________"
            
            
  If Dir(App.Path + "\GOLF.Inf") = "" Then
    If InStr(1, mRepCode, "Salary") <> 0 Or InStr(1, mRepCode, "Bonus") <> 0 Then
        Printer.CurrentY = mScaleHeight - 750
        Printer.CurrentX = RepLeftMargin
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "  ___________"
                
        Printer.CurrentY = mScaleHeight - 750
        Printer.CurrentX = RepLeftMargin + 4500
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "___________"
        
        Printer.CurrentY = mScaleHeight - 750
        Printer.CurrentX = RepLeftMargin + 9000
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "___________"
        
        
        Printer.CurrentY = mScaleHeight - 750
        Printer.CurrentX = RepLeftMargin + 13500
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "___________"
        
        '2
        Printer.CurrentY = mScaleHeight - 500
        Printer.CurrentX = RepLeftMargin
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print ""
                
        Printer.CurrentY = mScaleHeight - 500
        Printer.CurrentX = RepLeftMargin + 4500
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "Bipin Prajapati"
        
        Printer.CurrentY = mScaleHeight - 500
        Printer.CurrentX = RepLeftMargin + 9000
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "Shilpa Patel"
        
        
'        Printer.CurrentY = mScaleHeight - 500
'        Printer.CurrentX = RepLeftMargin + 13500
'        Printer.FontBold = True
'        Printer.FontSize = 8
'        Printer.FontName = "Verdana"
'        Printer.Print "Kunal Shah"
        
        
        '3
        Printer.CurrentY = mScaleHeight - 300
        Printer.CurrentX = RepLeftMargin
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "Dy. Manager"
                
        Printer.CurrentY = mScaleHeight - 300
        Printer.CurrentX = RepLeftMargin + 4500
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "Sr. Manager"
        
        Printer.CurrentY = mScaleHeight - 300
        Printer.CurrentX = RepLeftMargin + 9000
        Printer.FontBold = True
        Printer.FontSize = 8
        Printer.FontName = "Verdana"
        Printer.Print "  Director(Comm.)"
        
        
'        Printer.CurrentY = mScaleHeight - 300
'        Printer.CurrentX = RepLeftMargin + 13500
'        Printer.FontBold = True
'        Printer.FontSize = 8
'        Printer.FontName = "Verdana"
'        Printer.Print "MD & CEO"
        
    'Over
   Else
    
  End If
 End If
    Printer.CurrentY = mScaleHeight - 150
    Printer.CurrentX = RepLeftMargin
    Printer.FontBold = True
    Printer.FontSize = 6
    Printer.FontName = "Verdana"
    
    'Printer.Print "T.fat                    " + mFooter
    
    Printer.CurrentY = mScaleHeight - 150
    Printer.CurrentX = RepLeftMargin + 800
    Printer.FontSize = 6
    Printer.FontName = "Verdana"
    'Printer.Line (mx, my)-(mx + mwidth, my + mheight), , B
    'Printer.Print "Email Generate From T.fat ERP"
    'Printer.Line (leftmargin, mcurrenty)-(mtcolwidth, mcurrenty)
    Printer.EndDoc
End Function

Public Function GridPrintHeadingNoDatabase(mPageNo As Integer, gridrep As MSHFlexGrid, mTColWidth As Integer, mScaleWidth As Integer)
    Dim xleftmargin As Integer
    Dim xtopmargin As Integer, RepLeftMargin As Integer
    Dim my As Integer, reptopmargin As Integer
    Dim mx As Integer
    Dim gRepHeading1 As String, gRepHeading2 As String
        gRepHeading1 = "SURAJ LIMITED "
        gRepHeading2 = "Rate Comparision "
    RepLeftMargin = 250
    reptopmargin = 0
    mPageNo = mPageNo + 1
    mx = RepLeftMargin
    my = 100
    
    Printer.CurrentX = RepLeftMargin
    Printer.CurrentY = my
    Printer.FontName = "Tahoma"
    Printer.FontBold = True
    Printer.FontSize = 10
    Printer.FontBold = True
    my = my + 100
    Printer.CurrentX = RepLeftMargin
    Printer.CurrentY = my
    Dim myFix As Integer
    myFix = my
    
    Call Printpicture(my, RepLeftMargin, 500, 2000, "SL.jpg", App.Path, "")
    Printer.CurrentX = RepLeftMargin
    my = my + 500
    Printer.CurrentX = RepLeftMargin
    Printer.CurrentY = my
    Printer.Print "       THOL PLANT"
    Printer.FontBold = True
    Printer.FontSize = 9
    
    'Printer.CurrentX = RepLeftMargin
    '
    Dim m As Integer
    m = TextWidth("RATE COMPARISION ")
    Printer.CurrentX = mTColWidth - m
    Printer.CurrentY = myFix
    Printer.Print "RATE COMPARISION "
    Printer.FontBold = True
    Printer.FontSize = 9
    
    Printer.CurrentX = (mTColWidth / 2) - m
    Printer.CurrentY = myFix
    Printer.Print "F/PUR/03 REV NO. 08 "
    Printer.FontBold = True
    Printer.FontSize = 9
    
    
    
    Printer.CurrentX = mTColWidth - m
    Printer.CurrentY = myFix + 500
    Printer.Print "Sr. No. " + CStr(mSrl)
    Printer.FontBold = True
    Printer.FontSize = 9
    
    
    'my = my + 350
    Dim mRepstr1 As String
    my = my + 350
    
    Dim mStr As String
    
    mStr = Replace(GridQuote.TextMatrix(0, 1), "Description   |  ", "")
    Printer.CurrentX = RepLeftMargin
    Printer.CurrentY = my
    Printer.Print "Plant :  " + mPlant + "                      " + mStr + "                    Child Group : " + FindValue("Select isnull((Select Name From Itemmaster Where Code = i.Grp),'') From TmpQuoteStk Inner JOin Itemmaster I On TmpQuoteStk.Code = I.Code")
    Printer.CurrentX = mTColWidth - m
    Printer.CurrentY = my
    Printer.Print "Date : " + CStr(Format(mDocdatePrint, "DD/MM/YYYY"))
    'Printer.FontBold = True
    'my = my + 350
    'Printer.CurrentY = my
    
    'Printer.FontBold = False
    my = my + 700
    Printer.CurrentX = RepLeftMargin
    Printer.CurrentY = my
End Function

Public Function FitInHeight(mRepCode As String, Grid As MSHFlexGrid)
    Dim i As Integer, n As Integer
    Dim rsReport As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim mCol As Integer
    Dim mheading As Integer
    Dim msqlstr1 As String
    Dim mdataval As String
    Grid.Visible = True
    i = 2
    mCol = 0
    Dim mbodynarrheight As Integer
    'rsReport.CursorLocation = adUseClient
   ' rsReport.Open "Select * from " + gDataBaseGen + ".Dbo.TfatSearch  where Code='" + mRepCode + "' order by Sno", cnndata, adOpenStatic, adLockReadOnly, adCmdText
    i = 1
    For i = Grid.FixedRows To Grid.Rows - 1
     '   rsReport.MoveFirst
        mCol = 0
      '  Do While Not rsReport.EOF
            mCol = mCol + 1
            Grid.Col = mCol
            Grid.Row = i
            If rsReport.Fields("FitToHeight") = True Then
                Dim mtext As String
                Dim mWidth As Integer
                mbodynarrheight = 0
                mWidth = Grid.ColWidth(mCol)
                mtext = Grid.TextMatrix(i, mCol)
                Grid.Row = i
                Grid.Col = mCol
                Printer.FontName = Grid.CellFontName
                Printer.FontSize = Grid.CellFontSize + 1
                If Printer.TextWidth(mtext) > mWidth And InStr(1, mtext, " ") > 0 Then
                    Dim mStr As String
                    Dim mstr1 As String
                    Dim mline As String
                    Dim mPos As Integer
                    Dim mx1 As Integer, mx As Integer
                    Dim mLen As Integer
                    mStr = mtext
                    mLen = Len(mtext)
                    mline = ""
                    Do While mLen <> 0
                        mPos = InStr(1, mStr, " ")
                        mx1 = InStr(1, mStr, Chr(10))
                        If mx1 <> 0 And mx1 < mPos Then mbodynarrheight = mbodynarrheight + 250
                        If mPos <> 0 Then
                            mstr1 = Mid(mStr, 1, mPos)
                            If mPos = Len(mStr) Then
                                mStr = ""
                                mLen = 0
                            End If
                        Else
                            mPos = Len(mStr)
                            mstr1 = Mid(mStr, 1, mPos)
                            mLen = 0
                        End If
                        If mx1 = 0 And mPos = 0 Then mLen = 0
                        If mbodynarrheight > 32000 Then Exit Do
                        If (Printer.TextWidth(mline) + Printer.TextWidth(mstr1)) > mWidth Then
                            Printer.CurrentX = mx
                            mbodynarrheight = mbodynarrheight + 250
                            mline = ""
                        End If
                        mline = mline + mstr1
                        If (mLen - Len(mstr1)) <> 0 Then
                            mStr = Mid(mStr, mPos + 1, Len(mStr))
                        End If
                        If mLen = 0 Then
                            Printer.CurrentX = mx
                            mbodynarrheight = mbodynarrheight + 250
                        End If
                    Loop
                    Grid.Row = i
                    Grid.Col = mCol
                    If Grid.MergeCol(mCol) = False Then
                        If Grid.RowHeight(i) < mbodynarrheight Then Grid.RowHeight(i) = mbodynarrheight
                    Else
                        If Grid.TextMatrix(i - 1, mCol) <> Grid.TextMatrix(i, mCol) Then
                            If Grid.RowHeight(i) < mbodynarrheight Then Grid.RowHeight(i) = mbodynarrheight
                        End If
                    End If
                End If
            End If
    Next
    Grid.Visible = True
End Function

Public Function GridPrintRowHeight(mRow As Integer, mCol As Integer, mtext As String, mWidth As Integer, mcurx As Integer, mcury As Integer)
       If Printer.TextWidth(mtext) > mWidth Then
          Dim mStr As String
          Dim mstr1 As String
          Dim mline As String
          Dim mPos As Integer
          Dim mx1 As Integer, mbodynarrheight As Integer
          Dim mfindmx1 As Boolean
          Dim mLen As Integer
          mbodynarrheight = 0
          mStr = mtext
          mLen = Len(mtext)
          mline = ""
          mfindmx1 = False
          Do While mLen <> 0
             mPos = InStr(1, mStr, " ")
             mx1 = InStr(1, mStr, Chr(10))
             If mx1 <> 0 And mx1 < mPos Then
                mstr1 = Mid(mStr, 1, mx1 - 2)
                mPos = mx1
                mfindmx1 = True
             Else
                mfindmx1 = False
                If mPos <> 0 Then
                   mstr1 = Mid(mStr, 1, mPos)
                   If mPos = Len(mStr) Then
                      mStr = ""
                      mLen = 0
                   End If
                Else
                   mPos = Len(mStr)
                   mstr1 = Mid(mStr, 1, Len(mStr))
                   mStr = ""
                   mLen = 0
                End If
             End If
             If mPos = 0 Then mLen = 0
             If mfindmx1 = True Then
                mline = mline + mstr1
                Printer.CurrentX = mcurx
                mbodynarrheight = mbodynarrheight + 250
                mline = ""
                mstr1 = ""
             ElseIf (Printer.TextWidth(mline) + Printer.TextWidth(mstr1)) > mWidth Then
                Printer.CurrentX = mcurx
                Printer.Print mline
                mbodynarrheight = mbodynarrheight + 260
                mline = ""
             End If
             mline = mline + mstr1
             If (mLen - Len(mstr1)) <> 0 Then
                mStr = Mid(mStr, mPos + 1, Len(mStr))
             End If
             If mLen = 0 Then
                Printer.CurrentX = mcurx
                Printer.Print mline
                mbodynarrheight = mbodynarrheight + 260
             End If
          Loop
       End If
End Function

Public Function GetSnoIncrease(mSno As String, mLength As Integer, Optional mtable As String = "", Optional mField As String = "", Optional mCond As String = "", Optional mSkipLengthFill As Boolean = False)
    If mtable = "" Then
        mSno = CStr(Val(mSno) + 1)
    Else
        mSno = CStr(Val(FindData(mtable, "IsNull(Max(Cast(" + mField + " as int)),0)", mCond)) + 1)
    End If
    If Len(mSno) < mLength And mSkipLengthFill = False Then
        mSno = Replace(Space(mLength - Len(mSno)), " ", "0") + mSno
    End If
    GetSnoIncrease = mSno
End Function
Public Function FindData(mtable As String, mField As String, Optional mCond As String = "", Optional mdatatype As String = "T", Optional msqlstr As String = "", Optional mTSGeneral As Boolean = False)
    On Error GoTo mError
    If mdatatype = "N" Then
        mField = "IsNull(" + mField + ",0)"
    Else
        mField = "IsNull(" + mField + ",'')"
    End If
    
    If mTSGeneral = True Then mtable = "TSGeneral.Dbo." + mtable
    
    Dim rsFind As New ADODB.Recordset
    rsFind.CursorLocation = adUseClient
    rsFind.Open IIf(msqlstr <> "", msqlstr, "Select " + mField + " From " + mtable + IIf(mCond <> "", " Where " + mCond, "")), cnnDataBase, adOpenStatic, adLockReadOnly, adCmdText
    If rsFind.RecordCount > 0 Then
        FindData = rsFind.Fields(0)
    End If
    rsFind.Close
mError:
    If Err.Number <> 0 Then MsgBox Err.Description
End Function

Public Function Printpicture(mtop As Integer, mleft As Integer, mHeight As Integer, mWidth As Integer, mfldtext As String, mdatapath As String, mfldcond As String)
    FrmDesktop.ImagePicture.Picture = LoadPicture(mdatapath + "\" + mfldtext)
    Printer.PaintPicture FrmDesktop.ImagePicture.Picture, mleft, mtop, mWidth, mHeight
    FrmDesktop.Hide
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload frmpickdata
    Unload FrmPrintRep
    Unload FrmDesktop
    Unload FrmReport
End Sub

