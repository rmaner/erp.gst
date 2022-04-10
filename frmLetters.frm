VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmLetters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Letters..."
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   Icon            =   "frmLetters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8835
   Begin TabDlg.SSTab SSTab1 
      Height          =   7905
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   13944
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Letter Details..."
      TabPicture(0)   =   "frmLetters.frx":114DA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "rtb"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Preview..."
      TabPicture(1)   =   "frmLetters.frx":114F6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox rtb 
         Height          =   5640
         Left            =   -74925
         TabIndex        =   19
         Top             =   2205
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9948
         _Version        =   393217
         ScrollBars      =   3
         MousePointer    =   3
         DisableNoScroll =   -1  'True
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmLetters.frx":11512
      End
      Begin VB.Frame Frame6 
         Caption         =   "Memo Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -74910
         TabIndex        =   6
         Top             =   345
         Width           =   8625
         Begin VB.PictureBox pboxA 
            Height          =   1560
            Left            =   75
            ScaleHeight     =   1500
            ScaleWidth      =   8400
            TabIndex        =   7
            Top             =   210
            Width           =   8460
            Begin RichTextLib.RichTextBox txtSubject 
               Height          =   375
               Left            =   705
               TabIndex        =   20
               Top             =   720
               Width           =   7515
               _ExtentX        =   13256
               _ExtentY        =   661
               _Version        =   393217
               TextRTF         =   $"frmLetters.frx":11594
            End
            Begin VB.CommandButton cmdRender 
               DownPicture     =   "frmLetters.frx":11616
               Height          =   315
               Left            =   7350
               Picture         =   "frmLetters.frx":11959
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   30
               UseMaskColor    =   -1  'True
               Width           =   570
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "&Save"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   6300
               TabIndex        =   17
               Top             =   30
               Width           =   1035
            End
            Begin VB.TextBox txtLetterRef 
               Height          =   330
               Left            =   720
               TabIndex        =   10
               Top             =   30
               Width           =   990
            End
            Begin VB.CommandButton cmdSelect 
               Height          =   315
               Left            =   5730
               Picture         =   "frmLetters.frx":11CE3
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   30
               Width           =   570
            End
            Begin VB.CommandButton cmdNew 
               Caption         =   "&NEW"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4680
               TabIndex        =   8
               Top             =   30
               Width           =   1035
            End
            Begin MSComCtl2.DTPicker txtLetterDate 
               Height          =   330
               Left            =   2640
               TabIndex        =   11
               Top             =   30
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd-MM-yy hh:mm tt"
               Format          =   46137347
               CurrentDate     =   38023
            End
            Begin RichTextLib.RichTextBox txtSender 
               Height          =   375
               Left            =   705
               TabIndex        =   21
               Top             =   1095
               Width           =   7515
               _ExtentX        =   13256
               _ExtentY        =   661
               _Version        =   393217
               TextRTF         =   $"frmLetters.frx":12026
            End
            Begin RichTextLib.RichTextBox txtReceipient 
               Height          =   375
               Left            =   705
               TabIndex        =   22
               Top             =   360
               Width           =   7515
               _ExtentX        =   13256
               _ExtentY        =   661
               _Version        =   393217
               TextRTF         =   $"frmLetters.frx":120A8
            End
            Begin VB.Label Label1 
               Caption         =   "By:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   30
               TabIndex        =   16
               Top             =   1170
               Width           =   855
            End
            Begin VB.Label Label35 
               Caption         =   "Dated:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   2100
               TabIndex        =   15
               Top             =   75
               Width           =   1080
            End
            Begin VB.Label Label3 
               Caption         =   "Letter#:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   30
               TabIndex        =   14
               Top             =   60
               Width           =   855
            End
            Begin VB.Label Label35 
               Caption         =   "To:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   30
               TabIndex        =   13
               Top             =   435
               Width           =   1335
            End
            Begin VB.Label Label37 
               Caption         =   "Subject:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   15
               TabIndex        =   12
               Top             =   795
               Width           =   855
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7455
         Left            =   105
         TabIndex        =   1
         Top             =   360
         Width           =   8595
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   5670
            ScaleHeight     =   285
            ScaleWidth      =   2895
            TabIndex        =   2
            Top             =   135
            Width           =   2895
            Begin VB.CommandButton cmdChangePageOrientation 
               Caption         =   "Orientation"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1425
               TabIndex        =   4
               Top             =   0
               Width           =   1425
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "&Print"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   1425
            End
         End
         Begin VSPrinter8LibCtl.VSPrinter vp 
            Height          =   7260
            Left            =   60
            TabIndex        =   5
            Top             =   135
            Width           =   8475
            _cx             =   14949
            _cy             =   12806
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            MousePointer    =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoRTF         =   -1  'True
            Preview         =   -1  'True
            DefaultDevice   =   0   'False
            PhysicalPage    =   0   'False
            PalettePicture  =   "frmLetters.frx":1212A
            AbortWindow     =   -1  'True
            AbortWindowPos  =   0
            AbortCaption    =   "Printing..."
            AbortTextButton =   "Cancel"
            AbortTextDevice =   "on the %s on %s"
            AbortTextPage   =   "Now printing Page %d of"
            FileName        =   ""
            MarginLeft      =   1440
            MarginTop       =   720
            MarginRight     =   360
            MarginBottom    =   720
            MarginHeader    =   0
            MarginFooter    =   0
            IndentLeft      =   0
            IndentRight     =   0
            IndentFirst     =   0
            IndentTab       =   720
            SpaceBefore     =   0
            SpaceAfter      =   0
            LineSpacing     =   100
            Columns         =   1
            ColumnSpacing   =   180
            ShowGuides      =   2
            LargeChangeHorz =   300
            LargeChangeVert =   300
            SmallChangeHorz =   30
            SmallChangeVert =   30
            Track           =   0   'False
            ProportionalBars=   -1  'True
            Zoom            =   64
            ZoomMode        =   0
            ZoomMax         =   400
            ZoomMin         =   10
            ZoomStep        =   25
            EmptyColor      =   -2147483636
            TextColor       =   0
            HdrColor        =   0
            BrushColor      =   0
            BrushStyle      =   0
            PenColor        =   0
            PenStyle        =   0
            PenWidth        =   0
            PageBorder      =   0
            Header          =   ""
            Footer          =   ""
            TableSep        =   "|;"
            TableBorder     =   7
            TablePen        =   0
            TablePenLR      =   0
            TablePenTB      =   0
            NavBar          =   1
            NavBarColor     =   -2147483633
            ExportFormat    =   0
            URL             =   ""
            Navigation      =   3
            NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
            AutoLinkNavigate=   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
      End
   End
End
Attribute VB_Name = "frmLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNew_Click()
    X = MsgBox("Create new letter?", vbYesNo)
    If X = vbYes Then
        sSQL(0) = "ADDNEW_LETTER": dbOpen (0)
        Set recs(0) = recs(0).NextRecordset
        txtLetterRef.Text = recs(0)!LetterRef
        dbClose (0)
        sSQL(0) = ""
    End If
    txtLetterRef_LostFocus
End Sub

Private Sub cmdRender_Click()
    RenderLetter
    SSTab1.Tab = 1
End Sub

Private Sub cmdSelect_Click()
    sSQL(0) = "Select * from Letters ORDER BY 1 Desc"
    frmShow.Init sSQL(0)
    If Val(sArray(0)) <> 0 Then
        txtLetterRef.Text = sArray(0)
    End If
    txtLetterRef_LostFocus
    sSQL(0) = ""
End Sub

Private Sub Form_Load()
    Me.Move 0, 0
End Sub

Private Sub txtLetterRef_LostFocus()
    sSQL(0) = "Select * from Letters where LetterRef=" & Val(txtLetterRef.Text)
    Call dbOpen(0): Call ClearsArray(0): Call FillsArray(0): Call dbClose(0)
    If Val(sArray(0)) <> 0 Then
        txtLetterDate.Value = sArray(1)
        txtReceipient.Text = sArray(2)
        txtSubject.Text = sArray(3)
        txtSender.Text = sArray(4)
    End If
End Sub

Private Sub cmdSave_Click()
    If Val(txtLetterRef.Text) <> 0 Then
        sSQL(0) = "Update Letters Set LetterDate=" & Chr(39) & txtLetterDate.Value & Chr(39) & ", "
        sSQL(0) = sSQL(0) & " Receipient=" & Chr(39) & txtReceipient.Text & Chr(39) & ", "
        sSQL(0) = sSQL(0) & " Subject=" & Chr(39) & txtSubject.Text & Chr(39) & ", "
        sSQL(0) = sSQL(0) & " Sender=" & Chr(39) & txtSender.Text & Chr(39) & " Where LetterRef=" & Val(txtLetterRef.Text)
        dbOpen (0)
        dbClose (0)
    End If
    txtLetterRef.SetFocus
    RenderLetter
    SSTab1.Tab = 1
End Sub

Public Sub RenderLetter()
    vp.PaperSize = pprA4
    vp.MarginLeft = 700: vp.MarginRight = 500
    vp.MarginTop = 500: vp.MarginRight = 500
    
    vp.StartDoc
        RenderHead0
        RenderHead1
        RenderBody1
        RenderFooter
    vp.EndDoc
    Call RenderOverlay
End Sub

Public Sub RenderHead0()
    With vp
        .TextAlign = taCenterTop
        SetFont ("[frmPrintBill-Font00]")
        vp.DrawPicture mdiOne.ImgList.ListImages(1).Picture, 700, 700
        .CurrentY = 700
        .TextAlign = taLeftTop
        SetFont ("[frmPrintBill-Font01]")
        .StartTable
            .TableBorder = tbNone
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "0.5in"
            .TableCell(tcColWidth, , 2) = "3.5in"
            .TableCell(tcColAlign, , 2) = taLeftTop
            .TableCell(tcText, 1, 2) = CompanyName
        .EndTable
        
        SetFont ("[frmPrintBill-Font02]")
        .StartTable
            .TableBorder = tbNone
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "0.5in"
            .TableCell(tcColWidth, , 2) = "3.5in"
            .TableCell(tcColAlign, , 2) = taLeftTop
            .TableCell(tcText, 1, 2) = AboutCompany & vbCrLf & CompanyAddress & ", " & CompanyPhone & vbCrLf & CompanyFax & ", " & CompanyEmail & vbCrLf
        .EndTable
    End With
End Sub

Public Sub RenderHead1()
    With vp
        .IndentLeft = 150
        SetFont ("[frmPrintBill-Font03]")
        .StartTable
            .TableBorder = tbTop
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "3.5in": .TableCell(tcColWidth, , 2) = "3.5in"
            .TableCell(tcColAlign, , 1) = taLeftTop: .TableCell(tcColAlign, , 2) = taRightTop
            .TableCell(tcText, 1, 1) = "Letter No." & txtLetterRef.Text & vbCrLf
            .TableCell(tcText, 1, 2) = "Date: " & txtLetterDate.Value
        .EndTable
    
        .StartTable
            .TableBorder = tbNone
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "7.0in"
            .TableCell(tcColAlign, , 1) = taLeftTop
            .TableCell(tcText, 1, 1) = "To," & vbCrLf & vbTab & txtReceipient.Text & vbCrLf & vbCrLf & "Subject: " & txtSubject.Text
        .EndTable
        .CurrentY = .CurrentY + 100
    End With
End Sub

Public Sub RenderBody1()
    vp.Text = rtb.TextRTF
End Sub

Public Sub RenderFooter()
    With vp
        .CurrentY = .CurrentY + 150
        .StartTable
            .TableBorder = tbNone
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "7.0in": .TableCell(tcColAlign, , 1) = taRightTop
            .TableCell(tcText, 1, 1) = "Yours sincerely, " & vbCrLf & vbCrLf & vbCrLf & txtSender.Text
        .EndTable
    End With
End Sub
    
Public Sub RenderOverlay()
    With vp
        SetFont ("[frmPrintBill-Font02]")
        .TextAlign = taRightBottom
        For i = 1 To vp.PageCount
            vp.StartOverlay i: vp.CurrentY = 250
            vp.Text = "Page " & i & " of " & vp.PageCount
            vp.EndOverlay
        Next
    End With
End Sub

Private Sub SetFont(S As String)
    vp.FontName = ReadFont(S, 0)
    vp.FontSize = ReadFont(S, 1)
    vp.FontBold = ReadFont(S, 2)
    vp.FontItalic = ReadFont(S, 3)
End Sub

