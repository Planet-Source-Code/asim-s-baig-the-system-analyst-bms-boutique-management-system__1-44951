VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmprint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View & Print Report"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   1080
   ClientWidth     =   11880
   Icon            =   "frmprint.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11880
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4425
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox rpt1 
      Height          =   6510
      Left            =   15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   11483
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmprint.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00E2D1D3&
      Caption         =   "&Print Report"
      Height          =   435
      Left            =   -15
      MaskColor       =   &H00987758&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6510
      UseMaskColor    =   -1  'True
      Width           =   11865
   End
End
Attribute VB_Name = "frmprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdprint_Click()
On Error GoTo check_pr
    CD1.Flags = cdlPDDisablePrintToFile Or cdlPDNoSelection Or cdlPDReturnDC
    
    CD1.ShowPrinter
    For i = 1 To CD1.Copies
        rpt1.SelPrint CD1.hDC
    Next i
    
    Exit Sub
check_pr:
    If Err.Number = 32755 Then
          
    Else
        MsgBox "Error Occured: " & Err.Number & " " & Err.Description
    End If
End Sub

Private Sub Form_Load()
    Select Case GL_REPORT
        Case "Selling"
            Call show_selling
        Case "PL"
            Call show_PL
        Case "Customer"
            Call show_customer
        Case Else
    End Select
End Sub


Public Sub show_selling()
    
    On Error GoTo check_err1
    
    rpt1.FileName = App.Path & "/Reports/sellings.rtf"
    DoEvents
    DoEvents
    
    
    temp = rpt1.TextRTF
    temp = Replace(temp, "::FROM.DATE::", Format(reports.dtfrom.Value, "MMMM DD, YYYY"))
    temp = Replace(temp, "::TO.DATE::", Format(reports.dtto.Value, "MMMM DD, YYYY"))
    
    
    Call openconn

    sqlstr = "select S1.serialno, C1.name, C1.description, S1.customerid , " & _
            "C2.name , S1.qty , S1.amount , S1.dated " & _
            "from clothes C1, customer C2, selling S1 " & _
            "where (S1.dated between #" & reports.dtfrom.Value & "# and #" & reports.dtto.Value & "#) and (S1.serialno = C1.serialno) and (S1.customerid = C2.customerid) order by S1.dated asc"

    Call rs(sqlstr)
    
    grand = 0
    While Not (adoRS.EOF)
        
        temp = Replace(temp, "::Customer::", "Customer: " & "\tab " & adoRS.Fields(3) & Space(3) & adoRS.Fields(4), 1, 1)
        temp = Replace(temp, "::Cloth::", "Clothes: " & "\tab " & adoRS.Fields(0) & Space(3) & adoRS.Fields(1) & Space(3) & adoRS.Fields(2), 1, 1)
        temp = Replace(temp, "::DATE::", "Dated: " & "\tab " & "\tab " & Format(adoRS.Fields(7), "MMMM DD YYYY"), 1, 1)
        
        x = "\par " & "\par " & "::Customer::" & "\par " & "::Cloth::" & "\par " & "::DATE::" & "\par " & "::CALC::"
        
        temp = Replace(temp, "::CALC::", "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & adoRS.Fields(5) & "\tab " & "\tab " & numdisp(adoRS.Fields(6)) & "\tab " & "\tab " & numdisp(adoRS.Fields(5) * adoRS.Fields(6)) & x, 1, 1)
        
        grand = grand + (adoRS.Fields(5) * adoRS.Fields(6))
        adoRS.MoveNext
    Wend
        
    Call closeconn
    
        temp = Replace(temp, "::Customer::", "\del ")
        temp = Replace(temp, "\par ::Cloth::", "\del ")
        temp = Replace(temp, "\par ::DATE::", "\del ")
        temp = Replace(temp, "\par ::CALC::", "\del ")
    
    Call openconn
    
    sqlstr = "select I.customerid, C.name, I.amount, I.dated from income I , customer C where (I.dated between #" & reports.dtfrom.Value & "# and #" & reports.dtto.Value & "#) and (I.customerid = C.customerid) order by I.dated asc"
    
    Call rs(sqlstr)
    
            
    While Not (adoRS.EOF)
        temp = Replace(temp, "::Customer2::", "Customer: " & "\tab " & adoRS.Fields(0) & Space(3) & adoRS.Fields(1), 1, 1)
        temp = Replace(temp, "::DATE2::", "Dated: " & "\tab " & "\tab " & Format(adoRS.Fields(3), "MMMM DD YYYY"), 1, 1)
        
        x = "\par " & "\par " & "::Customer2::" & "\par " & "::DATE2::" & "\par " & "::CALC2::"
        
        temp = Replace(temp, "::CALC2::", "Amount: " & "\tab " & numdisp(adoRS.Fields(2)) & x, 1, 1)
        
        grand = grand + adoRS.Fields(2)
        adoRS.MoveNext
    Wend
    
    Call closeconn
        temp = Replace(temp, "::Customer2::", "\del ")
        temp = Replace(temp, "\par ::DATE2::", "\del ")
        temp = Replace(temp, "\par ::CALC2::", "\del ")

        temp = Replace(temp, "::GRAND::", numdisp(grand))
    
    
    
    rpt1.TextRTF = temp
   
    
    Exit Sub
check_err1:
    If Err.Number = 75 Then
        MsgBox "File: " & App.Path & "/Reports/sellings.rtf" & " is Open, Please Close It", vbCritical
        MsgBox "File: " & App.Path & "/Reports/sellings.rtf" & " is Open, Please Close It", vbCritical
        End
    Else
        MsgBox "Some Error Occured: " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Program May Not Work Properly", vbCritical
        MsgBox "Some Error Occured: " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Program May Not Work Properly", vbCritical
        End
    End If
End Sub

Public Sub show_PL()
    On Error GoTo check_err3
    
    rpt1.FileName = App.Path & "/Reports/PL.rtf"
    DoEvents
    DoEvents
    
    
    temp = rpt1.TextRTF
    temp = Replace(temp, "::FROM.DATE::", Format(reports.dtfrom.Value, "MMMM DD, YYYY"))
    temp = Replace(temp, "::TO.DATE::", Format(reports.dtto.Value, "MMMM DD, YYYY"))
    
    
    Call openconn

    sqlstr = "select S1.serialno, C1.name, C1.description, S1.customerid , " & _
            "C2.name , S1.qty , C1.costprice , S1.amount , S1.dated " & _
            "from clothes C1, customer C2, selling S1 " & _
            "where (S1.dated between #" & reports.dtfrom.Value & "# and #" & reports.dtto.Value & "#) and (S1.serialno = C1.serialno) and (S1.customerid = C2.customerid) order by S1.dated asc"

    Call rs(sqlstr)
    
    grand = 0
    While Not (adoRS.EOF)
        
        temp = Replace(temp, "::Customer::", "Customer: " & "\tab " & adoRS.Fields(3) & Space(3) & adoRS.Fields(4), 1, 1)
        temp = Replace(temp, "::Cloth::", "Clothes: " & "\tab " & adoRS.Fields(0) & Space(3) & adoRS.Fields(1) & Space(3) & adoRS.Fields(2), 1, 1)
        temp = Replace(temp, "::DATE::", "Dated: " & "\tab " & "\tab " & Format(adoRS.Fields(8), "MMMM DD YYYY"), 1, 1)
        
        x = "\par " & "\par " & "::Customer::" & "\par " & "::Cloth::" & "\par " & "::DATE::" & "\par " & "::CALC::"
        
        temp = Replace(temp, "::CALC::", "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & adoRS.Fields(5) & "\tab " & numdisp(adoRS.Fields(6)) & "\tab " & numdisp(adoRS.Fields(7)) & "\tab " & numdisp((adoRS.Fields(5) * adoRS.Fields(7)) - (adoRS.Fields(5) * adoRS.Fields(6))) & x, 1, 1)
        
        grand = grand + ((adoRS.Fields(5) * adoRS.Fields(7)) - (adoRS.Fields(5) * adoRS.Fields(6)))
        adoRS.MoveNext
    Wend
        
    Call closeconn
    
        temp = Replace(temp, "::Customer::", "\del ")
        temp = Replace(temp, "\par ::Cloth::", "\del ")
        temp = Replace(temp, "\par ::DATE::", "\del ")
        temp = Replace(temp, "\par ::CALC::", "\del ")
    
        If (grand > 0) Then
            temp = Replace(temp, "::PL::", "Profit")
            temp = Replace(temp, "::GRAND::", numdisp(grand))
        Else
            temp = Replace(temp, "::PL::", "Loss")
            temp = Replace(temp, "::GRAND::", numdisp(grand * -1))
        End If
            
    
    rpt1.TextRTF = temp
   
    
    Exit Sub
check_err3:
    If Err.Number = 75 Then
        MsgBox "File: " & App.Path & "/Reports/sellings.rtf" & " is Open, Please Close It", vbCritical
        MsgBox "File: " & App.Path & "/Reports/sellings.rtf" & " is Open, Please Close It", vbCritical
        End
    Else
        MsgBox "Some Error Occured: " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Program May Not Work Properly", vbCritical
        MsgBox "Some Error Occured: " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Program May Not Work Properly", vbCritical
        End
    End If
End Sub

Public Sub show_customer()
    On Error GoTo check_err2
    
    rpt1.FileName = App.Path & "/Reports/dues.rtf"
    DoEvents
    DoEvents
    
    
    temp = rpt1.TextRTF
    
    Call openconn

    If reports.Option1.Value Then
        sqlstr = "select * from customer"
    Else
        sqlstr = "select * from customer where dues > 0"
    End If

    Call rs(sqlstr)
    
    grand = 0
    While Not (adoRS.EOF)
        
        temp = Replace(temp, "::Account::", "Account: " & "\tab " & adoRS.Fields("customerid"), 1, 1)
        temp = Replace(temp, "::Customer::", "Customer: " & "\tab " & adoRS.Fields("name"), 1, 1)
        temp = Replace(temp, "::Phone::", "Phone: " & "\tab " & "\tab " & adoRS.Fields("phones"), 1, 1)
        
        x = "\par " & "\par " & "::Account::" & "\par " & "::Customer::" & "\par " & "::Phone::" & "\par " & "::Dues::"
        
        temp = Replace(temp, "::Dues::", "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & "\tab " & numdisp(adoRS.Fields("Dues")) & x, 1, 1)
        
        grand = grand + adoRS.Fields("dues")
        adoRS.MoveNext
    Wend
        
    Call closeconn
    
        temp = Replace(temp, "::Account::", "\del ")
        temp = Replace(temp, "\par ::Customer::", "\del ")
        temp = Replace(temp, "\par ::Phone::", "\del ")
        temp = Replace(temp, "\par ::Dues::", "\del ")
    
        temp = Replace(temp, "::GRAND::", numdisp(grand))
    
    rpt1.TextRTF = temp
   
    
    Exit Sub
check_err2:
    If Err.Number = 75 Then
        MsgBox "File: " & App.Path & "/Reports/sellings.rtf" & " is Open, Please Close It", vbCritical
        MsgBox "File: " & App.Path & "/Reports/sellings.rtf" & " is Open, Please Close It", vbCritical
        End
    Else
        MsgBox "Some Error Occured: " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Program May Not Work Properly", vbCritical
        MsgBox "Some Error Occured: " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Program May Not Work Properly", vbCritical
        End
    End If
End Sub

Private Sub rpt1_GotFocus()
    cmdprint.SetFocus
End Sub
