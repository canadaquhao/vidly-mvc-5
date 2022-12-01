# vidly-mvc-5
A new line of code


```
Option Compare Database
Option Explicit

Private Sub calVisitDate_AfterUpdate()
On Error GoTo err_handler

    If fraConfirmationReport = 1 Then
        txtStartDate = calVisitDate
        fraConfirmationReport = 2
    Else
        txtEndDate = calVisitDate
    End If

err_exit:
    Exit Sub
    
err_handler:
    Call modError.errorHandler(Err.Number)
    Resume err_exit
End Sub

Private Sub cmdPrintReport_Click()
On Error GoTo err_handler
'Tony added at Dec.1st, 2021, get chamber status
If IsDate(txtStartDate) = True And IsDate(txtEndDate) = True Then
    If cboReportType = 0 Then
        Call generatechamberstatus(txtStartDate, txtEndDate)
    End If
End If

    Dim strWhere As String
    If IsDate(txtStartDate) = True And IsDate(txtEndDate) = False Then
        strWhere = "visitDate = #" & txtStartDate & "#"
    ElseIf IsDate(txtStartDate) = True And IsDate(txtEndDate) = True Then
    'Tony comment at July 10th, 2014, date format issue
        'strWhere = "visitDate BETWEEN '" & txtStartDate & "' AND '" & txtEndDate & "'"
         strWhere = "visitDate BETWEEN #" & txtStartDate & "# AND #" & txtEndDate & "#"
    Else
        MsgBox "Please select a valid date", vbInformation, "Invalid Visit Date"
    End If
    
    If cboReportType = 1 Then
        strWhere = strWhere & " AND bookingTypeID = 22"
    ElseIf cboReportType = 2 Then
        strWhere = strWhere & " AND bookingTypeID <> 22"
    End If
    
    If Len(strWhere) > 0 Then
        Dim reportname As String
        If cboReportType = 0 Then
            reportname = "rpt_gallery_visitor_listing_new"
        Else
            reportname = "rpt_gallery_visitor_listing_new2"
        End If
        DoCmd.OpenReport reportname, acViewPreview, , strWhere, acWindowNormal, cboReportType.Column(1)
        DoCmd.SelectObject acForm, "frmrptgalleryvisitorlisting"
        'DoCmd.Close
        DoCmd.SelectObject acReport, reportname
    End If
    
err_exit:
    Exit Sub
    
err_handler:
    Call modError.errorHandler(Err.Number)
    Resume err_exit
End Sub

Private Sub Form_Load()
On Error GoTo err_handler
    
    Call modProgram.chkLogin
    Call setFormDefaults
    
err_exit:
    Exit Sub
    
err_handler:
    Call modError.errorHandler(Err.Number)
    Resume err_exit
End Sub

Private Sub setFormDefaults()
On Error GoTo err_handler

    calVisitDate.Value = DateAdd("d", 1, Date)
    Call calVisitDate_AfterUpdate
    
err_exit:
    Exit Sub
    
err_handler:
    Call modError.errorHandler(Err.Number)
    Resume err_exit
End Sub

Private Sub fraConfirmationReport_AfterUpdate()
On Error GoTo err_handler

    If fraConfirmationReport.Value = 1 Then
        txtEndDate.Value = Null
    End If
    
err_exit:
    Exit Sub
    
err_handler:
    Call modError.errorHandler(Err.Number)
    Resume err_exit
End Sub

Private Sub generatechamberstatus(stdt As Date, enddt As Date)
       Dim dbConnection As New ADODB.Connection
       dbConnection.CursorLocation = adUseClient
       ''Dim rstChamberStatus As ADODB.Recordset
        ''Set dbConnection = CurrentProject.AccessConnection
       dbConnection.Open "LARS_VisitorServices"
       dbConnection.Execute ("exec [sp_rpt_chamber_status_generation] '" + Format(stdt, "yyyy-MM-dd") + "', '" + Format(enddt, "yyyy-MM-dd") + "'")
       dbConnection.Close
       Set dbConnection = Nothing

End Sub
```
