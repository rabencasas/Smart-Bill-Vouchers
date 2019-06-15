Attribute VB_Name = "mod_Functions"
Public Function GenerateId(Yes As Boolean)
    If Yes Then
        Sheet1.Range("A51").Value = "[" & Format(DateTime.Now, "yy.mm.dd") & "][" & Format(DateTime.Now, "hh.mm.ss") & "]"
    Else
        Sheet1.Range("A51").Value = "ID not visible for preview"
    End If
End Function

Public Function OpenSheet()
    Sheet1.Unprotect ("admin.pass")
End Function

Public Function CloseSheet()
    Sheet1.Protect ("admin.pass")
End Function

Public Function SaveCertification()
    If Sheet5.Range("H7").Value <> "" Then
        If Sheet1.Range("A51").Value <> "ID not visible for preview" Then
            Dim file As String
            Dim textfile As Integer
            
            file = Sheet5.Range("H7").Value & "\" & Sheet1.Range("A51").Value & " - " & UCase(Sheet4.Range("H10").Value) & ".certificate"
            
            textfile = FreeFile
            
            Open file For Output As textfile
            
            Print #textfile, Sheet1.Range("A51").Value & " - " & UCase(Sheet4.Range("H10").Value) & vbNewLine
            
            Print #textfile, "Name:" & vbTab & vbTab & vbTab & UCase(Sheet4.Range("H10").Value)
            Print #textfile, "Age:" & vbTab & vbTab & vbTab & Sheet4.Range("H12").Value
            Print #textfile, "Address:" & vbTab & vbTab & UCase(Sheet4.Range("H14").Value)
            Print #textfile, "Assistance type:" & vbTab & UCase(Sheet4.Range("H16").Value)
            Print #textfile, "Amount:" & vbTab & vbTab & vbTab & Format(Sheet4.Range("H18").Value, "Standard")
            Print #textfile, "Beneficiary:" & vbTab & vbTab & Sheet4.Range("H20").Value
            Print #textfile, "Relationship to Beneficiary:" & vbTab & UCase(Sheet4.Range("H22").Value)
            Print #textfile, "Date Issued:" & vbTab & vbTab & Format(Sheet4.Range("H8").Value, "mmmm dd, yyyy") & "__" & WeekdayName(Weekday(Format(Sheet4.Range("H8").Value, "mmmm dd, yyyy")))
            Print #textfile, "Clerk:" & vbTab & vbTab & vbTab & Sheet4.Range("H5").Value
            
            Close textfile
            
            ' Show successfull log message to user
            Sheet4.Range("E24").Value = "Certifiation to " & UCase(Sheet4.Range("H10").Value) & " is successfully saved."
        Else
            MsgBox "Cannot continue to save. The certification has not been printed yet.", vbCritical
        End If
        
    Else
        MsgBox "Please set the log folder path in order to save."
    End If
End Function
