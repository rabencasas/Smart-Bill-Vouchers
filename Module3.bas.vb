Attribute VB_Name = "Module3"
Public Function CopyData()
    ''
    '' VOUCHER
    ''
    
    'due date
    Sheet3.Range("C26").Value = Sheet1.Range("B8").Value
    
    'type
    Sheet3.Range("B22").Value = Sheet1.Range("C8").Value
    'account name
    Sheet3.Range("C21").Value = Sheet1.Range("D8").Value
    'officer (alobs)
    Sheet4.Range("B25").Value = Sheet1.Range("D8").Value
    'account number
    Sheet3.Range("C23").Value = Sheet1.Range("E8").Value
    'srn/ mobile
    Sheet3.Range("C22").Value = Sheet1.Range("F8").Value
    
    ' payee
    Sheet3.Range("C12").Value = Sheet1.Range("C3").Value
    ' payee address
    Sheet3.Range("C14").Value = Sheet1.Range("C4").Value
    
    ' total charges
    Sheet3.Range("J19").Value = Format(Sheet1.Range("G14").Value, "Standard")
    ' vat
    Sheet3.Range("I23").Value = Sheet1.Range("H14").Value
    ' wt
    Sheet3.Range("I24").Value = Sheet1.Range("I14").Value
    ' total tax
    Sheet3.Range("J25").Value = Sheet1.Range("J14").Value
    'bal
    Sheet3.Range("J28").Value = Sheet1.Range("K14").Value
    'bill period
    Sheet3.Range("C24").Value = Format(Sheet1.Range("L8").Value, "MMMM dd, YYYY") & " - " & Format(Sheet1.Range("M8").Value, "MMMM dd, YYYY")
    
    ''
    '' ALOBS
    ''
    ' total charges
    Sheet4.Range("G15").Value = Format(Sheet1.Range("G14").Value, "Standard")
    
End Function
