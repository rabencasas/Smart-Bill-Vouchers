VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    frmPrint.Show
End Sub

Private Sub CommandButton2_Click()
    frmaddpayment.Show
End Sub

Private Sub CommandButton3_Click()
    OpenSheet

    'save to Bill Record sheets
        
        ' insert spacing (2 rows)
        Sheet5.Range("A5").EntireRow.Insert
        
        Sheet5.Range("A5").EntireRow.Insert
        
        'type:a
        
        'srn/mobile no:b
        
        'account no:c
        
        'acc name:d
        
        'bill period from:e
        
        'to:f
        
        'amount paid:i
        
        'due date:j
        
        
        Sheet5.Range("B15").Value = Sheets("FORM").Range("C" & i).Value 'ACCOUNT NUMBER
        Sheet5.Range("E15").Value = Sheets("FORM").Range("B" & i).Value 'BILL NUMBER
        Sheet5.Range("H15").Value = Sheets("FORM").Range("I" & i).Value 'BILL PERIOD FROM
        Sheet5.Range("J15").Value = Sheets("FORM").Range("J" & i).Value 'BILL PERIOD TO
        'Sheet5.Range("L15").Value = Sheets("FORM").Range("H" & i).Value 'PREVIOUS BALANCE
        'Sheet5.Range("O15").Value = Sheets("FORM").Range("H" & i).Value 'TOTAL CURRENT BILL
        Sheet5.Range("R15").Value = Sheets("FORM").Range("D" & i).Value 'AMOUNT PAID
        Sheet5.Range("U15").Value = DateTime.Now 'DATE PRINTED
    
        MsgBox "Voucher successfully logged!", vbOKOnly
        
        CloseSheet
End Sub

Private Sub CommandButton4_Click()
    OpenSheet

    CopyData

    Sheet4.PrintPreview
    Sheet3.PrintPreview
    
    CloseSheet
End Sub
