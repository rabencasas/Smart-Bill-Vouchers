VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmaddpayment 
   ClientHeight    =   13395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12930
   OleObjectBlob   =   "frmaddpayment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmaddpayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboaccounts_Change()
    
    cbosrn.ListIndex = cboaccounts.ListIndex
    
    cbohead.ListIndex = cboaccounts.ListIndex
    
    cbotype.ListIndex = cboaccounts.ListIndex
    
    'With Sheets("INNOVE COMMUNICATIONS INC.").Range("B10:B500")
    'Set c = .Find(cboaccounts.Text, LookIn:=xlValues)
    'If Not c Is Nothing Then

        'Dim billno As Integer
        'Dim row As String
        'row = CStr(c.row)
        'billno = CInt(Sheets("INNOVE COMMUNICATIONS INC.").Range("E" + row).Value) + 1
        
        'Dim billfrom, billto, newbillfrom, newbillto As Date
        'billfrom = DateValue(Sheets("INNOVE COMMUNICATIONS INC.").Range("H" + row).Value)
        'billto = DateValue(Sheets("INNOVE COMMUNICATIONS INC.").Range("J" + row).Value)

        'newbillfrom = DateAdd("m", 1, billfrom)
        'newbillto = DateAdd("m", 1, billto)

        'txtperiodfrom.Text = newbillfrom
        'txtperiodto.Text = newbillto

        'txtbillno.Text = billno
    'End If
    'End With
    
End Sub

Private Sub cbosrn_Change()
    cbosrn.ListIndex = cboaccounts.ListIndex
End Sub

Private Sub cmdaddpayment_Click()
    
    OpenSheet
            
    'due date
    If txtdue.Text = "IMMEDIATELY" Then
        Sheet1.Range("B8").Value = txtdue.Text
    Else
        Sheet1.Range("B8").Value = Format(txtdue.Text, "MM/dd/yyyy")
    End If
            
    'type
    Sheet1.Range("C8").Value = cbotype.Text
    'officer
    Sheet1.Range("D8").Value = cbohead.Text
    'acc number
    Sheet1.Range("E8").Value = cboaccounts.Text
    'srn/ mobile
    Sheet1.Range("F8").Value = cbosrn.Text
    'charges
    Sheet1.Range("G8").Value = txtcharge.Text
    'vat
    Sheet1.Range("H8").Value = txtvat.Text
    'wt
    Sheet1.Range("I8").Value = txtwt.Text
    'total tax
    Sheet1.Range("J8").Value = txttotaltax.Text
    'balance
    Sheet1.Range("K8").Value = txtbal.Text
    'bill period from
    Sheet1.Range("L8").Value = txtperiodfrom.Text
    'bill period to
    Sheet1.Range("M8").Value = txtperiodto.Text
    
    CloseSheet
    
    Me.Hide
        
End Sub

Private Sub cmdcalculate_Click()
    ' deductions
    txtvat.Text = Format(txtcharge.Text / 1.12 * 0.05, "Standard")
    lblvat.Caption = txtcharge.Text / 1.12 * 0.05
    txtwt.Text = Format(txtcharge.Text / 1.12 * 0.02, "Standard")
    lblwt.Caption = txtcharge.Text / 1.12 * 0.02
    ' total tax
    txttotaltax.Text = CDec(txtvat.Text) + CDec(txtwt.Text)
    ' balance
    txtbal.Text = CDec(txtcharge.Text) - CDec(txttotaltax.Text)
End Sub

Private Sub txtcharge_Change()
    If txtcharge.Text <> "" Or txtcharge.Text <> 0 Then
        ' deductions
        txtvat.Text = Format(txtcharge.Text / 1.12 * 0.05, "Standard")
        lblvat.Caption = txtcharge.Text / 1.12 * 0.05
        txtwt.Text = Format(txtcharge.Text / 1.12 * 0.02, "Standard")
        lblwt.Caption = txtcharge.Text / 1.12 * 0.02
        ' total tax
        txttotaltax.Text = CDec(txtvat.Text) + CDec(txtwt.Text)
        ' balance
        txtbal.Text = CDec(txtcharge.Text) - CDec(txttotaltax.Text)
    End If
End Sub

Private Sub txtdue_Change()

End Sub

Private Sub txtperiodfrom_Change()

End Sub

Private Sub txtperiodfrom_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtperiodfrom.Text = Format(txtperiodfrom.Text, "dd mmm yyyy")
End Sub

Private Sub txtperiodfrom_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
End Sub

Private Sub txtperiodto_Change()

End Sub

Private Sub txtperiodto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtperiodto.Text = Format(txtperiodto.Text, "dd mmm yyyy")
End Sub

Private Sub UserForm_Activate()
        
End Sub

Private Sub UserForm_Click()

End Sub

Public Function AddToPayment()
End Function

Private Sub UserForm_Initialize()
    With cboaccounts
        .AddItem "0769674519"
        .AddItem "126639570"
        .AddItem "0729763977"
        .AddItem "718341310"
    End With
    
    With cbosrn
        .AddItem "9479177529"
        .AddItem "1010196857"
        .AddItem "9989859588"
        .AddItem "1011118367"
    End With
    
    With cbohead
        .AddItem "EMMANUEL L. IWAY"
        .AddItem "JOCELYN S. LIMKAICHONG"
        .AddItem "EMMANUEL L. IWAY"
        .AddItem "ADELA ARAULA"
    End With
    
    With cbotype
        .AddItem "SRN:"
        .AddItem "SRN:"
        .AddItem "Mobile Number:"
        .AddItem "SRN:"
    End With
End Sub
