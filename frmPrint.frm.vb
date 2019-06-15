VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrint 
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9315
   OleObjectBlob   =   "frmPrint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcalculate_Click()
    OpenSheet

    CopyData
        
    If txtvoucher.Text <> "" Or txtvoucher.Text <> 0 Then
        Sheet3.PrintOut from:=1, to:=1, copies:=txtvoucher.Text
    End If
    
    If txtalobs.Text <> "" Or txtalobs.Text <> 0 Then
        Sheet4.PrintOut from:=1, to:=1, copies:=txtalobs.Text
    End If
    
    CloseSheet
    
    Hide
End Sub

Private Sub CommandButton1_Click()
    OpenSheet

    Sheet4.PrintPreview
    Sheet3.PrintPreview
    
    CloseSheet
End Sub
