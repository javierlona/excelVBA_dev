VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Exhibit12 
   Caption         =   "Affiliates Worksheet ver. 8_3_2016"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9300
   OleObjectBlob   =   "Exhibit12.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Exhibit12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Activesheet.Unprotect ("txcdc1!")
    
    Dim MyValOCEPC As Integer
    Dim MyValGA As Integer
    Dim MyValNGA As Integer
    
    MyValOCEPC = Me.tbOCEPC.Value
    MyValGA = Me.tbGA.Value
    MyValNGA = Me.tbNGA.Value
    
    Dim string1 As String
    If MyValOCEPC <> 0 Then
        string1 = "NAME OF OC and/or EPC ENTITIES"
        Call addEntity(string1, MyValOCEPC)
    End If
    
    If MyValGA <> 0 Then
        string1 = "NAME OF NON-GUARANTOR AFFILIATES"
        Call addEntity(string1, MyValNGA)
    End If
    
    If MyValNGA <> 0 Then
        string1 = "NAME OF GUARANTOR AFFILIATES"
        Call addEntity(string1, MyValGA)
    End If
    
    Call countentities
    
    Unload Me
    Exhibit12.Hide
    Activesheet.Protect Password:="txcdc1!", AllowDeletingRows:=True, AllowInsertingRows:=True, AllowFormattingRows:=True

End Sub

Private Sub CommandButton2_Click()

    Activesheet.Unprotect ("txcdc1!")
    
    Call countentities
    
    Unload Me
    Exhibit12.Hide
    Activesheet.Protect Password:="txcdc1!", AllowDeletingRows:=True, AllowInsertingRows:=True, AllowFormattingRows:=True

End Sub

Private Sub CommandButton3_Click()

    Activesheet.Unprotect ("txcdc1!")
    
    Call calculateNetWorthNetProfit
    Call calculateSums
    
    Unload Me
    Exhibit12.Hide
    Activesheet.Protect Password:="txcdc1!", AllowDeletingRows:=True, AllowInsertingRows:=True, AllowFormattingRows:=True

End Sub
