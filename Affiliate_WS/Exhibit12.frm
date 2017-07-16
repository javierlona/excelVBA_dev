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
Public MyValOCEPC As Integer
Public MyValGA As Integer
Public MyValNGA As Integer

Private Sub CommandButton1_Click()

    Activesheet.Unprotect ("txcdc1!")
    MyValOCEPC = Me.tbOCEPC.Value
    MyValGA = Me.tbGA.Value
    MyValNGA = Me.tbNGA.Value
    
    Dim int1 As Integer
    Dim string1 As String
    
    string1 = "NAME OF OC and/or EPC ENTITIES"
    int1 = MyValOCEPC
    Call showNS(string1, int1)
    
    string1 = "NAME OF NON-GUARANTOR AFFILIATES"
    int1 = MyValNGA
    Call showNS(string1, int1)
    
    string1 = "NAME OF GUARANTOR AFFILIATES"
    int1 = MyValGA
    Call showNS(string1, int1)
    
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
