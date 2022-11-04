VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3510
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    Dim CoteNominal As Double
    Dim CoteMax As Double
    Dim CoteMin As Double
    Dim Cote As Double
    
    Set MyDimension = CATIA.ActiveDocument.Selection.Item(1).Value
    MyDimensionValue = MyDimension.GetValue.Value
    
    CoteNominal = Round(MyDimensionValue, 3)
        
    MyDimension.GetTolerances oTolType, oTolName, oUpTol, oLowTol, odUpTol, odLowTol, oDisplayMode
    
    If oTolType = 1 Then
        CoteMin = CoteNominal + odLowTol
        CoteMax = CoteNominal + odUpTol
    End If
    If oTolType = 2 Then
        CoteMin = CoteNominal + oLowTol
        CoteMax = CoteNominal + oUpTol
    End If
    
    MsgBox ("Cote nominale: " & CoteNominal & Chr(13) & Chr(10) & "Cote max: " & CoteMax & Chr(13) & Chr(10) & "Cote min:" & CoteMin)
    
End Sub

Private Sub ListBox1_Click()

End Sub
