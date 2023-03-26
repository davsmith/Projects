VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmButtons 
   Caption         =   "Macros"
   ClientHeight    =   3210
   ClientLeft      =   75
   ClientTop       =   255
   ClientWidth     =   5235
   OleObjectBlob   =   "Budget (Current Year) - Source_frmButtons.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnRefresh_Click()
    refreshTransactionData
End Sub

Private Sub btnPurge_Click()
    deleteTempSheets
End Sub

Private Sub btnClearAllFilter_Click()
    clearAllFilters
End Sub

Private Sub rbExpenses_Click()
    setSummaryFilters getViewParameters
End Sub

Private Sub rbIncome_Click()
    setSummaryFilters getViewParameters
End Sub

Private Sub rbIncomeAndExpenses_Click()
    setSummaryFilters getViewParameters
End Sub

Private Sub rbTaxes_Click()
    setSummaryFilters getViewParameters
End Sub


Private Sub chkAnnual_Click()
    Debug.Print "chkAnnual: " + CStr(chkAnnual.Value)
    setSummaryFilters getViewParameters
End Sub

Private Sub chkProjected_Click()
    setSummaryFilters getViewParameters
End Sub

Private Sub chkVacation_Click()
    setSummaryFilters getViewParameters
End Sub


Private Function getViewParameters()
    Dim strParameters As String
    
    strParameters = ""
    
    If (rbIncome.Value = True) Then
        strParameters = strParameters + "I"
    End If
    
    If (rbExpenses.Value = True) Then
        strParameters = strParameters + "E"
    End If
    
    If (rbIncomeAndExpenses.Value = True) Then
        strParameters = strParameters + "IE"
    End If
    
    If (rbTaxes.Value = True) Then
        strParameters = strParameters + "TIEAV"
    End If
    
    If (chkAnnual.Value = True) Then
        strParameters = strParameters + "A"
    End If
    
    If (chkProjected.Value = True) Then
        strParameters = strParameters + "P"
    End If
    
    If (chkVacation.Value = True) Then
        strParameters = strParameters + "V"
    End If
    
    getViewParameters = strParameters
End Function

Private Sub UserForm_Click()

End Sub
