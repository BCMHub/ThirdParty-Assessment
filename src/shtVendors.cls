VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtVendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    modVendors.RefreshVendorView
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim vendorID As Long
    If Target.Row < 5 Then Exit Sub
    vendorID = CLng(Val(CStr(Cells(Target.Row, 1).Value2)))
    If vendorID > 0 And modIdentity.RequireMutation() Then
        Cancel = True
        frmVendor.LoadVendor vendorID
        frmVendor.Show
    End If
End Sub
