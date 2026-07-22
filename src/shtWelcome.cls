VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Range("B8").Value = "Windows identity: " & IIf(Len(modIdentity.CurrentUser()) > 0, modIdentity.CurrentUser(), "unresolved")
    Range("B9").Value = "Access: " & IIf(modIdentity.SessionCanMutate(), modIdentity.CurrentRole(), "Browse-only")
    Range("B10").Value = modIdentity.MutationBlockReason()
End Sub
