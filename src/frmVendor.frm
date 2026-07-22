VERSION 5.00
Begin VB.UserForm frmVendor
   Caption         =   "Vendor"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   StartUpPosition =   1
   Begin MSForms.TextBox txtVendorID
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
   Begin MSForms.TextBox txtVendorName
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   5160
   End
   Begin MSForms.TextBox txtOwnerName
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   5160
   End
   Begin MSForms.TextBox txtOwnerEmail
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   5160
   End
   Begin MSForms.TextBox txtContactPerson
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   2160
      Width           =   5160
   End
   Begin MSForms.TextBox txtContactEmail
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   2640
      Width           =   5160
   End
   Begin MSForms.TextBox txtContactPhone
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   5160
   End
   Begin MSForms.TextBox txtNotes
      Height          =   1500
      Left            =   1800
      MultiLine       =   -1
      ScrollBars      =   2
      TabIndex        =   7
      Top             =   3600
      Width           =   5160
   End
   Begin MSForms.CommandButton cmdSave
      Caption         =   "Save"
      Height          =   420
      Left            =   4800
      TabIndex        =   8
      Top             =   5640
      Width           =   1080
   End
   Begin MSForms.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   420
      Left            =   6000
      TabIndex        =   9
      Top             =   5640
      Width           =   1080
   End
   Begin MSForms.Label lblID
      Caption         =   "Vendor ID (edit only)"
      Height          =   240
      Left            =   240
      Top             =   300
      Width           =   1440
   End
   Begin MSForms.Label lblVendor
      Caption         =   "Vendor name *"
      Height          =   240
      Left            =   240
      Top             =   780
      Width           =   1440
   End
   Begin MSForms.Label lblOwner
      Caption         =   "Owner name *"
      Height          =   240
      Left            =   240
      Top             =   1260
      Width           =   1440
   End
   Begin MSForms.Label lblOwnerEmail
      Caption         =   "Owner email *"
      Height          =   240
      Left            =   240
      Top             =   1740
      Width           =   1440
   End
   Begin MSForms.Label lblContact
      Caption         =   "Contact person"
      Height          =   240
      Left            =   240
      Top             =   2220
      Width           =   1440
   End
   Begin MSForms.Label lblContactEmail
      Caption         =   "Contact email"
      Height          =   240
      Left            =   240
      Top             =   2700
      Width           =   1440
   End
   Begin MSForms.Label lblPhone
      Caption         =   "Contact phone"
      Height          =   240
      Left            =   240
      Top             =   3180
      Width           =   1440
   End
   Begin MSForms.Label lblNotes
      Caption         =   "Notes"
      Height          =   240
      Left            =   240
      Top             =   3660
      Width           =   1440
   End
End
Attribute VB_Name = "frmVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSubmitting As Boolean

Private Sub UserForm_Initialize()
    txtVendorID.Enabled = False
End Sub

Public Sub LoadVendor(ByVal vendorID As Long)
    Dim lo As ListObject
    Dim lr As ListRow
    Set lo = modUtil.GetTable(modUtil.TBL_VENDORS)
    Set lr = modUtil.FindLongRow(lo, "VendorID", vendorID)
    If lr Is Nothing Then Exit Sub
    txtVendorID.Value = vendorID
    txtVendorName.Value = modUtil.RowValue(lr, lo, "VendorName")
    txtOwnerName.Value = modUtil.RowValue(lr, lo, "BusinessOwnerName")
    txtOwnerEmail.Value = modUtil.RowValue(lr, lo, "BusinessOwnerEmail")
    txtContactPerson.Value = modUtil.RowValue(lr, lo, "VendorContactPerson")
    txtContactEmail.Value = modUtil.RowValue(lr, lo, "VendorContactEmail")
    txtContactPhone.Value = modUtil.RowValue(lr, lo, "VendorContactPhone")
    txtNotes.Value = modUtil.RowValue(lr, lo, "Notes")
End Sub

Private Sub cmdSave_Click()
    Dim ok As Boolean
    If mSubmitting Then Exit Sub
    mSubmitting = True: cmdSave.Enabled = False
    If Len(Trim$(txtVendorID.Value)) = 0 Then
        ok = (modVendors.CreateVendor(txtVendorName.Value, txtOwnerName.Value, txtOwnerEmail.Value, txtContactPerson.Value, txtContactEmail.Value, txtContactPhone.Value, txtNotes.Value) > 0)
    Else
        ok = modVendors.UpdateVendor(CLng(txtVendorID.Value), txtVendorName.Value, txtOwnerName.Value, txtOwnerEmail.Value, txtContactPerson.Value, txtContactEmail.Value, txtContactPhone.Value, txtNotes.Value)
    End If
    cmdSave.Enabled = True: mSubmitting = False
    If ok Then Unload Me
End Sub

Private Sub cmdCancel_Click(): Unload Me: End Sub
