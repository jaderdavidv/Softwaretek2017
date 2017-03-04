VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_nit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Comando108_Click()
Me.UpdateFIX = Me.UpdateFIX + 1
DoCmd.RunCommand acCmdSave
DoCmd.RunCommand acCmdRecordsGoToNew
End Sub
