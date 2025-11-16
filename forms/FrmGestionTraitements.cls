Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***     Copyright                                                                       ***
'******************************************************************************
'***    FORM:                                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'******************************************************************************
Option Explicit

'******************************************************************************
'***    Declaration D'Evenement Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   Set TxtTitTraitements = TxtTitValeur

   Set TxtFonTraitements = TxtFonValeur

   Set TxtObjTraitements = TxtObjValeur

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.Description, , "Form_Load"

   Resume Exit_Form_Load
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

   On Error GoTo Err_Form_Unload

   Set TxtTitTraitements = Nothing

   Set TxtFonTraitements = Nothing

   Set TxtObjTraitements = Nothing

Exit_Form_Unload:

   Exit Sub

Err_Form_Unload:

   MsgBox Err.Number & " " & Err.Description, , "Form_Unload"

   Resume Exit_Form_Unload
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_Timer()

   On Error GoTo Err_Form_Timer

   BteProgression.Left = BteProgressionFond.Left + 1

   BteProgression.Top = BteProgressionFond.Top + 1

   Select Case (PbrTraitements.Max = PbrTraitements.Min) Or (PbrTraitements.Value = PbrTraitements.Min)
      Case True

         BteProgression.Width = 0

      Case False

         Select Case ((PbrTraitements.Value - PbrTraitements.Min) / (PbrTraitements.Max - PbrTraitements.Min))
            Case Is <= 1

               BteProgression.Width = BteProgressionFond.Width * ((PbrTraitements.Value - PbrTraitements.Min) / (PbrTraitements.Max - PbrTraitements.Min)) - 2

            Case Else

         End Select

   End Select

   BteProgression.Height = BteProgressionFond.Height - 2

Exit_Form_Timer:

   Exit Sub

Err_Form_Timer:

   MsgBox Err.Number & " " & Err.Description, , "Form_Timer"

   Resume Exit_Form_Timer
End Sub
