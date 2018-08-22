Attribute VB_Name = "Module1"
Public i As Integer, x As Single, y As Single, MémoriseY As Single
Public j As Integer, Cont As Integer, k As Integer, m As Integer, n As Integer
Public l As Integer, Motion As Integer, FlagMouvement As Integer
Public NumSoc As Integer, NumUser As Integer
Public TypeSoc As Integer, InitialeUser As String
Public NomSoc As String, Identificateur As String
Public LieuSoc As String, P1 As Integer
Public Chemin As String, License As String
Public NomUser As String, PrenomUser As String, PassWord As String
Public Index As Integer, ChoixFiche As Integer, ChoixFiche1 As Integer
Public Reponse As String, Réponse As Integer
Public Flag As Integer, Tex As String
Public Disque As String, PassWordOk As Integer
Public ChoixPassWord As Integer '1= lancement, 2=modification mot de passe
Public Exercice As Integer, Débiteur As String
Public TypeVéhicule As Integer
Public TauxTva As Double, VariableAide As Integer
Public DateLimiteSup As Date, DateLimiteinf As Date
Public MO As Double, NuméroDébiteur As String, Visualiser As Integer
Public Pas As Single, Création As Integer, NumFiche As Integer
Public AncienKM As Integer, DernierDébiteurCréer As Long
Public KilomètreA As Integer, CheminV As String, energie As String





Public Function ControleAscii(ascii As Integer, NumType As Integer) As Integer

Select Case NumType
Case 1 'Uniquement des chiffres
 If ascii = 8 Then ControleAscii = ascii: Exit Function
 If ascii < 48 Or ascii > 57 Then
  If ascii <> 13 Then ControleAscii = 0
 Else
  ControleAscii = ascii
 End If

Case 2 'Uniquement des chiffres + le point
 If ascii = 8 Or ascii = 46 Then ControleAscii = ascii: Exit Function
 If ascii < 48 Or ascii > 57 Then
  If ascii <> 13 Then
  ControleAscii = 0
  Else
  ControleAscii = ascii
  End If
 Else
  ControleAscii = ascii
 End If


Case 3 'Uniquement des chiffres , le point , le + le -
 If ascii = 8 Or ascii = 46 Then ControleAscii = ascii: Exit Function
 If ascii = 43 Or ascii = 45 Then
  ControleAscii = ascii: Exit Function
 End If
 
 If ascii < 48 Or ascii > 57 Then
  If ascii <> 13 Then ControleAscii = 0
 Else
  ControleAscii = ascii
 End If




Case 4 'Force en majuscule
 VarStringA = Chr(ascii)
 ControleAscii = Asc(UCase(VarStringA))

Case 5 'Force en minuscule
 VarStringA = Chr(ascii)
 ControleAscii = Asc(LCase(VarStringA))


Case 6 'Uniquement des chiffres , et -, et /
 If ascii = 8 Or ascii = 45 Or ascii = 47 Then ControleAscii = ascii: Exit Function
 If ascii < 48 Or ascii > 57 Then
  If ascii <> 13 Then ControleAscii = 0
 Else
  ControleAscii = ascii
 End If

Case 7 'Oui ou Non
 VarStringA = Chr(ascii)
 VarStringA = Asc(UCase(VarStringA))
 If VarStringA = 78 Or VarStringA = 79 Then
  ControleAscii = VarStringA
 Else
   If ascii <> 13 Then ControleAscii = 0
 End If

Case 8 'Uniquement - et +
 If ascii = 8 Or ascii = 45 Or ascii = 43 Then
  ControleAscii = ascii
 Else
 If ascii <> 13 Then ControleAscii = 0
 End If

Case 10 'Aucune entrée clavier n'est tolérée
 ControleAscii = 0

Case Else

End Select


End Function
Public Function Arrondi(Montant As Double)
 Select Case FormatMonnaie
 Case "0.00"
 Arrondi = Int((Montant * 100) + 0.5) / 100
 Case "0.000"
 Arrondi = Int((Montant * 1000) + 0.5) / 1000
 Case "0.0000"
 Arrondi = Int((Montant * 10000) + 0.5) / 10000
 Case "0.00000"
 Arrondi = Int((Montant * 100000) + 0.5) / 100000
End Select




End Function

Public Function ControleDate(DateaControler As String) As Boolean
If Val(DateaControler) > 0 Then
 If IsDate(DateaControler) = False Then
    ControleDate = False
    Else
    ControleDate = True
 End If
 Else
 ControleDate = True
End If
End Function

Public Sub Surbrillance()
  '  SendKeys "{Home}+{End}"
End Sub

