Attribute VB_Name = "modSRCA"
Option Explicit

Public Characters(1 To 1000) As Character
Public gCurrent As Integer

Type Character
    stName As String
    
    iBody As Integer
    iQuickness As Integer
    iStrength As Integer
    iCharisma As Integer
    iIntelligence As Integer
    iWillpower As Integer
    iHealth As Integer
    iInitiative As Integer    'Just initiative dice
    iReaction As Integer
    iMagic As Integer
    iEssence As Double
    
    iDamage As Integer
    iDamageLevel As String    'L, M, S, D
    iBArmour As String
    iIArmour As String
    
    bInUse As Boolean
    
    stInventory As String
    stSkills As String ' Firearms 5
                       ' Athletics 3
                       ' etc.
    iImageNum As Integer
    
End Type




Public Function LoadArchtypeFromFile(Name As String) As Character
    'All archtypes must be saved in a file called "[name].dat" in the same folder
    'as the exe
    Dim temp As Character
    'DEBUG
    Open App.Path & "\archtypes\" & Name & ".dat" For Input As #1
    Input #1, temp.stName
    Input #1, temp.iBody
    Input #1, temp.iQuickness
    Input #1, temp.iStrength
    Input #1, temp.iCharisma
    Input #1, temp.iIntelligence
    Input #1, temp.iWillpower
    
    Input #1, temp.iEssence
    Input #1, temp.iMagic
    Input #1, temp.iReaction
    Input #1, temp.iInitiative
    
    Input #1, temp.iDamage
    Input #1, temp.iDamageLevel
    Input #1, temp.iBArmour
    Input #1, temp.iIArmour
    
    Input #1, temp.stInventory
    temp.stInventory = Replace(temp.stInventory, "\n", vbCrLf)
    
    Input #1, temp.stSkills
    temp.stSkills = Replace(temp.stSkills, "\n", vbCrLf)
    
    Input #1, temp.iImageNum
    
    LoadArchtypeFromFile = temp
    
    Close #1
End Function

Public Sub SaveArchtypeToFile(guy As Character)
    'DEBUG
    'First, add it to the index file:
    Open App.Path & "\archtypes\index.dat" For Append As #1
    Print #1, guy.stName
    Close #1
    
    'Now add the info to the file
    Open App.Path & "\archtypes\" & guy.stName & ".dat" For Output As #1
    Print #1, guy.stName
    Print #1, guy.iBody
    Print #1, guy.iQuickness
    Print #1, guy.iStrength
    Print #1, guy.iCharisma
    Print #1, guy.iIntelligence
    Print #1, guy.iWillpower
    
    Print #1, guy.iEssence
    Print #1, guy.iMagic
    Print #1, guy.iReaction
    Print #1, guy.iInitiative
    
    Print #1, guy.iDamage
    Print #1, guy.iDamageLevel
    Print #1, guy.iBArmour
    Print #1, guy.iIArmour
    guy.stInventory = Replace(guy.stInventory, vbCrLf, "\n")
    Print #1, guy.stInventory
    
    guy.stSkills = Replace(guy.stSkills, vbCrLf, "\n")
    Print #1, guy.stSkills
    
    Print #1, guy.iImageNum
    
    Close #1
    
    
    
    
End Sub



Public Function MakeCharacterVariable(stName, _
                                      stBody As String, _
                                      stQuickness As String, _
                                      stStrength As String, _
                                      stCharisma As String, _
                                      stIntelligence As String, _
                                      stWillpower As String, _
                                      stEssence As String, _
                                      stMagic As String, _
                                      stReaction As String, _
                                      stInitiative As String, _
                                      stDamage As String, _
                                      stBArmour As String, _
                                      stIArmour As String, _
                                      stInventory As String, _
                                      stSkills As String, _
                                      iImageNum As Integer) As Character
                                                             
    Dim temp As Character
                                                             
    temp.bInUse = True
                                                             
    temp.stName = stName
    temp.iBody = Val(stBody)
    temp.iQuickness = Val(stQuickness)
    temp.iStrength = Val(stStrength)
    temp.iCharisma = Val(stCharisma)
    temp.iIntelligence = Val(stIntelligence)
    temp.iWillpower = Val(stWillpower)
    
    temp.iReaction = Val(stReaction)
    temp.iInitiative = Val(stInitiative)
    temp.iEssence = Val(stEssence)
    temp.iMagic = Val(stMagic)
  
    temp.iDamageLevel = Right(stDamage, 1)
    temp.iDamage = Left(stDamage, Len(stDamage) - 1)
    temp.iBArmour = Val(stBArmour)
    temp.iIArmour = Val(stIArmour)
    
    temp.stInventory = stInventory
    temp.stSkills = stSkills
    
    temp.iImageNum = iImageNum
    
    
    MakeCharacterVariable = temp
                                                             

End Function


Public Sub CreateCharacter(iNum As Integer, TheCharacter As Character)
    Load frmMain.imgActive(iNum)
    Load frmMain.lblHealth(iNum)
    Load frmMain.lblName(iNum)
    
    
    With frmMain.lblName(iNum)
        .Caption = TheCharacter.stName
        .Left = 0
        .Top = 0
        .BorderStyle = 1
        .Visible = True
    End With
    
    With frmMain.lblHealth(iNum)
        .Left = 0
        .Top = 0 + frmMain.lblName(iNum).Height
        .BorderStyle = 1
        .Visible = True
    End With
    
    With frmMain.imgActive(iNum)
        .BorderStyle = 1
        .Picture = frmMain.imgCharacter(TheCharacter.iImageNum).Picture
        .Visible = True
        .Left = 0
        .Top = 0 + frmMain.lblName(iNum).Height + frmMain.lblHealth(iNum).Height
        .Stretch = True
        .Height = .Height / (.Width / 1515)
        .Width = 1515
        .DragMode = 1
    End With

End Sub




Public Sub ShowStats(guy As Character)
    Load frmStats
    
    frmStats.txtName.Text = guy.stName
    frmStats.Caption = "Stats of " & guy.stName
    
    frmStats.txtBody.Text = guy.iBody
    frmStats.txtQuickness.Text = guy.iQuickness
    frmStats.txtStrength.Text = guy.iStrength
    frmStats.txtCharisma.Text = guy.iCharisma
    frmStats.txtIntelligence.Text = guy.iIntelligence
    frmStats.txtWillpower.Text = guy.iWillpower

    frmStats.txtReaction.Text = guy.iReaction
    frmStats.txtInitiative.Text = guy.iInitiative
    frmStats.txtMagic.Text = guy.iMagic
    frmStats.txtEssence.Text = guy.iEssence

    frmStats.txtSkills.Text = guy.stSkills
    frmStats.txtInventory.Text = guy.stInventory
    
    frmStats.txtDamage.Text = guy.iDamage & guy.iDamageLevel
    frmStats.txtbArmour.Text = guy.iBArmour
    frmStats.txtiArmour.Text = guy.iIArmour
    
    frmStats.imgPreview.Picture = frmMain.imgCharacter(guy.iImageNum).Picture
    
    
    frmStats.Show vbModeless

End Sub

Public Sub EditNPC(GuyNum As Integer)
    Load frmBlankNPC
    Dim iIndex As Integer
    Dim guy As Character
    guy = Characters(GuyNum)
    
    frmBlankNPC.txtName.Text = guy.stName
    frmBlankNPC.Caption = "Stats of " & guy.stName
    
    frmBlankNPC.txtBody.Text = guy.iBody
    frmBlankNPC.txtQuickness.Text = guy.iQuickness
    frmBlankNPC.txtStrength.Text = guy.iStrength
    frmBlankNPC.txtCharisma.Text = guy.iCharisma
    frmBlankNPC.txtIntelligence.Text = guy.iIntelligence
    frmBlankNPC.txtWillpower.Text = guy.iWillpower

    frmBlankNPC.txtReaction.Text = guy.iReaction
    frmBlankNPC.txtInitiative.Text = guy.iInitiative
    frmBlankNPC.txtMagic.Text = guy.iMagic
    frmBlankNPC.txtEssence.Text = guy.iEssence

    frmBlankNPC.txtSkills.Text = guy.stSkills
    frmBlankNPC.txtInventory.Text = guy.stInventory
    
    frmBlankNPC.txtDamage.Text = guy.iDamage & guy.iDamageLevel
    frmBlankNPC.txtbArmour.Text = guy.iBArmour
    frmBlankNPC.txtiArmour.Text = guy.iIArmour
    
    frmBlankNPC.imgPreview.Picture = frmMain.imgCharacter(guy.iImageNum).Picture
        
    frmBlankNPC.txtEditGuy.Text = GuyNum
    
    For iIndex = 0 To frmBlankNPC.lstPicture.ListCount
        If frmBlankNPC.lstPicture.ItemData(iIndex) = guy.iImageNum Then
            frmBlankNPC.lstPicture.ListIndex = iIndex
            Exit For
        End If
    Next
        
    frmBlankNPC.Show vbModeless

End Sub


Public Function RollDice(Dice As Integer, Target As Integer)
    Dim iIndex As Integer
    Dim Rolled As Integer
    Dim total As Integer
    Dim Successes As Integer
    
    Randomize Timer
    
    Successes = 0
    For iIndex = 1 To Dice
        Rolled = 6
        total = 0
        
        While Rolled = 6
            Rolled = Int(Rnd * 6) + 1
            total = total + Rolled
        Wend
        
        If total >= Target Then
            Successes = Successes + 1
        End If
    Next
    
    RollDice = Successes
            
End Function

Public Function Roll(number As Integer, sides As Integer)
    Dim iIndex As Integer
    Dim total As Integer
    
    Randomize Timer
    For iIndex = 1 To number
        total = total + Int(Rnd * sides) + 1
    Next
    Roll = total
End Function


Public Sub SelAll(source As TextBox)
    source.SelStart = 0
    source.SelLength = Len(source.Text)
End Sub

