
Private Sub cmdDelete_Click()
    Dim X As Long
        Dim Y As Long
        X = Sheets("WIP").Range("A" & Rows.Count).End(xlUp).Row
        
        If MsgBox("Are you sure you want to DELETE this project? ", vbYesNo + vbQuestion, "Question") = vbNo Then
        Exit Sub
        End If
        
        For Y = 7 To X
        If Sheets("WIP").Cells(Y, 2).Value = cmbSearch.Text Then
        Rows(Y).Delete
        
        End If
        Next Y
        
         '''''''''''Clear & Reset Boxes''''''''''''
        Me.cmbSearch.Value = ""
        Me.cmbStatus.Value = ""
        Me.txtProjectName.Value = ""
        Me.cmbEstimator.Value = ""
        Me.txtAddress.Value = ""
        Me.txtCity.Value = ""
        Me.cmbState.Value = ""
        Me.txtZip.Value = ""
        Me.cmbVACounty.Value = " "
        Me.cmbBonded.Value = ""
        Me.cmbWageScale.Value = ""
        Me.txtOriginalContract.Value = ""
        Me.txtEstimatedGross.Value = ""
        Me.txtBillings.Value = ""
        Me.txtIncurred.Value = ""
        Me.txtCostToComplete.Value = ""
        
    
        MsgBox "Project has been deleted.", vbInformation
        cmbStatus.SetFocus

End Sub

Private Sub cmdExit_Click()
        Unload Me
End Sub

Private Sub cmdReset_Click()
        Unload Me
        UserForm1.Show
End Sub

Private Sub cmdSave_Click()
        Dim sh As Worksheet
        Set sh = ThisWorkbook.Sheets("WIP")
        Dim le As Long
        LR = sh.Range("A" & Rows.Count).End(xlUp).Row
         
        ''''''''''Validation'''''''''''''
        If Me.cmbStatus <> "Open" And Me.cmbStatus <> "Completed" Then
        MsgBox "Select Project status from drop down.", vbCritical
        Exit Sub
        End If
        
        If Me.cmbEstimator = "" Then
        MsgBox "Please select an Estimator.", vbCritical
        Exit Sub
        End If
        
       If Me.cmbWageScale <> "No" And Me.cmbWageScale <> "Yes" Then
        MsgBox "Please indicate Wage Scale status.", vbCritical
        Exit Sub
        End If
       
       If Me.cmbState <> "MD" And Me.cmbState <> "VA" And Me.cmbState <> "DC" Then
        MsgBox "Please select a state.", vbCritical
        Exit Sub
        End If
        
        If Application.WorksheetFunction.CountIf(sh.Range("B:B"), Me.txtProjectName.Text) > 0 Then
            MsgBox "Name already exists! Select UPDATE option for current projects.", vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        
       '''''''''''Add data om Excel Sheet''''''''''
       If MsgBox("Are you sure you want to save a NEW project? ", vbYesNo + vbQuestion, "Question") = vbNo Then
        Exit Sub
        End If
        
        With sh
            .Cells(LR + 1, 1).Value = Me.cmbStatus.Value
            .Cells(LR + 1, "B").Value = Me.txtProjectName.Value
            .Cells(LR + 1, "C").Value = Me.cmbEstimator.Value
            .Cells(LR + 1, "D").Value = Me.txtAddress.Value
            .Cells(LR + 1, "E").Value = Me.txtCity.Value
            .Cells(LR + 1, "F").Value = Me.cmbState.Value
            .Cells(LR + 1, "G").Value = Me.txtZip.Value
            .Cells(LR + 1, "H").Value = Me.cmbVACounty.Value
            .Cells(LR + 1, "I").Value = Me.cmbBonded.Value
            .Cells(LR + 1, "J").Value = Me.cmbWageScale.Value
            .Cells(LR + 1, "K").Value = Me.txtOriginalContract.Value
            .Cells(LR + 1, "L").Value = Me.txtEstimatedGross.Value
            .Cells(LR + 1, "M").Value = Me.txtBillings.Value
            .Cells(LR + 1, "M").Value = Me.txtIncurred.Value
            .Cells(LR + 1, "O").Value = Me.txtCostToComplete.Value
            .Cells(LR + 1, "P").Value = Application.UserName & "-" & Format(Now(), "MM/DD/YYYY, HH:MM AM/PM")

        End With
        
         '''''''''''Clear & Reset Boxes''''''''''''
            Me.cmbStatus.Value = ""
            Me.txtProjectName.Value = ""
            Me.cmbEstimator.Value = ""
            Me.txtAddress.Value = ""
            Me.txtCity.Value = ""
            Me.cmbState.Value = ""
            Me.txtZip.Value = ""
            Me.cmbVACounty.Value = " "
            Me.cmbBonded.Value = ""
            Me.cmbWageScale.Value = ""
            Me.txtOriginalContract.Value = ""
            Me.txtEstimatedGross.Value = ""
            Me.txtBillings.Value = ""
            Me.txtIncurred.Value = ""
            Me.txtCostToComplete.Value = ""
        
            Call Refresh_data
            
            MsgBox "Your NEW project has been successfully added.", vbInformation
            cmbStatus.SetFocus
            
         
            LR = Range("B" & Rows.Count).End(xlUp).Row
            Application.EnableEvents = False
            Range("B7:BB" & LR).Sort Key1:=Range("B7"), Order1:=xlAscending, _
            Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            Application.EnableEvents = True
            

        
        End Sub



Sub Refresh_data()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("WIP")
    Dim le As Long
    LR = Sheets("WIP").Range("A" & Rows.Count).End(xlUp).Row
            
    If LR = 6 Then LR = 7
    With Me.ListBox1
            .ColumnCount = 15
            .ColumnHeads = True
            .ColumnWidths = "35,140,40,90,65,20,35,95,30,25,50,50,50,50,50"
            .RowSource = "WIP!A7:O" & LR
    End With
    
End Sub


Private Sub cmdSearch_Click()
    Dim X As Long
    Dim Y As Long
    X = Sheets("WIP").Range("A" & Rows.Count).End(xlUp).Row
    
    For Y = 7 To X
    If Sheets("WIP").Cells(Y, 2).Value = cmbSearch.Text Then
    cmbStatus = Sheets("WIP").Cells(Y, 1).Value
    txtProjectName = Sheets("WIP").Cells(Y, 2).Value
    cmbEstimator = Sheets("WIP").Cells(Y, 3).Value
    txtAddress = Sheets("WIP").Cells(Y, 4).Value
    txtCity = Sheets("WIP").Cells(Y, 5).Value
    cmbState = Sheets("WIP").Cells(Y, 6).Value
    txtZip = Sheets("WIP").Cells(Y, 7).Value
    cmbVACounty = Sheets("WIP").Cells(Y, 8).Value
    cmbBonded = Sheets("WIP").Cells(Y, 9).Value
    cmbWageScale = Sheets("WIP").Cells(Y, 10).Value
    txtOriginalContract = Sheets("WIP").Cells(Y, 11).Value
    txtEstimatedGross = Sheets("WIP").Cells(Y, 12).Value
    txtBillings = Sheets("WIP").Cells(Y, 13).Value
    txtIncurred = Sheets("WIP").Cells(Y, 14).Value
    txtCostToComplete = Sheets("WIP").Cells(Y, 15).Value
    
    
    End If
    Next Y

End Sub

Private Sub cmdUpdate_Click()
        Dim X As Long
        Dim Y As Long
        X = Sheets("WIP").Range("A" & Rows.Count).End(xlUp).Row
    
        
        ''''''''''Validation'''''''''''''
        If Me.cmbStatus <> "Open" And Me.cmbStatus <> "Completed" Then
        MsgBox "Select Project status from drop down.", vbCritical
        Exit Sub
        End If
        
        If Me.cmbEstimator = "" Then
        MsgBox "Please select an Estimator.", vbCritical
        Exit Sub
        End If
        
        If Me.cmbWageScale <> "No" And Me.cmbWageScale <> "Yes" Then
        MsgBox "Please indicate Wage Scale status.", vbCritical
        Exit Sub
        End If
        
        If Me.cmbState <> "MD" And Me.cmbState <> "VA" And Me.cmbState <> "DC" Then
        MsgBox "Please select a state.", vbCritical
        Exit Sub
        End If
        
        
        '''''''''''''Add Update Data to excel'''''''''''
        
        If MsgBox("Are you sure you want to UPDATE an existing project? ", vbYesNo + vbQuestion, "Question") = vbNo Then
        Exit Sub
        End If
        
    
        For Y = 7 To X
        If Sheets("WIP").Cells(Y, 2).Value = cmbSearch.Text Then
        Sheets("WIP").Cells(Y, 2).Value = txtProjectName
        Sheets("WIP").Cells(Y, 1).Value = cmbStatus
        Sheets("WIP").Cells(Y, 3).Value = cmbEstimator
        Sheets("WIP").Cells(Y, 4).Value = txtAddress
        Sheets("WIP").Cells(Y, 5).Value = txtCity
        Sheets("WIP").Cells(Y, 6).Value = cmbState
        Sheets("WIP").Cells(Y, 7).Value = txtZip
        Sheets("WIP").Cells(Y, 8).Value = cmbVACounty
        Sheets("WIP").Cells(Y, 9).Value = cmbBonded
        Sheets("WIP").Cells(Y, 10).Value = cmbWageScale
        Sheets("WIP").Cells(Y, 11).Value = txtOriginalContract
        Sheets("WIP").Cells(Y, 12).Value = txtEstimatedGross
        Sheets("WIP").Cells(Y, 13).Value = txtBillings
        Sheets("WIP").Cells(Y, 14).Value = txtIncurred
        Sheets("WIP").Cells(Y, 15).Value = txtCostToComplete
    
        End If
        Next Y
        
        '''''''''''Clear & Reset Boxes''''''''''''
        Me.cmbSearch.Value = ""
        Me.cmbStatus.Value = ""
        Me.txtProjectName.Value = ""
        Me.cmbEstimator.Value = ""
        Me.txtAddress.Value = ""
        Me.txtCity.Value = ""
        Me.cmbState.Value = ""
        Me.txtZip.Value = ""
        Me.cmbVACounty.Value = " "
        Me.cmbBonded.Value = ""
        Me.cmbWageScale.Value = ""
        Me.txtOriginalContract.Value = ""
        Me.txtEstimatedGross.Value = ""
        Me.txtBillings.Value = ""
        Me.txtIncurred.Value = ""
        Me.txtCostToComplete.Value = ""
        
        
        MsgBox "Project has been successfully updated.", vbInformation
        cmbStatus.SetFocus
        
         LR = Range("B" & Rows.Count).End(xlUp).Row
            Application.EnableEvents = False
            Range("B7:BB" & LR).Sort Key1:=Range("B7"), Order1:=xlAscending, _
            Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            Application.EnableEvents = True

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        cmbSearch = ListBox1.Column(1)
        If cmbSearch.Text = ListBox1.Column(1) Then
        cmbStatus.Text = ListBox1.Column(0)
        txtProjectName.Text = ListBox1.Column(1)
        cmbEstimator.Text = ListBox1.Column(2)
        txtAddress.Text = ListBox1.Column(3)
        txtCity.Text = ListBox1.Column(4)
        cmbState.Text = ListBox1.Column(5)
        txtZip.Text = ListBox1.Column(6)
        cmbVACounty.Text = ListBox1.Column(7)
        cmbBonded.Text = ListBox1.Column(8)
        cmbWageScale.Text = ListBox1.Column(9)
        txtOriginalContract.Text = ListBox1.Column(10)
        txtEstimatedGross.Text = ListBox1.Column(11)
        txtBillings.Text = ListBox1.Column(12)
        txtIncurred.Text = ListBox1.Column(13)
        txtCostToComplete.Text = ListBox1.Column(14)

        End If
        
        
End Sub


Private Sub txtBillings_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Me.txtBillings = Format(Me.txtBillings, "$#,##0.00")
End Sub

Private Sub txtCostToComplete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Me.txtCostToComplete = Format(Me.txtCostToComplete, "$#,##0.00")
End Sub

Private Sub txtEstimatedGross_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Me.txtEstimatedGross = Format(Me.txtEstimatedGross, "$#,##0.00")
End Sub


Private Sub txtIncurred_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Me.txtIncurred = Format(Me.txtIncurred, "$#,##0.00")
End Sub

Private Sub txtOriginalContract_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Me.txtOriginalContract = Format(Me.txtOriginalContract, "$#,##0.00")
End Sub

Private Sub UserForm_Activate()
        cmbState.List = Array("MD", "VA", "DC")
        cmbBonded.List = Array("Yes", "No")
        cmbWageScale.List = Array("Yes", "No")
        cmbStatus.List = Array("Open", "Completed")
        cmbEstimator.List = Array("Kata", "Kayla", "Kerri", "Sawsan", "Tony", "Christian")
        Call Refresh_data
        
        Me.Left = Application.Left + (Application.Width - Me.Width) / 2
        
    
End Sub
