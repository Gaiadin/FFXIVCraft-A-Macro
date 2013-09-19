Public Class frmMain
    Dim intMacroIndex As Integer = 2
    Dim intLineIndex() As Integer
    Dim strMacros() As String
    Dim intCP(49, 15) As Integer
    Dim intTotalCP As Integer
    Dim intBaseCp As Integer


    Private Sub btnBasicSynth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBasicSynth.Click

        CheckFirst(0)
        txtOutput.Text += "/ac ""Basic Synthesis"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnStandardSynth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStandardSynth.Click

        CheckFirst(15)
        txtOutput.Text += "/ac ""Standard Synthesis"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)
    End Sub

    Private Sub btnBasicTouch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBasicTouch.Click

        CheckFirst(18)
        txtOutput.Text += "/ac ""Basic Touch"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)
    End Sub

    Private Sub btnStandardTouch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStandardTouch.Click

        CheckFirst(32)
        txtOutput.Text += "/ac ""Standard Touch"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)
    End Sub

    Private Sub btnAdvancedTouch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdvancedTouch.Click

        CheckFirst(48)
        txtOutput.Text += "/ac ""Advanced Touch"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Setup all tooltip information
        ToolTip.SetToolTip(btnBasicSynth, "Basic Synthesis" & vbNewLine & "Increases progress" & vbNewLine & "Efficiency: 100%" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnStandardSynth, "Standard Synthesis" & vbNewLine & "CP: 15" & vbNewLine & "Increases progress" & vbNewLine & "Efficiency: 150%" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnBasicTouch, "Basic Touch" & vbNewLine & "CP: 18" & vbNewLine & "Increases quality" & vbNewLine & "Efficiency: 100%" & vbNewLine & "Success Rate: 70%")
        ToolTip.SetToolTip(btnStandardTouch, "Standard Touch" & vbNewLine & "CP: 32" & vbNewLine & "Increases quality" & vbNewLine & "Efficiency: 125%" & vbNewLine & "Success Rate: 80%")
        ToolTip.SetToolTip(btnAdvancedTouch, "Advanced Touch" & vbNewLine & "CP: 48" & vbNewLine & "Increases quality" & vbNewLine & "Efficiency: 150%" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnMastersMend, "Master's Mend" & vbNewLine & "CP: 92" & vbNewLine & "Restores item durability by 30")
        ToolTip.SetToolTip(btnMastersMend2, "Master's Mend II" & vbNewLine & "CP: 160" & vbNewLine & "Restores item durability by 60")
        ToolTip.SetToolTip(btnSteadyHand, "Steady Hand" & vbNewLine & "CP: 22" & vbNewLine & "Improves action success rate by" & vbNewLine & "20% for the next five steps")
        ToolTip.SetToolTip(btnInnerQuiet, "Inner Quiet" & vbNewLine & "CP: 18" & vbNewLine & "Grants a bonus to control with" & vbNewLine & "every increase in quality")
        ToolTip.SetToolTip(btnObserve, "Observe" & vbNewLine & "CP: 14" & vbNewLine & "Do nothing for one step")
        ToolTip.SetToolTip(btnGreatStrides, "Great Strides" & vbNewLine & "CP: 32" & vbNewLine & "Doubles efficiency of next" & vbNewLine & "touch action. Effect active for" & vbNewLine & "three steps")
        ToolTip.SetToolTip(btnRapidSynth, "Rapid Synthesis" & vbNewLine & "Increases progress" & vbNewLine & "Efficiency: 250%" & vbNewLine & "Success Rate: 50%")
        ToolTip.SetToolTip(btnBrandOfIce, "Brand of Ice" & vbNewLine & "CP: 15" & vbNewLine & "Increases progress. Progress doubles" & vbNewLine & "when recipe affinity is ice" & vbNewLine & "Efficiency: 100% (200%)" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnBrandOfFire, "Brand of Fire" & vbNewLine & "CP: 15" & vbNewLine & "Increases progress. Progress doubles" & vbNewLine & "when recipe affinity is fire" & vbNewLine & "Efficiency: 100% (200%)" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnBrandOfWind, "Brand of Wind" & vbNewLine & "CP: 15" & vbNewLine & "Increases progress. Progress doubles" & vbNewLine & "when recipe affinity is wind" & vbNewLine & "Efficiency: 100% (200%)" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnBrandOfEarth, "Brand of Earth" & vbNewLine & "CP: 15" & vbNewLine & "Increases progress. Progress doubles" & vbNewLine & "when recipe affinity is earth" & vbNewLine & "Efficiency: 100% (200%)" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnBrandOfLightning, "Brand of Lightning" & vbNewLine & "CP: 15" & vbNewLine & "Increases progress. Progress doubles" & vbNewLine & "when recipe affinity is wind" & vbNewLine & "Efficiency: 100% (200%)" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnBrandOfWater, "Brand of Water" & vbNewLine & "CP: 15" & vbNewLine & "Increases progress. Progress doubles" & vbNewLine & "when recipe affinity is water" & vbNewLine & "Efficiency: 100% (200%)" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnPieceByPiece, "Piece by Piece" & vbNewLine & "CP: 15" & vbNewLine & "Increases remaining progress by 1/3" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnIngenuity, "Ingenuity" & vbNewLine & "CP: 24" & vbNewLine & "Lowers recipe level to current level" & vbNewLine & "for the next five steps")
        ToolTip.SetToolTip(btnIngenuity2, "Ingenuity II" & vbNewLine & "CP: 32" & vbNewLine & "Lowers recipe level three below" & vbNewLine & "current level for the next five steps")
        ToolTip.SetToolTip(btnRumination, "Rumination" & vbNewLine & "Removes Inner Quiet effect and restores CP" & vbNewLine & "proportional to the number of times" & vbNewLine & "control was increased")
        ToolTip.SetToolTip(btnByregotsBlessing, "Byregot's Blessing" & vbNewLine & "CP: 24" & vbNewLine & "Increases quality" & vbNewLine & "Efficiency: 100% plus 20% for each bonus" & vbNewLine & "to control granted by Inner Quiet" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnHastyTouch, "Hasty Touch" & vbNewLine & "Increases quality" & vbNewLine & "Efficiency: 100%" & vbNewLine & "Success Rate: 50%")
        ToolTip.SetToolTip(btnSteadyHand2, "Steady Hand II" & vbNewLine & "CP: 25" & vbNewLine & "Improves action success rate by 30%" & vbNewLine & "for the next five steps")
        ToolTip.SetToolTip(btnReclaim, "Reclaim" & vbNewLine & "CP: 55" & vbNewLine & "Increases the chance materials will not be" & vbNewLine & "lost after botched synthesis to 90%")
        ToolTip.SetToolTip(btnManipulation, "Manipulation" & vbNewLine & "CP: 88" & vbNewLine & "Restores 10 points of durability after" & vbNewLine & "each step for the next three steps")
        ToolTip.SetToolTip(btnFlawlessSynth, "Flawless Synthesis" & vbNewLine & "CP: 37" & vbNewLine & "Increases progress by 40" & vbNewLine & "Success Rate: 90%")
        ToolTip.SetToolTip(btnInnovation, "Innovation" & vbNewLine & "CP: 18" & vbNewLine & "Increases control by 50% for" & vbNewLine & "the next three steps")
        ToolTip.SetToolTip(btnWasteNot, "Waste Not" & vbNewLine & "CP: 56" & vbNewLine & "Reduces loss of durability by 50% for the next four steps")
        ToolTip.SetToolTip(btnWasteNot2, "Waste Not II" & vbNewLine & "CP: 98" & vbNewLine & "Reduces loss of durability by 50% for the next eight steps")
        ToolTip.SetToolTip(btnCarefulSynth, "Careful Synthesis" & vbNewLine & "Increases progress" & vbNewLine & "Efficiency: 90%" & vbNewLine & "Success Rate: 100%")
        ToolTip.SetToolTip(btnCarefulSynth2, "Careful Synthesis II" & vbNewLine & "Increases progress" & vbNewLine & "Efficiency: 120%" & vbNewLine & "Success Rate: 100%")
        ToolTip.SetToolTip(btnTricksOfTheTrade, "Tricks of the Trade" & vbNewLine & "Restores 20 cp. Can only be used" & vbNewLine & "when material condition is good")
        ToolTip.SetToolTip(btnComfortZone, "Comfort Zone" & vbNewLine & "CP: 66" & vbNewLine & "Restores 8 CP after each step" & vbNewLine & "for the next ten steps")
        ToolTip.SetToolTip(chkAutoCopy, "Automatically copies current macro to clipboard as it is changed")
        ToolTip.SetToolTip(chkAlert, "Show a chat alert as the last step of each" & vbNewLine & "macro in a multi-macro series to let you know" & vbNewLine & "when to use the next macro")
        ToolTip.SetToolTip(txtWaitBuff, "Wait time after buff-type abilities" & vbNewLine & "(Great Strides, Steady Hand, etc.)")
        ToolTip.SetToolTip(txtWaitAction, "Wait time after action-type abilities" & vbNewLine & "(Synthesis, Touch, etc.)")
        ToolTip.SetToolTip(btnClear, "Clear current macro")
        ToolTip.SetToolTip(btnClearAll, "Clear all macros")
        ToolTip.SetToolTip(btnEraseStep, "Go back one step")
        tabMacro.TabPages.Add(0, "Macro 1")
        ReDim strMacros(0)
        ReDim intLineIndex(0)
        tabMacro.SelectedIndex = 0
    End Sub

    Private Sub btnMastersMend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMastersMend.Click

        CheckFirst(92)
        txtOutput.Text += "/ac ""Master's Mend"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnMastersMend2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMastersMend2.Click

        CheckFirst(160)
        txtOutput.Text += "/ac ""Master's Mend II"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnSteadyHand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSteadyHand.Click

        CheckFirst(22)
        txtOutput.Text += "/ac ""Steady Hand"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnInnerQuiet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInnerQuiet.Click

        CheckFirst(18)
        txtOutput.Text += "/ac ""Inner Quiet"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnObserve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnObserve.Click

        CheckFirst(14)
        txtOutput.Text += "/ac ""Observe"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnGreatStrides_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGreatStrides.Click

        CheckFirst(32)
        txtOutput.Text += "/ac ""Great Strides"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnTricksOfTheTrade_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTricksOfTheTrade.Click

        CheckFirst(-20)
        txtOutput.Text += "/ac ""Tricks of the Trade"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnBrandOfWater_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrandOfWater.Click

        CheckFirst(15)
        txtOutput.Text += "/ac ""Brand of Water"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnComfortZone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnComfortZone.Click

        CheckFirst(66)
        txtOutput.Text += "/ac ""Comfort Zone"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnManipulation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManipulation.Click

        CheckFirst(88)
        txtOutput.Text += "/ac ""Manipulation"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnFlawlessSynth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFlawlessSynth.Click

        CheckFirst(37)
        txtOutput.Text += "/ac ""Flawless Synthesis"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnInnovation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInnovation.Click

        CheckFirst(18)
        txtOutput.Text += "/ac ""Innovation"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnIngenuity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIngenuity.Click

        CheckFirst(24)
        txtOutput.Text += "/ac ""Ingenuity"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnBrandOfFire_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrandOfFire.Click

        CheckFirst(15)
        txtOutput.Text += "/ac ""Brand of Fire"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnIngenuity2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIngenuity2.Click

        CheckFirst(32)
        txtOutput.Text += "/ac ""Ingenuity II"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnCarefulSynth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCarefulSynth.Click

        CheckFirst(0)
        txtOutput.Text += "/ac ""Careful Synthesis"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnBrandOfLightning_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrandOfLightning.Click

        CheckFirst(15)
        txtOutput.Text += "/ac ""Brand of Lightning"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnCarefulSynth2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCarefulSynth2.Click

        CheckFirst(0)
        txtOutput.Text += "/ac ""Careful Synthesis II"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnRapidSynth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRapidSynth.Click

        CheckFirst(0)
        txtOutput.Text += "/ac ""Rapid Synthesis"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnBrandOfIce_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrandOfIce.Click

        CheckFirst(15)
        txtOutput.Text += "/ac ""Brand of Ice"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnPieceByPiece_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPieceByPiece.Click

        CheckFirst(15)
        txtOutput.Text += "/ac ""Piece by Piece"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnHastyTouch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHastyTouch.Click

        CheckFirst(0)
        txtOutput.Text += "/ac ""Hasty Touch"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnSteadyHand2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSteadyHand2.Click

        CheckFirst(25)
        txtOutput.Text += "/ac ""Steady Hand II"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnReclaim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReclaim.Click

        CheckFirst(55)
        txtOutput.Text += "/ac ""Reclaim"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnRumination_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRumination.Click

        CheckFirst(0)
        txtOutput.Text += "/ac ""Rumination"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnBrandOfWind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrandOfWind.Click

        CheckFirst(15)
        txtOutput.Text += "/ac ""Brand of Wind"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnByregotsBlessing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnByregotsBlessing.Click

        CheckFirst(24)
        txtOutput.Text += "/ac ""Byregot's Blessing"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnWasteNot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWasteNot.Click

        CheckFirst(56)
        txtOutput.Text += "/ac ""Waste Not"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnBrandOfEarth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrandOfEarth.Click

        CheckFirst(15)
        txtOutput.Text += "/ac ""Brand of Earth"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(False)

    End Sub

    Private Sub btnWasteNot2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWasteNot2.Click

        CheckFirst(98)
        txtOutput.Text += "/ac ""Waste Not II"" <me>"
        intLineIndex(tabMacro.SelectedIndex) += 1
        Wait(True)

    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        If intMacroIndex = 2 Then
            ClearAll()
            Return
        End If
        txtOutput.Clear()
        ClearMacro()
        intMacroIndex -= 1
    End Sub

    Private Sub btnEraseStep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEraseStep.Click
        'Create a list of lines from the text box
        Dim lstText As List(Of String) = txtOutput.Lines.ToList()

        If intLineIndex(tabMacro.SelectedIndex) = 0 Then
            'Nothing to erase
            Return
        End If
        If intLineIndex(tabMacro.SelectedIndex) = 2 Then
            'First step in macro
            If intMacroIndex = 2 Then
                ClearAll()
                Return
            End If
            intMacroIndex -= 1

            lstText.RemoveAt(txtOutput.Lines.Count - 1)
            lstText.RemoveAt(txtOutput.Lines.Count - 2)
            intLineIndex(tabMacro.SelectedIndex) -= 2
            txtOutput.Lines = lstText.ToArray
            ClearMacro()
            Return
        ElseIf intLineIndex(tabMacro.SelectedIndex) = 15 Then
            'The last step in a macro, remove only one line
            lstText.RemoveAt(txtOutput.Lines.Count - 1)
            intLineIndex(tabMacro.SelectedIndex) -= 1
            txtOutput.Lines = lstText.ToArray
        Else

            lstText.RemoveAt(txtOutput.Lines.Count - 1)
            lstText.RemoveAt(txtOutput.Lines.Count - 2)
            intLineIndex(tabMacro.SelectedIndex) -= 2
            txtOutput.Lines = lstText.ToArray
        End If
        Try
            intTotalCP += intCP(tabMacro.SelectedIndex, intLineIndex(tabMacro.SelectedIndex))
            intCP(tabMacro.SelectedIndex, intLineIndex(tabMacro.SelectedIndex)) = 0
            txtCP.Text = intTotalCP
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ClearMacro()

        Dim intTempCP As Integer
        'Add all of the CP values for each line in current macro
        For i = 0 To 14
            intTempCP += intCP(tabMacro.SelectedIndex, i)
        Next
        'Give CP back for skills undone
        intTotalCP += intTempCP
        txtCP.Text = intTotalCP
        If tabMacro.SelectedIndex = UBound(strMacros) Then
            'Last macro
            tabMacro.TabPages.RemoveByKey(UBound(strMacros))
            ReDim Preserve strMacros(UBound(strMacros) - 1)
            ReDim Preserve intLineIndex(UBound(intLineIndex) - 1)

            tabMacro.SelectedIndex = UBound(strMacros)
        Else
            'The first or a middle macro was deleted
            Dim count As Integer = tabMacro.SelectedIndex
            Do While count < UBound(strMacros)
                strMacros(count) = strMacros(count + 1)
                For i = 0 To 14
                    intCP(count, i) = intCP(count + 1, i)
                Next
                intLineIndex(count) = intLineIndex(count + 1)
                tabMacro.TabPages.Item(count + (UBound(strMacros) - count)).Text = tabMacro.TabPages.Item(count + (UBound(strMacros) - count) - 1).Text
                count += 1
            Loop
            tabMacro.TabPages.RemoveByKey(UBound(strMacros))
            ReDim Preserve strMacros(UBound(strMacros) - 1)
            ReDim Preserve intLineIndex(UBound(intLineIndex) - 1)
            tabMacro.SelectedIndex = count
        End If
    End Sub

    Private Sub txtOutput_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOutput.TextChanged
        Try
            strMacros(tabMacro.SelectedIndex) = txtOutput.Text
        Catch ex As Exception

        End Try
        txtOutput.SelectionStart = txtOutput.Text.Length
        txtOutput.ScrollToCaret()

        If chkAutoCopy.Checked = True Then
            CopyToClipboard()
        End If
    End Sub

    Private Sub CopyToClipboard()
        txtOutput.SelectAll()
        txtOutput.Copy()
    End Sub
    Private Sub CheckFirst(ByVal CP As Integer)
        If txtCP.ReadOnly = False Then
            txtCP.ReadOnly = True
            Try
                intBaseCp = txtCP.Text
            Catch ex As Exception

            End Try

        End If
        If intLineIndex(tabMacro.SelectedIndex) <> 15 And intLineIndex(tabMacro.SelectedIndex) <> 0 Then
            txtOutput.Text += vbNewLine
        End If
        If intLineIndex(tabMacro.SelectedIndex) = 14 Then
            'Final line of the macro.  Check for alert.
            If chkAlert.Checked = True Then
                txtOutput.Text += "/echo USE MACRO " & tabMacro.SelectedIndex + 2 & " NOW"
                intLineIndex(tabMacro.SelectedIndex) += 1
            End If
        End If
        If intLineIndex(tabMacro.SelectedIndex) = 15 Then
            'Macro full. Create new tab
            If tabMacro.SelectedIndex <> UBound(strMacros) Then
                'Not the last macro, append following action to next available page
                Dim i As Integer
                For i = tabMacro.SelectedIndex To UBound(strMacros)

                    If intLineIndex(i) < 15 Then
                        tabMacro.SelectedIndex = i
                        Exit For
                    End If
                Next
                'MessageBox.Show(UBound(strMacros))
                If i = UBound(strMacros) + 1 Then
                    'reached the last macro with no available space.  Create new tab
                    NewTab()

                ElseIf intLineIndex(tabMacro.SelectedIndex) <> 0 And intLineIndex(tabMacro.SelectedIndex) <> 15 Then
                    txtOutput.Text += vbNewLine
                End If
            Else
                NewTab()
            End If


            'txtOutput.Text += vbNewLine & "MACRO " & intMacroIndex & vbNewLine


            intCP(tabMacro.SelectedIndex, intLineIndex(tabMacro.SelectedIndex)) = CP
            intTotalCP -= CP
        Else
            intCP(tabMacro.SelectedIndex, intLineIndex(tabMacro.SelectedIndex)) = CP
            intTotalCP -= CP
        End If
        txtCP.Text = intTotalCP
    End Sub

    Private Sub NewTab()
        ReDim Preserve strMacros(UBound(strMacros) + 1)
        ReDim Preserve intLineIndex(UBound(intLineIndex) + 1)
        tabMacro.TabPages.Add(UBound(strMacros), "Macro " & intMacroIndex)
        intMacroIndex += 1
        tabMacro.SelectedIndex = UBound(strMacros)
        intLineIndex(tabMacro.SelectedIndex) = 0
    End Sub

    Private Sub Wait(ByVal buff As Boolean)
        If intLineIndex(tabMacro.SelectedIndex) < 15 Then
            If buff = True Then
                txtOutput.Text += vbNewLine & "/wait " & txtWaitBuff.Text
            Else
                txtOutput.Text += vbNewLine & "/wait " & txtWaitAction.Text
            End If

            intLineIndex(tabMacro.SelectedIndex) += 1
        End If
        txtOutput.Focus()
        txtOutput.DeselectAll()
    End Sub

    Private Sub tabMacro_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tabMacro.SelectedIndexChanged
        If tabMacro.SelectedIndex >= 0 Then
            txtOutput.Text = strMacros(tabMacro.SelectedIndex)
        End If

    End Sub

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click
        ClearAll()
    End Sub
    Private Sub ClearAll()
        intTotalCP = intBaseCp
        txtCP.ReadOnly = False
        ReDim intLineIndex(0)
        ReDim strMacros(0)
        intMacroIndex = 2
        txtOutput.Clear()
        tabMacro.TabPages.Clear()
        tabMacro.TabPages.Add(UBound(strMacros), "Macro 1")
        tabMacro.SelectedIndex = 0
        txtCP.Text = intTotalCP
    End Sub

    Private Sub txtCP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCP.TextChanged
        intTotalCP = txtCP.Text
    End Sub
End Class
