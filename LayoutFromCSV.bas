Attribute VB_Name = "Layouts"





Sub layoutFromCSV()


    Dim FilePath As String
    Open "C:\Users\Big Cell Engraver\Desktop\LayoutData.csv" For Input As #1



    ' Determine spacing for layout
    Dim spacing As Double
    Line Input #1, LineFromFile
    LineItems = Split(LineFromFile, ",")

    spacing = LineItems(1)

    ' Determine layout width from csv
    Dim layoutWidth As Double
    Line Input #1, LineFromFile
    LineItems = Split(LineFromFile, ",")
    layoutWidth = LineItems(1)


    'Get max text widths from csv
    Line Input #1, LineFromFile
    maxTextWidths = Split(LineFromFile, ",")


    ' Set offset variables
    Dim xOffset As Double
    Dim xCounter As Double
    Dim yOffset As Double

    xOffset = 0
    xCounter = ActiveSelection.SizeWidth + spacing
    yOffset = 0

    Do Until EOF(1)

        If xCounter + ActiveSelection.SizeWidth + spacing > layoutWidth Then
            xOffset = 0 - xCounter + ActiveSelection.SizeWidth + spacing
            xCounter = ActiveSelection.SizeWidth + spacing
            yOffset = 0 - ActiveSelection.SizeHeight - spacing

        Else
            xOffset = ActiveSelection.SizeWidth + spacing
            xCounter = xCounter + xOffset
            yOffset = 0
        End If


       Line Input #1, LineFromFile
       LineItems = Split(LineFromFile, ",")

       Dim tempRange As ShapeRange
       Set tempRange = ActiveSelectionRange.Duplicate(xOffset, yOffset)


       Dim sh As Shape
       For Each sh In ActiveSelection.Shapes

            Dim signWidth As Double
            signWidth = ActiveSelection.SizeWidth

            If sh.Type = cdrTextShape Then

                Dim alignmentPos As Double
                If sh.Text.Story.Alignment = cdrCenterAlignment Then

                    alignmentPos = sh.CenterX
                End If

                If sh.Text.Story.Alignment = cdrLeftAlignment Then
                    alignmentPos = sh.LeftX

                End If

                If sh.Text.Story.Alignment = cdrRightAlignment Then
                    alignmentPos = sh.RightX
                End If

                ' Loop through all possible text variables
                Dim index As Integer
                For index = 1 To 7
                    If sh.Text.Story = "Text" + CStr(index) Then

                        ' Check to see if we are passing empty cells
                        If LineItems(index - 1) = "" Then
                            tempRange.Delete
                            Close #1
                            Exit Sub
                        End If


                        ' Set the text
                        sh.Text.Story = LineItems(index - 1)

                        ' Adjusting Width for the text
                        Dim maxWidth As Double
                        Dim check As Double

                        If CDbl(maxTextWidths(index - 1)) = CDbl(-1) Then

                            ' Try and interpret a maxwidth
                            Dim edgeMargin As Double
                            edgeMargin = 1 / 8


                            maxWidth = 50
                            If sh.Text.Story.Alignment = cdrCenterAlignment Then
                                maxWidth = signWidth - (edgeMargin * 2)

                            ElseIf sh.Text.Story.Alignment = cdrLeftAlignment Then
                                maxWidth = signWidth - (alignmentPos - ActiveSelection.LeftX) - edgeMargin

                            ElseIf sh.Text.Story.Alignment = cdrRightAlignment Then
                                maxWidth = signWidth - (ActiveSelection.RightX - alignmentPos) - edgeMargin

                            Else
                                MsgBox ("Specify an alignment for all text")
                                Close #1
                                Exit Sub
                            End If

                        Else
                            maxWidth = CDbl(maxTextWidths(index - 1))

                        End If

                        If (maxWidth < sh.SizeWidth) Then
                            Call sh.SetSize(maxWidth, sh.SizeHeight)
                        End If


                        ' Adjusting Alignment for the sign
                        If sh.Text.Story.Alignment = cdrCenterAlignment Then
                            sh.CenterX = alignmentPos
                        End If

                        If sh.Text.Story.Alignment = cdrLeftAlignment Then
                            sh.LeftX = alignmentPos
                        End If

                        If sh.Text.Story.Alignment = cdrRightAlignment Then
                            sh.RightX = alignmentPos
                        End If



                    End If
                Next index

                ' LOOP THROUGH ALL POSSIBLE BRAILLE VARIABLES
                For index = 1 To 7
                    Dim braille As String
                    braille = "Braille" & CStr(index)
                    If InStr(sh.Text.Story, braille) Then
                        
                        
                        ' Check if cell is empty
                        If LineItems(index - 1) = "" Then
                            tempRange.Delete
                            Close #1
                            Exit Sub
                        End If
                        
                        ' Set the braille
                        Dim cmd As String
                        
                        cmd = "cmd.exe /c echo " & LineItems(index - 1) & " |  ""C:\Program Files\nodejs\node"" ""C:\Users\Big Cell Engraver\Documents\ADASignageBrailleTranslator\src\model\main.js"""
                        Dim oShell As Object
                        Set oShell = CreateObject("WScript.Shell")

                        Dim oExec As Object
                        Dim oOutput As Object

                        Set oExec = oShell.Exec(cmd)
                        Set oOutput = oExec.StdOut
                        
                        
                        Dim sLine As String
                        sLine = oOutput.ReadLine

                        Dim newLineText As String
                        newLineText = Replace(sh.Text.Story, braille, sLine)
                        

                        sh.Text.Story = newLineText
               

                        ' Check for california compliance
                        If sh.Text.Story.Font <> "CA Compliant Braille 060" Then
                            MsgBox ("Not California Braille Compliant")
                            Close #1
                            Exit Sub
                        End If



                    ' Adjusting Alignment for the sign
                        If sh.Text.Story.Alignment = cdrCenterAlignment Then
                            sh.CenterX = alignmentPos

                        ElseIf sh.Text.Story.Alignment = cdrLeftAlignment Then
                            sh.LeftX = alignmentPos

                        ElseIf sh.Text.Story.Alignment = cdrRightAlignment Then
                            sh.RightX = alignmentPos

                        Else
                            MsgBox ("Specify an alignment for the braille")
                            Close #1
                            Exit Sub
                        End If
                    End If
                Next index

            End If


       Next sh

       tempRange.CreateSelection

    Loop

    Close #1

End Sub
