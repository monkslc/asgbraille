Attribute VB_Name = "Layouts"





Sub layoutFromCSV()


    Dim LayoutDataPath As String
    Dim NodePath As String
    Dim TranslatorPath As String
    LayoutDataPath = "C:\Users\Michelle\Documents\LayoutData.csv"
    NodePath = "C:\Program Files\nodejs\node"
    TranslatorPath = "C:\Users\Michelle\Documents\ADASignageBrailleTranslator\src\model\main.js"
    
    Open LayoutDataPath For Input As #1 ' File Path


    ' Determine if we should enforce california braille
    Dim enforceCaliforniaBraille As Integer
    enforceCaliforniaBraille = 0
    Line Input #1, LineFromFile
    LineItems = Split(LineFromFile, ",")
    If LineItems(1) = "yes" Then
        enforceCaliforniaBraille = 1
    End If


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
    
    ' HERE We want to create an array that will store all of the transformations we need to make
    Dim brailleText() As String
    Dim brailleObjects As New ShapeRange
    Dim brailleAlignments() As Double
    Dim brailleReplacements() As String
 
    Dim brailleCounter As Integer
    brailleCounter = 0
    
    Dim objectPosition() As Integer
    Dim multipleReplacementsCounter As Integer
    multipleReplacementsCounter = 0
    
    ' Create variables to keep track of the sign number
    Dim signNumber As Integer
    signNumber = 1
    Dim textShrinks() As String
    Dim numOfTextShrinks As Integer
    numOfTextShrink = 0

    Do Until EOF(1)

        If xCounter + ActiveSelection.SizeWidth > layoutWidth Then
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

                    alignmentPos = sh.centerX
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
                        If LineItems(index - 1) = "delete" Then
                            ' Something goes here
                            sh.Text.Story = ""
                            
                        Else
                            sh.Text.Story = LineItems(index - 1)
                        
                        End If
                        

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
                            ' Let the user know we shrunk this text object
                            ReDim Preserve textShrinks(numOfTextShrinks)
                            textShrinks(numOfTextShrinks) = CStr(signNumber)
                            numOfTextShrinks = numOfTextShrinks + 1
                            Call sh.SetSize(maxWidth, sh.SizeHeight)
                        End If


                        ' Adjusting Alignment for the sign
                        If sh.Text.Story.Alignment = cdrCenterAlignment Then
                            sh.centerX = alignmentPos
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
                Dim numBrailleReplacements As Integer
                numBrailleReplacements = 0
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
                        
                        
                        
                        
                        
                        
                        ' Add the braille object and the braille string to the array
                        ReDim Preserve brailleText(brailleCounter)
                        ReDim Preserve brailleReplacements(brailleCounter)
                        ReDim Preserve objectPosition(brailleCounter)
                        ReDim Preserve brailleAlignments(brailleCounter)
                        brailleText(brailleCounter) = LineItems(index - 1)
                        brailleReplacements(brailleCounter) = braille
                        brailleAlignments(brailleCounter) = alignmentPos
                        brailleObjects.Add sh
                        
                        ' Ensure were tracking the correct position of the braille object
                        If numBrailleReplacements <> 0 Then
                            multipleReplacementsCounter = multipleReplacementsCounter + 1
                        End If
                        
                        objectPosition(brailleCounter) = brailleCounter - multipleReplacementsCounter
                        brailleCounter = brailleCounter + 1
                        
                        numBrailleReplacements = numBrailleReplacements + 1
               

                        ' Check for california compliance
                        If sh.Text.Story.Font <> "CA Compliant Braille 060" And enforceCaliforniaBraille <> 0 Then
                            MsgBox ("Not California Braille Compliant")
                            Close #1
                            Exit Sub
                        End If



                    
                    End If
                Next index

            End If


       Next sh
       signNumber = signNumber + 1
       tempRange.CreateSelection
       
       
       

    Loop

    ' Lets see if we got all of our braille objects
    ' Don't forget to put alignment here
    
    ' Set the braille
    Dim cmd As String
    Dim brailleCommand As String
    brailleCommand = Join(brailleText, "~")
    
    cmd = "cmd.exe /c echo " & brailleCommand & " | " & """" & NodePath & """" & " " & """" & TranslatorPath & """"

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    Dim oExec As Object
    Dim oOutput As Object

    Set oExec = oShell.Exec(cmd)
    Set oOutput = oExec.StdOut
    
    
    Dim sLine As String
    sLine = oOutput.ReadLine
    
    ' We want to loop through each braille translation we made and replace it in the appropriate text object
    Dim translations() As String
    translations = Split(sLine, "~")
    Dim i As Integer
    For i = 0 To UBound(translations) - LBound(translations) ' + 1
        ' Here is where we want to actually make the translation
    
        Dim brailleObj As Shape
        Set brailleObj = brailleObjects.Shapes(objectPosition(i) + 1)
        
        Dim storyText As String
        storyText = brailleObj.Text.Story
        If InStr(storyText, "delete") Then
            brailleObj.Text.Story = Replace(brailleObj.Text.Story, brailleReplacements(i), "")
            
        Else
             brailleObj.Text.Story = Replace(brailleObj.Text.Story, brailleReplacements(i), translations(i))
        End If
       
        
        ' Now we want to make the alignment adjustments
        If brailleObj.Text.Story.Alignment = cdrCenterAlignment Then
            brailleObj.centerX = brailleAlignments(i)

        ElseIf brailleObj.Text.Story.Alignment = cdrLeftAlignment Then
            brailleObj.LeftX = brailleAlignments(i)

        ElseIf sh.Text.Story.Alignment = cdrRightAlignment Then
            brailleObj.RightX = brailleAlignments(i)

        Else
            MsgBox ("Specify an alignment for the braille")
            Close #1
            Exit Sub
        End If
    Next i
    
    ' Let the user know where we shrunk text
    Dim shrinkMessage As String
    shrinkMessage = Join(textShrinks, ", ")
    If shrinkMessage <> "" Then
        MsgBox ("Shrunk text on signs: " & shrinkMessage)
    End If
    

    Close #1

End Sub
