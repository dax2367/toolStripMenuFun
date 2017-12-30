Option Strict On

Public Class ToolStripMenuFum
    '§=================================================================================================================
    ' Author:	Holly Eaton
    ' 
    ' Program:	Tool Strip Menu Fun
    ' 
    ' Dev Env:	Visual Studio
    ' 
    ' Description:
    '   Purpose:    Project that explores the Tool Strip Menu inside of the U.S. Capitals Study Game.
    '  
    '   Input:      The user can select items from three menus. 
    '
    '   Process:    The File menu contains Reset and Exit options.
    '               The View menu contains Font and Color selection options.
    '               THe Help menu contains Show me the answer and About options.
    '               (The Show me the answer menu option is just a non-functioning place holder at this time.)
    '  
    '   The Tool Strip Menu code starts just after the code for the U.S. Capitals Study Game; on line 187
    '
    '   Comments for the U.S. Capitals Study Game have been removed for this project. If you wish to see the comments
    '   and code for the U.S. Capitals Game please check out that repo at https://github.com/dax2367/stateCapitalsGame
    ' 
    '==================================================================================================================
    '==================================================================================================================  
    '==================================================================================================================
    '==================================================================================================================
    '    
    'Class level variables use in multiple subroutines
    Private intScore As Integer
    Private intTotalQuestions As Integer
    Private strCorrectCapital As String
    Private sngPercentCorrect As Single
    Private strStatesArray() As String = {"Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming"}
    Private strCapitalsArray() As String = {"Montgomery", "Juneau", "Phoenix", "Little Rock", "Sacramento", "Denver", "Hartford", "Dover", "Tallahassee", "Atlanta", "Honolulu", "Boise", "Springfield", "Indianapolis", "Des Moines", "Topeka", "Frankfort", "Baton Rouge", "Augusta", "Annapolis", "Boston", "Lansing", "St. Paul", "Jackson", "Jefferson City", "Helena", "Lincoln", "Carson City", "Concord", "Trenton", "Santa Fe", "Albany", "Raleigh", "Bismarck", "Columbus", "Oklahoma City", "Salem", "Harrisburg", "Providence", "Columbia", "Pierre", "Nashville", "Austin", "Salt Lake City", "Montpelier", "Richmond", "Olympia", "Charleston", "Madison", "Cheyenne"}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'things that happen when the application window first appears
        MessageBox.Show("Welcome to the State Capitals Game. " & vbCrLf & "Just click on a radiobutton to select your answer.")

        doNewQuestion()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub doNewQuestion()
        'Local variables:
        Dim intStateRandomNumber As Integer
        Dim intFirstRandomNumber As Integer
        Dim intSecondRandomNumber As Integer
        Dim intThirdRandomNumber As Integer
        Dim intForthRandomNumber As Integer
        Dim intReplaceRandomAnswer As Integer

        'Instantiate a new instance of the random class
        Dim RandomClassInstance As New Random

        'Use the random number to pull out a state from the states array and the matching capital from the capitals array.
        'Generate a random number between 0 and 49
        intStateRandomNumber = RandomClassInstance.Next(0, 50)

        'Update the lblStateQuestion label question
        lblStateQuestion.Text = "What is the capital of " & strStatesArray(intStateRandomNumber) & "?"
        'Save the correct answer, the capital name, in strCorrectCapital
        strCorrectCapital = strCapitalsArray(intStateRandomNumber)

        'Update each of the four answers with randomly selected capital names
        intFirstRandomNumber = RandomClassInstance.Next(0, 50)
        rad1.Text = strCapitalsArray(intFirstRandomNumber)
        While intFirstRandomNumber = intSecondRandomNumber
            intSecondRandomNumber = RandomClassInstance.Next(0, 50)
        End While
        rad2.Text = strCapitalsArray(intSecondRandomNumber)
        While intFirstRandomNumber = intThirdRandomNumber Or intSecondRandomNumber = intThirdRandomNumber
            intThirdRandomNumber = RandomClassInstance.Next(0, 50)
        End While
        rad3.Text = strCapitalsArray(intThirdRandomNumber)
        While intFirstRandomNumber = intForthRandomNumber Or intSecondRandomNumber = intForthRandomNumber Or intThirdRandomNumber = intForthRandomNumber
            intForthRandomNumber = RandomClassInstance.Next(0, 50)
        End While
        rad4.Text = strCapitalsArray(intForthRandomNumber)

        'Replace one of the four answers with the correct answer.
        intReplaceRandomAnswer = RandomClassInstance.Next(0, 4)

        'use a case statement and that number...
        Select Case intReplaceRandomAnswer
            Case 0
                rad1.Text = strCorrectCapital
            Case 1
                rad2.Text = strCorrectCapital
            Case 2
                rad3.Text = strCorrectCapital
            Case 3
                rad4.Text = strCorrectCapital

        End Select

        'Uncheck all four answers.
        rad1.Checked = False
        rad2.Checked = False
        rad3.Checked = False
        rad4.Checked = False





    End Sub

    Private Sub allRadioButtons_CheckedChanged(sender As Object, e As EventArgs) Handles rad1.CheckedChanged, rad2.CheckedChanged, rad3.CheckedChanged, rad4.CheckedChanged
        'detect which button has fired the event
        Dim selectedRadioButton As RadioButton
        selectedRadioButton = CType(sender, RadioButton)

        If selectedRadioButton.Checked Then
            Select Case doCheckAnswer(selectedRadioButton)
                Case True
                    'stuff to do if true/correct
                    'friendly Beep
                    Beep()
                    'or My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Exclamation)

                    'message box with congratulations etc.
                    MessageBox.Show("That's it!  Good Job!")

                    'increment the number of correct responses and total questions
                    intScore += 1
                    intTotalQuestions += 1
                    sngPercentCorrect = CSng(intScore / intTotalQuestions)

                    lblScore.Text = intScore.ToString & " of " & intTotalQuestions.ToString & " quesions correct."
                    lblPercent.Text = sngPercentCorrect.ToString("P") & " correct"

                    doNewQuestion()
                Case False
                    'stuff to do if false/incorrect

                    'you missed it beep
                    My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Asterisk)

                    'message box telling user their response was incorrect and show the correct answer.
                    MessageBox.Show("Sorry that is incorrect " & vbCrLf & "The correct answer is " & strCorrectCapital)

                    'increment the number of total questions and find the percent correct
                    intTotalQuestions += 1
                    sngPercentCorrect = CSng(intScore / intTotalQuestions)

                    lblScore.Text = intScore.ToString & " of " & intTotalQuestions.ToString & " quesions correct."
                    lblPercent.Text = sngPercentCorrect.ToString("P") & " correct"

                    doNewQuestion()
            End Select
        End If

    End Sub

    Private Function doCheckAnswer(selectedRadioButton As RadioButton) As Boolean
        Dim blnIsItCorrect As Boolean

        If selectedRadioButton.Text = strCorrectCapital Then
            blnIsItCorrect = True
        Else
            blnIsItCorrect = False
        End If

        Return blnIsItCorrect
    End Function

    Private Sub btnNewPlayer_Click(sender As Object, e As EventArgs) Handles btnNewPlayer.Click
        'reset the environment for the new player
        lblScore.Text = " "
        lblPercent.Text = " "
        intScore = 0
        intTotalQuestions = 0

        MessageBox.Show("New Player, Welcome to the State Capitals Game. " & vbCrLf & "Just click on a radiobutton to select your answer.")

        doNewQuestion()

    End Sub

    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontToolStripMenuItem.Click
        'show a dialog box that will allow users to change the font in the form.
        Try
            fdFont.ShowDialog()
            'apply the font if the progam can
            Me.Font = fdFont.Font
            lblStateQuestion.Font = fdFont.Font
        Catch ex As Exception
            'let the user know there was a problem
            MessageBox.Show("Sorry, only TrueType fonts work")
        End Try
    End Sub

    Private Sub ColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ColorToolStripMenuItem.Click
        'when the color dialog box opens it shows the current background color of the form highlighted
        cdColor.Color = Me.BackColor
        cdColor.ShowDialog()

        'apply whatever new color users choose for the background
        Me.BackColor = cdColor.Color
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        'close the program
        Me.Close()
    End Sub

    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        'reset the environment
        lblScore.Text = " "
        lblPercent.Text = " "
        intScore = 0
        intTotalQuestions = 0

        MessageBox.Show("The State Capitals Game Has Been Reset . " & vbCrLf & "Just click on a radiobutton to select your answer.")

        doNewQuestion()
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        'close the program
        Me.Close()
    End Sub

    Private Sub ResetToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem1.Click
        'reset the environment
        lblScore.Text = " "
        lblPercent.Text = " "
        intScore = 0
        intTotalQuestions = 0

        MessageBox.Show("The State Capitals Game Has Been Reset . " & vbCrLf & "Just click on a radiobutton to select your answer.")

        doNewQuestion()
    End Sub

    Private Sub ShowMeTheAnswerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowMeTheAnswerToolStripMenuItem.Click
        'message box showing the user the correct answer.
        MessageBox.Show("The correct answer is " & vbCrLf & vbCrLf & strCorrectCapital)
    End Sub

    Private Sub ShowMeTheAnswerToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ShowMeTheAnswerToolStripMenuItem1.Click
        'message box showing the user the correct answer.
        MessageBox.Show("The correct answer is " & vbCrLf & vbCrLf & strCorrectCapital)

        'increment the number of total questions and find the percent correct
        intTotalQuestions += 1
        sngPercentCorrect = CSng(intScore / intTotalQuestions)

        lblScore.Text = intScore.ToString & " of " & intTotalQuestions.ToString & " quesions correct."
        lblPercent.Text = sngPercentCorrect.ToString("P") & " correct"

        'show a new question
        doNewQuestion()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        'message box showing the author and copyright.
        MessageBox.Show("Author: Holly Eaton " & vbCrLf & vbCrLf & "Making menus is fun!")
    End Sub
End Class