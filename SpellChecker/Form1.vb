Public Class Form1
    
    Dim Dictionary() As String              'As we don't know exact number of words in dictionary that might be used, we use dynamic array
    Dim options As String                   'If word is not in dictionary, here will be stored other options
    Dim showD As Boolean = False            'If user wants to show loaded dictionary
    Dim lineCount As Integer                'lineCount will store the number of lines in file chosen by user
    Dim showProgress As Boolean = True      'showing progress slows down sorting so user have option to switch it off
    Dim StopApp As Boolean          'This is for a function which stops loading of dictionary
    ' Dim WithEvents Timer1 As Timer
    Shared _timer As Timer


    Public Sub LoadDictionary(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim fdlg As OpenFileDialog = New OpenFileDialog()   'code taken from class material
        Dim lineNB As Integer
        Dim x As Integer
        Dim refresh As Integer = 0

        StopApp = False
        lineCount = 0

        fdlg.Title = "Open text file containing Dictionary"                                                         '\
        fdlg.InitialDirectory = "C:\"                                                                               ' }code taken from class material
        fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"                                                     '/

        If fdlg.ShowDialog() = DialogResult.OK Then 'if user choose file and press OK then...

            If System.IO.File.Exists(fdlg.FileName) = True Then 'in case user types in the name of the file

                Dim objReader1 As New System.IO.StreamReader(fdlg.FileName)

                Do While objReader1.Peek() <> -1    'this is to find out how many words is in dictionary
                    objReader1.ReadLine()
                    lineCount = lineCount + 1
                Loop

                If lineCount = 0 Then
                    MsgBox("It seems that the file that you chose is empty. Please choose a different one", , "Empty file selected")
                    Exit Sub
                End If
                objReader1.Close()

                ProgressBar1.Maximum = (lineCount * 2) 'set the progress bar to twice the size of dictionary file because it runs once in sorting and second time in printing out the dictionary
                ProgressBar1.Value = 0


                ReDim Dictionary(0 To (lineCount - 1)) 'change size of array to the number of words in file

                Dim objReader2 As New System.IO.StreamReader(fdlg.FileName)

                Do While objReader2.Peek() <> -1    'fills array with words from file
                    For lineNB = 0 To lineCount - 1
                        Dictionary(lineNB) = LCase(objReader2.ReadLine())
                    Next lineNB
                Loop

                sorting()

                TextBox2.Text = "Dictionary:" & vbNewLine & vbNewLine 'shows word Dictionary before actual words in textbox2

                For x = 0 To lineCount - 1 'shows words in textbox2 

                    If StopApp = True Then 'in case user wants to cancel loading of dictionary
                        Exit Sub
                        RestartAll()
                    End If

                    TextBox2.Text = TextBox2.Text & Dictionary(x) & vbNewLine
                    ProgressBar1.Value = ProgressBar1.Value + 1                                         'increase progress bar value with each word
                    Label3.Text = "Step " & ProgressBar1.Value & " out of " & ProgressBar1.Maximum      'show which step of whole process is at the moment

                    If showProgress = True Then             'if user wants to see the progres with each step, application refreshes after each step
                        Application.DoEvents()

                    ElseIf showProgress = False Then        'to speed up user can switch off refreshing every step, instead it refreshes every 20th step to keep programs responsiveness during loading
                        refresh = refresh + 1

                        If refresh >= 20 Then
                            Application.DoEvents()
                            refresh = 0
                        End If

                    End If

                Next x

                ProgressBar1.Value = ProgressBar1.Maximum   'when sorting and printing is done, progress bar reach maximum
                Label3.Text = "Step " & ProgressBar1.Value & " out of " & ProgressBar1.Maximum


                objReader2.Close()

                Label1.Visible = True               '\
                Button1.Visible = False             ' \
                TextBox1.Visible = True             '  \
                Button2.Visible = True              '   \ 
                Button4.Visible = True              '    } hides loading buttons and shows all necessary buttons for word checking
                TextBox4.Visible = True             '   /
                ButtonProgress.Visible = False      '  /
                Label3.Visible = False              ' /
                ButtonNew.Visible = True            '/


                'As Timer function is a bit too complicated, I've put this function here to hide the Progress bar completely
                If MessageBox.Show("Dictionary has been loaded and sorted", "Done", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                    Me.Size = New System.Drawing.Size(448, 346)
                Else
                    RestartAll()
                End If
               

            Else
                MsgBox("File Does Not Exist", , "Error")
            End If

        End If


    End Sub


    Private Sub sorting()       'sorts array so binary search will be possible

        Dim x As Integer
        Dim y As Integer
        Dim swap1 As String
        Dim swap2 As String
        Dim refresh As Integer = 0



        For x = 0 To lineCount - 2
            For y = 0 To lineCount - 2
                If StrComp(Dictionary(y), Dictionary(y + 1), vbTextCompare) = 1 Then    'using string compare to swap sort the array
                    swap1 = Dictionary(y)
                    swap2 = Dictionary(y + 1)
                    Dictionary(y + 1) = swap1
                    Dictionary(y) = swap2
                End If
            Next y

            If StopApp = True Then 'in case user wants to cancel loading of dictionary
                Exit Sub
                RestartAll()
            End If

            ProgressBar1.Value = ProgressBar1.Value + 1                                         'increase progress bar value with each word
            Label3.Text = "Step " & ProgressBar1.Value & " out of " & ProgressBar1.Maximum      'show which step of whole process is at the moment

            If showProgress = True Then     'if user wants to see the progres with each step, application refreshes after each step
                Application.DoEvents()

            ElseIf showProgress = False Then    'to speed up user can switch off refreshing every step, instead it refreshes every 20th step to keep programs responsiveness during loading
                refresh = refresh + 1

                If refresh = 20 Then
                    Application.DoEvents()
                    refresh = 0
                End If

            End If

        Next x

    End Sub


    Public Sub Progress(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonProgress.Click 'to show or hide the steps of loading dictionary

        If ButtonProgress.Text = "Hide Steps" Then
            showProgress = False
            ButtonProgress.Text = "Show Steps"
            Label3.Visible = False

        ElseIf ButtonProgress.Text = "Show Steps" Then
            showProgress = True
            ButtonProgress.Text = "Hide Steps"
            Label3.Visible = True

        End If

    End Sub


    Private Sub TextBox1_Enter(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        'I was trying everything to make this work, in the end I turned to google
        'this piece of code taken from http://www.dreamincode.net/forums/topic/65794-how-to-make-textbox-responds-to-enter-key/
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then
            wordCheck()
        End If

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        wordCheck()
    End Sub


    Private Sub wordCheck() 'error checking and running a binary search


        If String.IsNullOrWhiteSpace(TextBox1.Text) Then 'if user didn't write anything or just pressed space then show error message
            MsgBox("Please first write a word which you want to be checked", , "Input box is empty")

        ElseIf TextBox1.Text.IndexOf(" ") <> -1 Or TextBox1.Text.IndexOf(".") <> -1 Or TextBox1.Text.IndexOf(",") <> -1 Then 'this check if entered word contain space full stop or comma, which might indicate that user is trying to search multiple words at the same time
            MsgBox("Please write only one word at the time and use only characters a-z", , "Error in text")

        Else

            TextBox3.Text = ""
            TextBox4.Text = ""
            Label2.Visible = False
            Button2.Visible = False
            Button3.Visible = True
            BinarySearch()

        End If


    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'after clicking on find similar words (button3) it prints string options to textbox 3 and 4. Both because user might choose to show dictionary or not

        TextBox3.Text = options
        TextBox4.Text = options

        If showD = False Then
            TextBox4.Visible = True
        End If

        Button3.Visible = False
        Button2.Visible = True

    End Sub


    Public Sub BinarySearch()

        TextBox1.Text = LCase(TextBox1.Text) 'this will change users word in lowercase

        Dim high As Integer
        Dim low As Integer
        Dim difference As Integer
        Dim middle As Integer
        Dim result As Boolean = False

        'first we set borders which will be moving and determinig where to search for the word
        low = LBound(Dictionary)
        high = UBound(Dictionary)


        ' math.round found on http://www.programmersheaven.com/mb/VBNET/274391/274435/re-rounding/
        'because binary search is looking for the value in the middle of two borders, it would never look on borders themselves. So we check them first
        If StrComp(TextBox1.Text, Dictionary(low), vbTextCompare) = 0 Or StrComp(TextBox1.Text, Dictionary(high), vbTextCompare) = 0 Then
            MsgBox("the word " & TextBox1.Text & " is correct", , "Spell checking complete")
            result = True
            Button2.Visible = True
            Button3.Visible = False

        Else
            'until we will have any result, this function will continue to look for middle of two borders
            Do While result <> True

                difference = Math.Round((high - low) / 2) 'this finds the middle value from difference of borders and then
                middle = low + difference                 'add it to low border to get middle value


                If (low + 1) = high Then    'when binary search reach the end (there is nothing in between them any more) it finishes

                    MsgBox("the word " & TextBox1.Text & " is not in dictionary", , "Spell checking complete")
                    result = True
                    Label2.Visible = True

                    'to show three possibilities we want to take three closest possibilities from binary search
                    If low <> LBound(Dictionary) Then
                        'if low border would be the lowest possible then we couldnt use word before it
                        options = Dictionary(low - 1) & vbNewLine & Dictionary(low) & vbNewLine & Dictionary(high)
                    ElseIf low = LBound(Dictionary) Then
                        options = Dictionary(low) & vbNewLine & Dictionary(high) & vbNewLine & Dictionary(high + 1)
                    Else
                        MsgBox("There seems to be something wrong, please start over")
                    End If

                ElseIf StrComp(TextBox1.Text, Dictionary(middle), vbTextCompare) = 0 Then
                    'then we take the users word and compare it with binary result, if it is the same, we found out that it is correct
                    MsgBox("the word " & TextBox1.Text & " is correct", , "Spell checking complete")
                    result = True
                    Button2.Visible = True
                    Button3.Visible = False

                ElseIf StrComp(TextBox1.Text, Dictionary(middle), vbTextCompare) = -1 Then
                    'if users word is supposed to be before the binary searched one, it will change the high border to the binary searched and continue searching
                    high = middle

                ElseIf StrComp(TextBox1.Text, Dictionary(middle), vbTextCompare) = 1 Then
                    'if users word is supposed to be after the binary searched one, it will change the low border to the binary searched and continue searching
                    low = middle

                End If

            Loop

        End If
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'this is to change layot so that user can see dictionary
        TextBox3.Visible = True
        TextBox2.Visible = True
        TextBox4.Visible = False
        Button5.Visible = True
        Button4.Visible = False

        showD = True

    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'This changes the layot so user don't see dictionary
        TextBox3.Visible = False
        TextBox2.Visible = False
        TextBox4.Visible = True
        Button5.Visible = False
        Button4.Visible = True

        showD = False

    End Sub


    Public Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'restart button
        RestartAll()

    End Sub


    Public Sub RestartAll() 'this will restart whole program

        Button1.Visible = True
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Button5.Visible = False
        ButtonNew.Visible = False
        ButtonProgress.Visible = True
        TextBox1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = True
        Label3.Text = "Step 0 out of 0"
        ProgressBar1.Value = 0
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        options = ""

        StopApp = True
        Me.Size = New System.Drawing.Size(448, 413)


    End Sub


    Private Sub ButtonNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNew.Click
        ' restart search, but keep dictionary
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        options = ""

        If showD = False Then
            TextBox4.Visible = True
        End If

        Button3.Visible = False
        Button2.Visible = True
        Label2.Visible = False

    End Sub


    'Folowing code is to show information about what each of controls do when user hover over it

    Private Sub ButtonProgress_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButtonProgress.MouseHover
        If ButtonProgress.Text = "Hide Steps" Then
            ToolTip1.SetToolTip(ButtonProgress, "By hiding programs steps, loading might be faster")
        ElseIf ButtonProgress.Text = "Show Steps" Then
            ToolTip1.SetToolTip(ButtonProgress, "By showing programs steps, loading might be slower")
        End If
    End Sub

    Private Sub TextBox3_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.MouseHover
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            ToolTip1.SetToolTip(TextBox3, "If you spell word incorrectly, here will be three alternatives for your word")
        Else
            ToolTip1.SetToolTip(TextBox3, "Here are three alternatives for your word")
        End If
    End Sub

    Private Sub TextBox4_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.MouseHover
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            ToolTip1.SetToolTip(TextBox4, "If you spell word incorrectly, here will be three alternatives for your word")
        Else
            ToolTip1.SetToolTip(TextBox4, "Here are three alternatives for your word")
        End If
    End Sub

    'Few bits to make the program complete
    'some pople will not keep mouse over button long enough to see tooltip, so to explain for them what to do:
    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        MsgBox("If you will hover with mouse over buttons or text fields" & vbNewLine & "box with information about what is it for will appear", , "Spell Checker - Help")
    End Sub

    'And as propper piece of software should have... "About"
    Private Sub AboutSpellCheckerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutSpellCheckerToolStripMenuItem.Click
        MsgBox("This Spell Checker was created and designed by" & vbNewLine & "Radek Lochman", , "About Spell Checker")
    End Sub

    Private Sub GreyAndWhiteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GreyAndWhiteToolStripMenuItem.Click
        Me.BackgroundImage = WindowsApplication1.My.Resources.Resources.grad
        Button1.BackgroundImage = Nothing
        Button2.BackgroundImage = Nothing
        Button3.BackgroundImage = Nothing
        Button4.BackgroundImage = Nothing
        Button5.BackgroundImage = Nothing
        ButtonNew.BackgroundImage = Nothing
        ButtonProgress.BackgroundImage = Nothing
        Label2.BackColor = Color.Transparent
    End Sub

    Private Sub FancydictionaryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FancydictionaryToolStripMenuItem.Click
        Me.BackgroundImage = WindowsApplication1.My.Resources.Resources.dict3
        Button1.BackgroundImage = Nothing
        Button2.BackgroundImage = Nothing
        Button3.BackgroundImage = Nothing
        Button4.BackgroundImage = Nothing
        Button5.BackgroundImage = Nothing
        ButtonNew.BackgroundImage = Nothing
        ButtonProgress.BackgroundImage = Nothing
        Label2.BackColor = Color.Transparent
    End Sub

    Private Sub BlueAndOrangeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BlueAndOrangeToolStripMenuItem.Click
        Me.BackgroundImage = WindowsApplication1.My.Resources.Resources.gradorange
        Button1.BackgroundImage = WindowsApplication1.My.Resources.Resources.bluebutton
        Button2.BackgroundImage = WindowsApplication1.My.Resources.Resources.bluebutton
        Button3.BackgroundImage = WindowsApplication1.My.Resources.Resources.bluebutton
        Button4.BackgroundImage = WindowsApplication1.My.Resources.Resources.bluebutton
        Button5.BackgroundImage = WindowsApplication1.My.Resources.Resources.bluebutton
        ButtonNew.BackgroundImage = WindowsApplication1.My.Resources.Resources.bluebutton
        ButtonProgress.BackgroundImage = WindowsApplication1.My.Resources.Resources.bluebutton
        Label2.BackColor = Color.FromKnownColor(KnownColor.InactiveCaption)
    End Sub


End Class

