' Program:      Simple Calculator
' Author:       Mark Bulmer
' Date:         March 2, 2022
' Purpose:      This application adds, subtracts, divides, and/or multiplies two
'               operands determined by user input.

Option Strict On

Public Class frmCalculator
    Dim currTextBox As TextBox
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtOperand1.Focus() ' Focuses the cursor on the first operand field.

    End Sub

    Private Sub txtTop_Enter(sender As Object, e As EventArgs) Handles txtOperand1.Enter
        currTextBox = txtOperand1 ' Number buttons correspond to first operand.
    End Sub

    Private Sub txtBottom_Enter(sender As Object, e As EventArgs) Handles txtOperand2.Enter
        currTextBox = txtOperand2 ' Number buttons correspond to second operand.
    End Sub
    Private Sub btnNumber1_Click(sender As Object, e As EventArgs) Handles btnNumber1.Click
        currTextBox.Text = currTextBox.Text & "1" ' Allows user to enter the number one.
    End Sub
    Private Sub btnNumber2_Click(sender As Object, e As EventArgs) Handles btnNumber2.Click
        currTextBox.Text = currTextBox.Text & "2" ' Allows user to enter the number two.
    End Sub

    Private Sub btnNumber3_Click(sender As Object, e As EventArgs) Handles btnNumber3.Click
        currTextBox.Text = currTextBox.Text & "3" ' Allows user to enter the number three.
    End Sub
    Private Sub btnNumber4_Click(sender As Object, e As EventArgs) Handles btnNumber4.Click
        currTextBox.Text = currTextBox.Text & "4" ' Allows user to enter the number four.
    End Sub
    Private Sub btnNumber5_Click(sender As Object, e As EventArgs) Handles btnNumber5.Click
        currTextBox.Text = currTextBox.Text & "5" ' Allows user to enter the number five.
    End Sub
    Private Sub btnNumber6_Click(sender As Object, e As EventArgs) Handles btnNumber6.Click
        currTextBox.Text = currTextBox.Text & "6" ' Allows user to enter the number six.
    End Sub
    Private Sub btnNumber7_Click(sender As Object, e As EventArgs) Handles btnNumber7.Click
        currTextBox.Text = currTextBox.Text & "7" ' Allows user to enter the number seven.
    End Sub
    Private Sub btnNumber8_Click(sender As Object, e As EventArgs) Handles btnNumber8.Click
        currTextBox.Text = currTextBox.Text & "8" ' Allows user to enter the number eight.
    End Sub
    Private Sub btnNumber9_Click(sender As Object, e As EventArgs) Handles btnNumber9.Click
        currTextBox.Text = currTextBox.Text & "9" ' Allows user to enter the number nine.
    End Sub
    Private Sub btnNumber0_Click(sender As Object, e As EventArgs) Handles btnNumber0.Click
        currTextBox.Text = currTextBox.Text & "0" ' Allows user to enter the number zero.
    End Sub
    Private Sub btnPoint_Click(sender As Object, e As EventArgs) Handles btnPoint.Click
        currTextBox.Text = currTextBox.Text & "." ' Allows user to enter a decimal point.
    End Sub
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' This event handler executes the addition operation between the two operands. It executes
        ' the desired operation and then displays the result.

        Dim Operand1 As Double
        Dim Operand2 As Double
        Dim Solution As Double

        Operand1 = Double.Parse(txtOperand1.Text)
        Operand2 = Double.Parse(txtOperand2.Text)
        Solution = Operand1 + Operand2

        lblSolution.Text = Operand1.ToString() & " + " & Operand2.ToString() & " = " & Solution.ToString()

    End Sub
    Private Sub btnSubtract_Click(sender As Object, e As EventArgs) Handles btnSubtract.Click
        ' This event handler executes the subtraction operation between the two operands. It executes
        ' the desired operation and then displays the result.

        Dim Operand1 As Double
        Dim Operand2 As Double
        Dim Solution As Double

        Operand1 = Double.Parse(txtOperand1.Text)
        Operand2 = Double.Parse(txtOperand2.Text)
        Solution = Operand1 - Operand2

        lblSolution.Text = Operand1.ToString() & " - " & Operand2.ToString() & " = " & Solution.ToString()
    End Sub
    Private Sub btnTimes_Click(sender As Object, e As EventArgs) Handles btnTimes.Click
        ' This event handler executes the multiplication operation between the two operands. It executes
        ' the desired operation and then displays the result.

        Dim Operand1 As Double
        Dim Operand2 As Double
        Dim Solution As Double

        Operand1 = Double.Parse(txtOperand1.Text)
        Operand2 = Double.Parse(txtOperand2.Text)
        Solution = Operand1 * Operand2

        lblSolution.Text = Operand1.ToString() & " * " & Operand2.ToString() & " = " & Solution.ToString()
    End Sub
    Private Sub btnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        ' This event handler executes the division operation between the two operands. It executes
        ' the desired operation and then displays the result.

        Dim Operand1 As Double
        Dim Operand2 As Double
        Dim Solution As Double

        Operand1 = Double.Parse(txtOperand1.Text)
        Operand2 = Double.Parse(txtOperand2.Text)
        Solution = Operand1 / Operand2

        lblSolution.Text = Operand1.ToString() & " / " & Operand2.ToString() & " = " & Solution.ToString()
    End Sub
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' This event handler is executed when the user clicks the Clear button. It 
        ' clears the solution label and both operands.

        txtOperand1.Clear()
        txtOperand2.Clear()
        lblSolution.Text = ""
        txtOperand1.Focus()
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExit.Click
        ' Close the window and terminate the application.

        Close()
    End Sub
End Class
