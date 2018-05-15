Public Class Form1

    Dim numberofYear As Integer
    Dim iMonthlyPayment, iTotalPayment As String
    Dim InterestRate, monthlyInterestRate, loanAmount, MonthlyPayment, TotalPayment As Double

    Private Sub btnLoan_Click(sender As Object, e As EventArgs) Handles btnLoan.Click
        TabControl1.SelectedTab = TabPage1
        TabPage1.Enabled = True
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False

        btnLoanSystem.Enabled = True
        btnReceiptSystem.Enabled = True
        btnResetSystem.Enabled = True
        btnExit.Enabled = True
    End Sub

    Private Sub btnBalance_Click(sender As Object, e As EventArgs) Handles btnBalance.Click
        TabControl1.SelectedTab = TabPage2
        TabPage2.Enabled = True
        TabPage1.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False
    End Sub

    Private Sub btnDeposit_Click(sender As Object, e As EventArgs) Handles btnDeposit.Click
        TabControl1.SelectedTab = TabPage4
        TabPage4.Enabled = True
        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = False
    End Sub

    Private Sub btnWithdrawal_Click(sender As Object, e As EventArgs) Handles btnWithdrawal.Click
        TabControl1.SelectedTab = TabPage3
        TabPage3.Enabled = True
        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage4.Enabled = False
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btnLoan.Enabled = False
        btnDeposit.Enabled = False
        btnBalance.Enabled = False
        btnWithdrawal.Enabled = False

        btnLoanSystem.Enabled = False
        btnReceiptSystem.Enabled = False
        btnResetSystem.Enabled = False
        btnExit.Enabled = False

        lblLoan.Visible = False
        lblDeposit.Visible = False
        lblBalance.Visible = False
        lblWithdrawal.Visible = False

        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False


    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim iExit As DialogResult
        iExit = MessageBox.Show("Confirm if you want to exit", "ATM System",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If iExit = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim iExit As DialogResult
        iExit = MessageBox.Show("Confirm if you want to exit", "ATM System",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If iExit = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub btnEnter_Click(sender As Object, e As EventArgs) Handles btnEnter.Click

        If lblPin.Text = "1234" Then
            btnLoan.Enabled = True
            btnDeposit.Enabled = True
            btnBalance.Enabled = True
            btnWithdrawal.Enabled = True

            lblLoan.Visible = True
            lblDeposit.Visible = True
            lblBalance.Visible = True
            lblWithdrawal.Visible = True
        Else lblPin.Text = "Invalid Pin"
        End If

    End Sub

    Private Sub num1_Click(sender As Object, e As EventArgs) Handles num1.Click
        lblPin.Text += "1"
    End Sub

    Private Sub num2_Click(sender As Object, e As EventArgs) Handles num2.Click
        lblPin.Text += "2"
    End Sub

    Private Sub num3_Click(sender As Object, e As EventArgs) Handles num3.Click
        lblPin.Text += "3"
    End Sub

    Private Sub num4_Click(sender As Object, e As EventArgs) Handles num4.Click
        lblPin.Text += "4"
    End Sub

    Private Sub num5_Click(sender As Object, e As EventArgs) Handles num5.Click
        lblPin.Text += "5"
    End Sub

    Private Sub num6_Click(sender As Object, e As EventArgs) Handles num6.Click
        lblPin.Text += "6"
    End Sub

    Private Sub num7_Click(sender As Object, e As EventArgs) Handles num7.Click
        lblPin.Text += "7"
    End Sub

    Private Sub num8_Click(sender As Object, e As EventArgs) Handles num8.Click
        lblPin.Text += "8"
    End Sub

    Private Sub num9_Click(sender As Object, e As EventArgs) Handles num9.Click
        lblPin.Text += "9"
    End Sub

    Private Sub btnReceiptSystem_Click(sender As Object, e As EventArgs) Handles btnReceiptSystem.Click
        rtfReceipt.AppendText("Loan Management Systems Calculator" + vbNewLine)
        rtfReceipt.AppendText("----------------------------------------------" + vbNewLine)
        rtfReceipt.AppendText("Enter amount of Loan" + vbTab + txtAmountOfLoan.Text + vbNewLine)
        rtfReceipt.AppendText("Enter Number of Year" + vbTab + txtNumberOfYears.Text + vbNewLine)
        rtfReceipt.AppendText("Enter Interest Rate" + vbTab + vbTab + txtInterestRate.Text + vbNewLine)
        rtfReceipt.AppendText("Monthly Payment" + vbTab + vbTab + lblMonthlyPayment.Text + vbNewLine)
        rtfReceipt.AppendText("Total Payment" + vbTab + vbTab + vbTab + lblTotalPayment.Text + vbNewLine)
        rtfReceipt.AppendText("----------------------------------------------" + vbNewLine)
        rtfReceipt.AppendText("--------------Thank You------------------" + vbNewLine)
    End Sub

    Private Sub btnLoanSystem_Click_1(sender As Object, e As EventArgs) Handles btnLoanSystem.Click

    End Sub

    Private Sub num0_Click(sender As Object, e As EventArgs) Handles num0.Click
        lblPin.Text += "0"
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        lblPin.Text = ""

        btnLoan.Enabled = False
        btnDeposit.Enabled = False
        btnBalance.Enabled = False
        btnWithdrawal.Enabled = False

        btnLoanSystem.Enabled = False
        btnReceiptSystem.Enabled = False
        btnResetSystem.Enabled = False
        btnExit.Enabled = False

        lblLoan.Visible = False
        lblDeposit.Visible = False
        lblBalance.Visible = False
        lblWithdrawal.Visible = False

        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False
    End Sub

    Private Sub btnResetSystem_Click(sender As Object, e As EventArgs) Handles btnResetSystem.Click
        txtAmountOfLoan.Clear()
        txtNumberOfYears.Clear()
        txtInterestRate.Clear()
        lblMonthlyPayment.Clear()
        lblTotalPayment.Clear()
        rtfReceipt.Clear()

        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False


    End Sub

    Private Sub Numbers_Only(sender As Object, e As KeyPressEventArgs) Handles txtNumberOfYears.KeyPress, txtAmountOfLoan.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub btnLoanSystem_Click(sender As Object, e As EventArgs)
        InterestRate = Convert.ToDouble(txtInterestRate.Text)
        monthlyInterestRate = InterestRate / 1200
        numberofYear = Convert.ToInt32(txtNumberOfYears.Text)
        loanAmount = Convert.ToDouble(txtAmountOfLoan.Text)

        MonthlyPayment = loanAmount * monthlyInterestRate / (1 - 1 / Math.Pow(1 + monthlyInterestRate,
            numberofYear * 12))

        iMonthlyPayment = Convert.ToString(MonthlyPayment)
        iMonthlyPayment = FormatCurrency(MonthlyPayment)
        lblMonthlyPayment.Text = (iMonthlyPayment)
        '.....................................................................
        TotalPayment = MonthlyPayment * numberofYear * 12
        iTotalPayment = FormatCurrency(TotalPayment)
        lblTotalPayment.Text = (iTotalPayment)

        txtAmountOfLoan.Text = FormatCurrency(txtAmountOfLoan.Text)
    End Sub
End Class
