Attribute VB_Name = "Module1"
Sub LoanAmortization()

    ' Declare variables
    Dim principal As Double
    Dim annualInterestRate As Double
    Dim monthlyInterestRate As Double
    Dim loanTerm As Integer
    Dim numPayments As Integer
    Dim monthlyPayment As Double
    Dim balance As Double
    Dim interestPaid As Double
    Dim principalPaid As Double
    Dim paymentNum As Integer
    
    ' Input values from the user
    principal = InputBox("Enter the loan principal (amount borrowed):")
    annualInterestRate = InputBox("Enter the annual interest rate (as a percentage):") / 100
    loanTerm = InputBox("Enter the loan term in years:")
    
    ' Calculate monthly interest rate and number of payments
    monthlyInterestRate = annualInterestRate / 12
    numPayments = loanTerm * 12
    ' Calculate monthly payment using the loan amortization formula
    monthlyPayment = (principal * monthlyInterestRate * (1 + monthlyInterestRate) ^ numPayments) / ((1 + monthlyInterestRate) ^ numPayments - 1)
    
    ' Print amortization schedule in Excel
    Range("A1:E1").Value = Array("Payment #", "Payment", "Interest Paid", "Principal Paid", "Remaining Balance")
    
    balance = principal
    
    ' Loop through each payment
    For paymentNum = 1 To numPayments
        interestPaid = balance * monthlyInterestRate
        principalPaid = monthlyPayment - interestPaid
        balance = balance - principalPaid
        
        ' Output the schedule into Excel
        Cells(paymentNum + 1, 1).Value = paymentNum
        Cells(paymentNum + 1, 2).Value = monthlyPayment
        Cells(paymentNum + 1, 3).Value = interestPaid
        Cells(paymentNum + 1, 4).Value = principalPaid
        Cells(paymentNum + 1, 5).Value = balance
    Next paymentNum

End Sub

Sub CalculateROI()

    ' Declare variables
    Dim initialInvestment As Double
    Dim totalReturn As Double
    Dim ROI As Double
    
    ' Input values from the user
    initialInvestment = InputBox("Enter the initial investment amount:")
    totalReturn = InputBox("Enter the total return on the investment:")
    
    ' Calculate ROI
    ROI = ((totalReturn - initialInvestment) / initialInvestment) * 100
    
    ' Output the result to the Excel sheet
    Range("A1").Value = "Initial Investment"
    Range("B1").Value = initialInvestment
    Range("A2").Value = "Total Return"
    Range("B2").Value = totalReturn
    Range("A3").Value = "Return on Investment (ROI)"
    Range("B3").Value = ROI & "%"
    
End Sub

Sub InvestmentGrowth()

    ' Declare variables
    Dim initialInvestment As Double
    Dim annualInterestRate As Double
    Dim numYears As Integer
    Dim futureValue As Double
    Dim yearNum As Integer
    
    ' Input values from the user
    initialInvestment = InputBox("Enter the initial investment amount:")
    annualInterestRate = InputBox("Enter the annual rate of return (as a percentage):") / 100
    numYears = InputBox("Enter the number of years:")
    
    ' Print the headers in Excel
    Range("A1:C1").Value = Array("Year", "Investment Growth", "Future Value")
    
    ' Loop through each year to calculate future value
    For yearNum = 1 To numYears
        futureValue = initialInvestment * (1 + annualInterestRate) ^ yearNum
        
        ' Output the year and future value into Excel
        Cells(yearNum + 1, 1).Value = yearNum
        Cells(yearNum + 1, 2).Value = initialInvestment * (1 + annualInterestRate) ^ (yearNum - 1)
        Cells(yearNum + 1, 3).Value = futureValue
    Next yearNum
    
End Sub


