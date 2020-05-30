vMsg = "Take your pick:" & vbcr & "L for loan, D for deposite, S for saving, P for payments"
vPick = UCase(inputbox(vMsg,, "L"))
vAPR = inputbox("APR",,0.05)
vYrs = inputbox("Number of years",,30)
select case vPick
        case "L"
                vLoan = inputbox("the amonut to borrow",,60000)
		msgbox "Monthly payments for " & vYrs & " years: " & formatcurrency(Loan(vLoan))
	case "D"
		vDeposit = inputbox("Monthly Deposit Amount",,500)
		msgbox "Compounded after " & vYrs & " years: " & formatcurrency(Monthly(vDeposit))
	case "S"
		vGoal = inputbox("whst is your goal after " & vYrs & " yrs.",,100000)
		Do
			vTryDeposit = vTryDeposit + S
			vTryGoal = Monthly(vTryDeposit)
		Loop while CCur(vTryGoal) < CCur(vGoal)      'CInt, CLng, CSng, CDbl
		msgbox "For your goal of " & vGoal & " you must deposit per month: " & formatcurrency(vTryDeposit,0)
	case "P"
		vBudget = Inputbox("what is your monthly budget?",,100)
		Do
			vTryLoan = vTryLoan + 5
			vTryMonthly = Loan(vTryLoan)
		Loop while CCur(vTryMonthly) < CCur(vBudget)
		msgbox "with your monthly budget of " & vBudget & " you can borrow: " & formatcurrency(vTryLoan,0)
	case else: msgbox "You did not choose L, D, S, or p"
end select
Function Monthly(InsideDeposit)
	for i = 1 to (vYrs*12)
		vInterim = vInterim * (1+vAPR/12) + InsideDeposit
	next
	Monthly = vInterim
end function
Function Loan(InsideLoan)
	vRate = (vAPR/12 + 1)^(vYrs*12)
	vCalc = (InsideLoan * vRate) / (vRate - 1) * (vAPR/12)
	Loan = vCalc
End Function
