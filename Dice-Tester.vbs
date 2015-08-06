Dim chiValue(21, 2)
chiValue(2, 0) = 2.706
chiValue(2, 1) = 6.635
chiValue(3, 0) = 4.605
chiValue(3, 1) = 9.210
chiValue(4, 0) = 6.251
chiValue(4, 1) = 11.341
chiValue(5, 0) = 7.779
chiValue(5, 1) = 13.277
chiValue(6, 0) = 9.236
chiValue(6, 1) = 15.086
chiValue(7, 0) = 10.645
chiValue(7, 1) = 16.812
chiValue(8, 0) = 12.017
chiValue(8, 1) = 18.475
chiValue(9, 0) = 13.362
chiValue(9, 1) = 20.090
chiValue(10, 0) = 14.684
chiValue(10, 1) = 21.666
chiValue(11, 0) = 15.987
chiValue(11, 1) = 23.209
chiValue(12, 0) = 17.275
chiValue(12, 1) = 24.725
chiValue(13, 0) = 18.549
chiValue(13, 1) = 26.217
chiValue(14, 0) = 19.812
chiValue(14, 1) = 27.688
chiValue(15, 0) = 21.064
chiValue(15, 1) = 29.141
chiValue(16, 0) = 22.307
chiValue(16, 1) = 30.578
chiValue(17, 0) = 23.542
chiValue(17, 1) = 32.000
chiValue(18, 0) = 24.769
chiValue(18, 1) = 33.409
chiValue(19, 0) = 25.989
chiValue(19, 1) = 34.805
chiValue(20, 0) = 27.204
chiValue(20, 1) = 36.191

Dim tally(20)
tally(0) = 0
tally(1) = 0
tally(2) = 0
tally(3) = 0
tally(4) = 0
tally(5) = 0
tally(6) = 0
tally(7) = 0
tally(8) = 0
tally(9) = 0
tally(10) = 0
tally(11) = 0
tally(12) = 0
tally(13) = 0
tally(14) = 0
tally(15) = 0
tally(16) = 0
tally(17) = 0
tally(18) = 0
tally(19) = 0

Dim i, numSides, eValue, numRolls, roll, result

Do
    WScript.StdOut.Write("How many sides does the die have? ")
    numSides = WScript.StdIn.ReadLine()

    If IsNumeric(numSides) Then
        numSides = CInt(numSides)
        if numSides > 1 and numSides <= 20 then
            exit do
        end if
    End If
    
    WScript.StdOut.WriteLine("Value must be between 1 and 20")
Loop While true

Do
    WScript.StdOut.Write("What iterative multiplier do you want to use? (default 10) ")
    eValue = WScript.StdIn.ReadLine()
    
    if (eValue = "") then
        eValue = 10
    end if

    If IsNumeric(eValue) Then
        eValue = CInt(eValue)
        if eValue >= 5 then
            exit do
        end if
    End If
    
    WScript.StdOut.WriteLine("Value must be greater than 5")
Loop While true

numRolls = numSides * eValue

WScript.StdOut.WriteLine("You will need to roll your die " & numRolls & " times")
WScript.StdOut.WriteLine("Please enter each number when prompted")

i = 0
do
    WScript.StdOut.Write("Input Roll: ")
    roll = WScript.StdIn.ReadLine()
    
    If IsNumeric(roll) Then
        roll = CInt(roll) - 1
        if roll >= 0 and roll < numSides then
            tally(roll) = tally(roll) + 1
            i = i + 1
        else
            WScript.StdOut.WriteLine("Invalid Roll")
        end if
    else
        WScript.StdOut.WriteLine("Invalid Roll")
    end if
    
loop while i < numRolls

WScript.StdOut.WriteLine()
WScript.StdOut.WriteLine("Totals")

result = 0
i = 0
do
    WScript.StdOut.WriteLine(i + 1 & ": " & tally(i))
    result = result + ((tally(i) - eValue) * (tally(i) - eValue))
    
    i = i + 1
loop while i < numSides

WScript.StdOut.WriteLine()
WScript.StdOut.WriteLine()

result = result / eValue

WScript.StdOut.WriteLine("You Chi-Squared value is " & result)
if result < chiValue(numSides, 0) then
    WScript.StdOut.WriteLine("Which is less than " & chiValue(numSides, 0))
    WScript.StdOut.WriteLine("Your dice appears to be statistically unbiased")
elseif result >= chiValue(numSides, 0) and result <= chiValue(numSides, 1) then
    WScript.StdOut.WriteLine("Which is greater than " & chiValue(numSides, 0) & " and less than " & chiValue(numSides, 1))
    WScript.StdOut.WriteLine("The results were inconclusive please try again with a higher iterative multiplier")
else
    WScript.StdOut.WriteLine("Which is greater than " & chiValue(numSides, 1))
    WScript.StdOut.WriteLine("Your dice appears to be statistically biased")
end if

WScript.Sleep(10*1000)
