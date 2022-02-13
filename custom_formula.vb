
Function Blackscholes(K As Double, t As Double, S As Double, r As Double, sigma As Double)
Dim d1 As Double
Dim d2 As Double

d1 = (WorksheetFunction.Ln(S / K) + (r + sigma ^ 2 / 2) * t) / (sigma * t ^ 0.5)
d2 = d1 - sigma * Sqr(t)

Blackscholes = S * WorksheetFunction.NormSDist(d1) - K * Exp(-r * t) * WorksheetFunction.NormSDist(d2)
End Function


Function americanput(n As Long, K As Double, t As Double, S As Double, r As Double, sigma As Double)

Dim u As Double
Dim d As Double
Dim deltaT As Double
Dim j As Integer
Dim putprice() As Double
Dim stockprice() As Double
Dim q As Double

ReDim putprice(n, n)
ReDim stockprice(n, n)


deltaT = t / n

u = Exp(sigma * Sqr(deltaT))
d = Exp(-sigma * Sqr(deltaT))

q = (Exp(r * deltaT) - d) / (u - d)



For j = 0 To n                'payoff at END

putprice(j, n) = WorksheetFunction.Max(K - S * u ^ (n - j) * d ^ j, 0) 'put prices at time N
                          
Next j



For j = 0 To n              'stock price at (i,j)

    For i = 0 To j
    
        stockprice(i, j) = S * u ^ (j - i) * d ^ i
                
    Next i

Next j




For j = n - 1 To 0 Step -1          'payoff's at (i,j)

    For i = 0 To j
    
    putprice(i, j) = WorksheetFunction.Max(Exp(-r * deltaT) * (q * putprice(i, j + 1) + (1 - q) * putprice(i + 1, j + 1)), K - stockprice(i, j))
    
    Next i

Next j


    
americanput = putprice(0, 0)


End Function


Function european(n As Long, K As Double, t As Double, S As Double, r As Double, sigma As Double)


Dim u As Double
Dim d As Double
Dim deltaT As Double
Dim j As Integer
Dim q As Double
Dim pay As Double
Dim cal As Double
Dim payoff() As Double
Dim callprice() As Double

ReDim payoff(n)
ReDim callprice(n)


deltaT = t / n

u = Exp(sigma * Sqr(deltaT))
d = Exp(-sigma * Sqr(deltaT))

q = (Exp(r * deltaT) - d) / (u - d)


For j = 0 To n

    payoff(j) = WorksheetFunction.Max(S * u ^ (n - j) * d ^ j - K, 0)

    callprice(j) = Exp(-r * t) * WorksheetFunction.Combin(n, j) * q ^ (n - j) * (1 - q) ^ j * payoff(j)

Next j

european = WorksheetFunction.Sum(callprice)

End Function


