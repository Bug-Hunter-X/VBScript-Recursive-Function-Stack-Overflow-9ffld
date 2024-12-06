Function factorial(x)
  Dim result, i
  result = 1
  For i = 1 To x
    result = result * i
  Next
  factorial = result
End Function
MsgBox factorial(5) 
'This function will now correctly calculate the factorial without errors.