StartTime = Timer()
Function armstrongFinder()
  num1 = InputBox("Enter first number","Enter Here")
  num2 = InputBox("Enter second number","Enter Here")
  for num = num1 to num2
            temp = cint(num)
            sum=0
            do while temp > 0
                      r = temp mod 10
                      sum = sum + (r*r*r)
                      temp = temp \ 10
            loop
            if(num = sum) then
                      message = message&"The Armstrong is: " &num&vblf
            end if 
  next

  msgbox message
End Function
armstrongFinder()
EndTime = Timer()
Wscript.Echo "Time Taken: " & FormatNumber(EndTime - StartTime, 2) 