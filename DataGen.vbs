Option Explicit

' add prefix zero for string
Public Function zFill(tStr, num)
    Dim less, i, tmpStr
	tmpStr = tStr
    less = num - len(tStr)
    If less <= 0 Then
        zFill = tmpStr
        Exit Function
    End If
    For i = 1 to less
        tmpStr = "0" & tmpStr
    Next
    zFill = tmpStr
End Function

' random number from numMin to numMax
public Function randInt(numMin, numMax)
    randInt = RandomNumber.Value(numMin, numMax)
End Function

' random number from numMin to numMax with dotDigit of demical number.
public Function randFloat(numMin, numMax, dotDigit)
   Dim intger
   dotDigit = 10 ^ dotDigit
   numMin = cInt(numMin*dotDigit)
   numMax = cInt(numMax*dotDigit)
   intger = RandomNumber.Value(numMin, numMax)
   randFloat = intger / dotDigit
End Function

' generate a random string in alphabet
public Function randAlpha(strLength)
    Dim i, str
    For i = 1 To strLength Step 1
        If RandomNumber.Value(0,1) = 0 Then
            str = str & chr(RandomNumber.Value(65,90))
        Else
            str = str & chr(RandomNumber.Value(97,122))
        End If     
    Next
    randAlpha = str
End Function

' generate unique string by date&time
public Function randUnique()
    Dim reg
    Set reg = new RegExp
    reg.global = True
    reg.pattern = "\D+"
    randUnique = reg.replace(Now(),"")
End Function

' randomSelect from a WebList/RadioGroup.return the selected value
Public Function randomSelect(obj)
   Dim ItemCount, RandIndex
   ItemCount = obj.GetROProperty("items count")
   RandIndex = RandomNumber.Value(0, ItemCount - 1)
   obj.Select "#" & cStr(RandIndex)
   randomSelect = obj.getROproperty("value")
End Function

' randomSelect
RegisterUserFunc "WebList", "randomSelect", "randomSelect"
RegisterUserFunc "WebRadioGroup", "randomSelect", "randomSelect"
