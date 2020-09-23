<div align="center">

## Capitalise First Letters


</div>

### Description

Takes an input string iof words and changes first letter of each word to a Capital letter
 
### More Info
 
cut and paste the code into a new project.

clicking the convert button 'capitalises' the words in the first text box to the second text box.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Pat Dolan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pat-dolan.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pat-dolan-capitalise-first-letters__1-6523/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub btnConvert_Click()
  Text2.Text = toCapitals(Text1.Text)
End Sub
Private Sub Form_Load()
Text1 = "the cat in the hat works in the c.i.a."
Text2 = ""
End Sub
Function toCapitals(strLowerCase)
  Dim ii, jj
  '--- determine how long the string to be converted is
  ii = Len(strLowerCase)
  '--- first letter of string will always be capitalised
  toCapitals = UCase(Mid(strLowerCase, 1, 1))
  '--- Check the rest of the unconverted string
  '--- We capitalise the next letter whenever we find a space or a break
  For jj = 1 To ii - 1
    If Mid(strLowerCase, jj, 1) = " " Or Mid(strLowerCase, jj, 1) = "." Then
      toCapitals = toCapitals & UCase(Mid(strLowerCase, jj + 1, 1))
    Else
      toCapitals = toCapitals & Mid(strLowerCase, jj + 1, 1)
    End If
  Next
End Function
```

