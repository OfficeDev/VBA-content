---
title: Using Select Case Statements
keywords: vbcn6.chm1076686
f1_keywords:
- vbcn6.chm1076686
ms.prod: office
ms.assetid: 0573a361-84d6-549f-8c51-5bc0fe17d156
ms.date: 06/08/2017
---


# Using Select Case Statements

Use the  **Select Case** statement as an alternative to using **ElseIf** in **If...Then...Else** statements when comparing one[expression](vbe-glossary.md) to several different values. While **If...Then...Else** statements can evaluate a different expression for each **ElseIf** statement, the **Select Case** statement evaluates an expression only once, at the top of the control structure.

In the following example, the  **Select Case** statement evaluates the argument that is passed to the procedure. Note that each **Case** statement can contain more than one value, a range of values, or a combination of values and[comparison operators](vbe-glossary.md). The optional  **Case Else** statement runs if the **Select Case** statement doesn't match a value in any of the **Case** statements.



```vb
Function Bonus(performance, salary) 
  Select Case performance 
    Case 1 
      Bonus = salary * 0.1 
    Case 2, 3 
      Bonus = salary * 0.09 
    Case 4 To 6 
      Bonus = salary * 0.07 
    Case Is > 8 
      Bonus = 100 
    Case Else 
      Bonus = 0 
  End Select 
End Function 
```


