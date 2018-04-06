---
title: GoTo Statement - VBA
keywords: vblr6.chm1008935
f1_keywords:
- vblr6.chm1008935
ms.prod: office
ms.assetid: 0fa45435-77cf-91f5-ade4-86ac0eb1a083
ms.date: 06/08/2017
---
# GoTo Statement

Branches unconditionally to a specified line within a [procedure](vbe-glossary.md).

 **Syntax**

 **GoTo** _line_

The required _line_ [argument](vbe-glossary.md) can be any [line label](vbe-glossary.md) or [line number](vbe-glossary.md).

 **Remarks**
 
 **GoTo** can branch only to lines within the procedure where it appears.

 >**Note**  Too many  **GoTo** statements can make code difficult to read and debug. Use structured control [statements](vbe-glossary.md) ( **Do...Loop**, **For...Next**, **If...Then...Else**, **Select Case** ) whenever possible.

## Example

This example uses the  **GoTo** statement to branch to line labels within a procedure.


```vb
Sub GotoStatementDemo() 
Dim Number, MyString 
 Number = 1 ' Initialize variable. 
 ' Evaluate Number and branch to appropriate label. 
 If Number = 1 Then GoTo Line1 Else GoTo Line2 
 
Line1: 
 MyString = "Number equals 1" 
 GoTo LastLine ' Go to LastLine. 
Line2: 
 ' The following statement never gets executed. 
 MyString = "Number equals 2" 
LastLine: 
 Debug.Print MyString ' Print "Number equals 1" in 
 ' the Immediate window. 
End Sub
```
