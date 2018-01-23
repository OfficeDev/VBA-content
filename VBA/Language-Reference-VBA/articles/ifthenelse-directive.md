---
title: If...Then...Else Directive
keywords: vblr6.chm1014461
f1_keywords:
- vblr6.chm1014461
ms.prod: office
ms.assetid: cdda62a6-f9e4-237e-c8b7-a2076e16ff7d
ms.date: 06/08/2017
---


# #If...Then...#Else Directive

Conditionally compiles selected blocks of Visual Basic code.

 **Syntax**

 **#If** _expression_ **Then**  
 _statements_  
[ **#ElseIf** _expression-n_ **Then**  
[ _elseifstatements_ ]]  
[ **#Else**  
[ _elsestatements_ ]]  
 **#End If**

The  **#If...Then...#Else** directive syntax has these parts:

|**Part**|**Description**|
|:-----|:-----|
| _expression_|Required. Any [expression](vbe-glossary.md), consisting exclusively of one or more [conditional compiler constants](vbe-glossary.md), literals, and operators, that evaluates to  **True** or **False**.|
| _statements_|Required. Visual Basic program lines or compiler directives that are evaluated if the associated expression is  **True**.|
| _expression-n_|Optional. Any expression, consisting exclusively of one or more conditional compiler constants, literals, and operators, that evaluates to  **True** or **False**.|
| _elseifstatements_|Optional. One or more program lines or compiler directives that are evaluated if  _expression-n_ is **True**.|
| _elsestatements_|Optional. One or more program lines or compiler directives that are evaluated if no previous  _expression_ or _expression-n_ is **True**.|
 **Remarks**
The behavior of the  **#If...Then...#Else** directive is the same as the **If...Then...Else** statement, except that there is no single-line form of the **#If**, **#Else**, **#ElseIf**, and **#End If** directives; that is, no other code can appear on the same line as any of the directives. Conditional compilation is typically used to compile the same program for different platforms. It is also used to prevent debugging code from appearing in an executable file. Code excluded during conditional compilation is completely omitted from the final executable file, so it has no size or performance effect.
Regardless of the outcome of any evaluation, all expressions are evaluated. Therefore, all [constants](vbe-glossary.md) used in expressions must be defined — any undefined constant evaluates as[Empty](vbe-glossary.md).

 **Note**  The  **Option Compare** statement does not affect expressions in **#If** and **#ElseIf** statements. Expressions in a conditional-compiler directive are always evaluated with **Option Compare Text**.


## Example

This example references conditional compiler constants in an  **#If...Then...#Else** construct to determine whether to compile certain statements.


```vb
' If Mac evaluates as true, do the statements following the #If. 
#If Mac Then 
 '. Place exclusively Mac statements here. 
 '. 
 '. 
' Otherwise, if it is a 32-bit Windows program, do this: 
#ElseIf Win32 Then 
 '. Place exclusively 32-bit Windows statements here. 
 '. 
 '. 
' Otherwise, if it is neither, do this: 
#Else 
 '. Place other platform statements here. 
 '. 
 '. 
#End If
```


