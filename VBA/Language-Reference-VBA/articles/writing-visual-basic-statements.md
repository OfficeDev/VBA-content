---
title: Writing Visual Basic Statements
keywords: vbcn6.chm1076695
f1_keywords:
- vbcn6.chm1076695
ms.prod: office
ms.assetid: a2d35638-995b-1a6f-2975-8deacddf93de
ms.date: 06/08/2017
---


# Writing Visual Basic Statements

A [statement](vbe-glossary.md) in Visual Basic is a complete instruction. It can contain[keywords](vbe-glossary.md), operators, [variables](vbe-glossary.md), [constants](vbe-glossary.md), and [expressions](vbe-glossary.md). Each statement belongs to one of the following three categories:



- Declaration statements, which name a variable, constant, or procedure and can also specify a data type.[Writing Declaration Statements](writing-declaration-statements.md)
    
- Assignment statements, which assign a value or expression to a variable or constant.[Writing Assignment Statements](writing-assignment-statements.md)
    
- Executable statements, which initiate actions. These statements can execute a method or function, and they can loop or branch through blocks of code. Executable statements often contain mathematical or conditional operators.[Writing Executable Statements](writing-executable-statements.md)
    


## Continuing a Statement over Multiple Lines

A statement usually fits on one line, but you can continue a statement onto the next line using a [line-continuation character](vbe-glossary.md). In the following example, the  **MsgBox** executable statement is continued over three lines:


```vb
Sub DemoBox() 'This procedure declares a string variable, 
 ' assigns it the value Claudia, and then displays 
 ' a concatenated message. 
 Dim myVar As String 
 myVar = "John" 
 MsgBox Prompt:="Hello " &; myVar, _ 
 Title:="Greeting Box", _ 
 Buttons:=vbExclamation 
End Sub
```


## Adding Comments

Comments can explain a procedure or a particular instruction to anyone reading your code. Visual Basic ignores comments when it runs your procedures. Comment lines begin with an apostrophe ( **'** ) or with **Rem** followed by a space, and can be added anywhere in a procedure. To add a comment to the same line as a statement, insert an apostrophe after the statement, followed by the comment. By default, comments are displayed as green text.


## Checking Syntax Errors

If you press ENTER after typing a line of code and the line is displayed in red (an error message may display as well), you must find out what's wrong with your statement, and then correct it.


