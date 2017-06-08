---
title: Using Parentheses in Code
keywords: vbcn6.chm1076685
f1_keywords:
- vbcn6.chm1076685
ms.prod: office
ms.assetid: 7894f174-ac01-dcc2-a30d-63d5c3625af6
ms.date: 06/08/2017
---


# Using Parentheses in Code

 **Sub** procedures, built-in[statements](vbe-glossary.md), and some [methods](vbe-glossary.md) don't return a value, so the[arguments](vbe-glossary.md) aren't enclosed in parentheses. For example:


```
MySub "stringArgument", integerArgument 

```


 **Function** procedures, built-in functions, and some methods do return a value, but you can ignore it. If you ignore the return value, don't include parentheses. Call the function just as you would call a **Sub** procedure. Omit the parentheses, list any arguments, and don't assign the function to a variable. For example:




```vb
MsgBox "Task Completed!", 0, "Task Box" 

```

To use the return value of a function, enclose the arguments in parentheses, as shown in the following example.



```
Answer3 = MsgBox("Are you happy with your salary?", 4, "Question 3") 

```

A statement in a  **Sub** or **Function** procedure can pass values to a called procedure using[named arguments](vbe-glossary.md). The guidelines for using parentheses apply, whether or not you use named arguments. When you use named arguments, you can list them in any order, and you can omit optional arguments. Named arguments are always followed by a colon and an equal sign ( **:=** ), and then the argument value.
The following example calls the  **MsgBox** function using named arguments, but it ignores the return value:



```vb
MsgBox Title:="Task Box", Prompt:="Task Completed!" 

```

The following example calls the  **MsgBox** function using named arguments and assigns the return value to the variable :



```
answer3 = MsgBox(Title:="Question 3", _ 
 Prompt:="Are you happy with your salary?", Buttons:=4) 

```


