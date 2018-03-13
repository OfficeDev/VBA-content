---
title: If...Then...Else Statement
keywords: vblr6.chm1008940
f1_keywords:
- vblr6.chm1008940
ms.prod: office
ms.assetid: 53514f63-ec20-27bf-2b61-5706540a4999
ms.date: 06/08/2017
---


# If...Then...Else Statement

Conditionally executes a group of [statements](vbe-glossary.md), depending on the value of an [expression](vbe-glossary.md).

 **Syntax**

 **If**_condition_**Then** [ _statements_ ] [ **Else**_elsestatements_ ]

Or, you can use the block form syntax:
 **If**_condition_**Then**
[ _statements_ ]
[ **ElseIf**_condition-n_**Then**
[ _elseifstatements_ ]
[ **Else**
[ _elsestatements_ ]]
 **End If**
The  **If...Then...Else** statement syntax has these parts:


| <strong>Part</strong>                                                                                                                                                                                                                                                                                                                                                                           | <strong>Description</strong>                                                                                                                                                                         |
|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>condition</em>                                                                                                                                                                                                                                                                                                                                                                              | Required. One or more of the following two types of expressions:                                                                                                                                     |
|                                                                                                                                                                                                                                                                                                                                                                                                 |                                                                                                                                                                                                      |
| A [numeric expression](vbe-glossary.md) or[string expression](vbe-glossary.md) that evaluates to <strong>True</strong> or <strong>False</strong>. If <em>condition</em> is[Null](vbe-glossary.md),  <em>condition</em> is treated as <strong>False</strong>.                                                                                                                                    |                                                                                                                                                                                                      |
|                                                                                                                                                                                                                                                                                                                                                                                                 |                                                                                                                                                                                                      |
| An expression of the form  <strong>TypeOf</strong><em>objectname</em><strong>Is</strong><em>objecttype</em>. The <em>objectname</em> is any object reference and <em>objecttype</em> is any valid object type. The expression is <strong>True</strong> if <em>objectname</em> is of the[object type](vbe-glossary.md) specified by <em>objecttype</em>; otherwise it is <strong>False</strong>. |                                                                                                                                                                                                      |
| <em>statements</em>                                                                                                                                                                                                                                                                                                                                                                             | Optional in block form; required in single-line form that has no  <strong>Else</strong> clause. One or more statements separated by colons; executed if <em>condition</em> is <strong>True</strong>. |
| <em>condition-n</em>                                                                                                                                                                                                                                                                                                                                                                            | Optional. Same as  <em>condition</em>.                                                                                                                                                               |
| <em>elseifstatements</em>                                                                                                                                                                                                                                                                                                                                                                       | Optional. One or more statements executed if associated  <em>condition-n</em> is <strong>True</strong>.                                                                                              |
| <em>elsestatements</em>                                                                                                                                                                                                                                                                                                                                                                         | Optional. One or more statements executed if no previous  <em>condition</em> or <em>condition-n</em> expression is <strong>True</strong>.                                                            |

 **Remarks**
You can use the single-line form (first syntax) for short, simple tests. However, the block form (second syntax) provides more structure and flexibility than the single-line form and is usually easier to read, maintain, and debug.

 **Note**  With the single-line form, it is possible to have multiple statements executed as the result of an  **If...Then** decision. All statements must be on the same line and separated by colons, as in the following statement:




```
If A > 10 Then A = A + 1 : B = B + A : C = C + B 
```

A block form  **If** statement must be the first statement on a line. The **Else**, **ElseIf**, and **End If** parts of the statement can have only a[line number](vbe-glossary.md) or[line label](vbe-glossary.md) preceding them. The block **If** must end with an **End If** statement.
To determine whether or not a statement is a block  **If**, examine what follows the **Then**[keyword](vbe-glossary.md). If anything other than a [comment](vbe-glossary.md) appears after **Then** on the same line, the statement is treated as a single-line **If** statement.
The  **Else** and **ElseIf** clauses are both optional. You can have as many **ElseIf** clauses as you want in a block **If**, but none can appear after an **Else** clause. Block **If** statements can be nested; that is, contained within one another.
When executing a block  **If** (second syntax), _condition_ is tested. If _condition_ is **True**, the statements following **Then** are executed. If _condition_ is **False**, each **ElseIf** condition (if any) is evaluated in turn. When a **True** condition is found, the statements immediately following the associated **Then** are executed. If none of the **ElseIf** conditions are **True** (or if there are no **ElseIf** clauses), the statements following **Else** are executed. After executing the statements following **Then** or **Else**, execution continues with the statement following **End If**.
 **Tip** **Select Case** may be more useful when evaluating a single expression that has several possible actions. However, the **TypeOf**_objectname_**Is**_objecttype_ clause can't be used with the **Select Case** statement.

 **Note**   **TypeOf** cannot be used with hard data types such as Long, Integer, and so forth other than Object.


## Example

This example shows both the block and single-line forms of the  **If...Then...Else** statement. It also illustrates the use of **If TypeOf...Then...Else**.


```vb
Dim Number, Digits, MyString 
Number = 53 ' Initialize variable. 
If Number < 10 Then 
 Digits = 1 
ElseIf Number < 100 Then 
' Condition evaluates to True so the next statement is executed. 
 Digits = 2 
Else 
 Digits = 3 
End If 

' Assign a value using the single-line form of syntax. 
If Digits = 1 Then MyString = "One" Else MyString = "More than one" 
```

Use  **If TypeOf** construct to determine whether the Control passed into a procedure is a text box.




```vb
Sub ControlProcessor(MyControl As Control) 
 IfTypeOf MyControl Is CommandButton Then 
 Debug.Print "You passed in a " &; TypeName(MyControl) 
 ElseIfTypeOf MyControl Is CheckBox Then 
 Debug.Print "You passed in a " &; TypeName(MyControl) 
 ElseIfTypeOf MyControl Is TextBox Then 
 Debug.Print "You passed in a " &; TypeName(MyControl) 
 End If 
End Sub
```


