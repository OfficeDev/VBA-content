---
title: Option Explicit Statement
keywords: vblr6.chm1008808
f1_keywords:
- vblr6.chm1008808
ms.prod: office
ms.assetid: f7871e28-9577-740b-e887-1109f52be30e
ms.date: 06/08/2017
---


# Option Explicit Statement

Used at [module level](vbe-glossary.md) to force explicit declaration of all[variables](vbe-glossary.md) in that[module](vbe-glossary.md).

 **Syntax**

 **Option Explicit**

 **Remarks**
If used, the  **Option** **Explicit** statement must appear in a module before any[procedures](vbe-glossary.md).
When  **Option Explicit** appears in a module, you must explicitly declare all variables using the **Dim**, **Private**, **Public**, **ReDim**, or **Static** statements. If you attempt to use an undeclared variable name, an error occurs at[compile time](vbe-glossary.md).
If you don't use the  **Option Explicit** statement, all undeclared variables are of **Variant** type unless the default type is otherwise specified with a **Def**_type_ statement.

 **Note**  Use  **Option Explicit** to avoid incorrectly typing the name of an existing variable or to avoid confusion in code where the[scope](vbe-glossary.md) of the variable is not clear.


## Example

This example uses the  **Option Explicit** statement to force explicit declaration of all variables. Attempting to use an undeclared variable causes an error at compile time. The **Option Explicit** statement is used at the module level only.


```vb
Option Explicit ' Force explicit variable declaration. 
Dim MyVar ' Declare variable. 
MyInt = 10 ' Undeclared variable generates error. 
MyVar = 10 ' Declared variable does not generate error. 

```


