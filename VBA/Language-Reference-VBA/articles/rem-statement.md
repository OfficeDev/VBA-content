---
title: Rem Statement
keywords: vblr6.chm1009000
f1_keywords:
- vblr6.chm1009000
ms.prod: office
ms.assetid: f3a8cec4-dc96-1dee-f428-32b13647aa85
ms.date: 06/08/2017
---


# Rem Statement

Used to include explanatory remarks in a program.

 **Syntax**

 **Rem**_comment_

You can also use the following syntax:
 **'**_comment_
The optional  _comment_[argument](vbe-glossary.md) is the text of any[comment](vbe-glossary.md) you want to include. A space is required between the **Rem**[keyword](vbe-glossary.md) and _comment_.
 **Remarks**
If you use [line numbers](vbe-glossary.md) or[line labels](vbe-glossary.md), you can branch from a  **GoTo** or **GoSub**[statement](vbe-glossary.md) to a line containing a **Rem** statement. Execution continues with the first executable statement following the **Rem** statement. If the **Rem** keyword follows other statements on a line, it must be separated from the statements by a colon ( **:** ).
You can use an apostrophe ( **'** ) instead of the **Rem** keyword. When you use an apostrophe, the colon is not required after other statements.

## Example

This example illustrates the various forms of the  **Rem** statement, which is used to include explanatory remarks in a program.


```vb
Dim MyStr1, MyStr2 
MyStr1 = "Hello": Rem Comment after a statement separated by a colon. 
MyStr2 = "Goodbye" ' This is also a comment; no colon is needed. 

```


