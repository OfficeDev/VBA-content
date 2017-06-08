---
title: While...Wend Statement
keywords: vblr6.chm1008811
f1_keywords:
- vblr6.chm1008811
ms.prod: office
ms.assetid: c905a6a3-fa70-42df-5ef0-c4e3193c2e10
ms.date: 06/08/2017
---


# While...Wend Statement

Executes a series of [statements](vbe-glossary.md) as long as a given condition is **True**.

 **Syntax**

 **While**_condition_
[ _statements_ ]

 **Wend**
The  **While...Wend** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _condition_|Required. [Numeric expression](vbe-glossary.md) or[string expression](vbe-glossary.md) that evaluates to **True** or **False**. If _condition_ is[Null](vbe-glossary.md),  _condition_ is treated as **False**.|
| _statements_|Optional. One or more statements executed while condition is  **True**.|
 **Remarks**
If  _condition_ is **True**, all _statements_ are executed until the **Wend** statement is encountered. Control then returns to the **While** statement and _condition_ is again checked. If _condition_ is still **True**, the process is repeated. If it is not **True**, execution resumes with the statement following the **Wend** statement.
 **While...Wend** loops can be nested to any level. Each **Wend** matches the most recent **While**.

 **Tip**  The  **Do...Loop** statement provides a more structured and flexible way to perform looping.


## Example

This example uses the  **While...Wend** statement to increment a counter variable. The statements in the loop are executed as long as the condition evaluates to **True**.


```vb
Dim Counter 
Counter = 0 ' Initialize variable. 
While Counter < 20 ' Test value of Counter. 
 Counter = Counter + 1 ' Increment Counter. 
Wend ' End While loop when Counter > 19. 
Debug.Print Counter ' Prints 20 in the Immediate window. 

```


