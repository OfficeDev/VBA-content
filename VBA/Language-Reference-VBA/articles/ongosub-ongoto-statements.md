---
title: On...GoSub, On...GoTo Statements
keywords: vblr6.chm1008986
f1_keywords:
- vblr6.chm1008986
ms.prod: office
ms.assetid: 9c182e3e-55ba-0d0e-b66c-6ae00189fec5
ms.date: 06/08/2017
---


# On...GoSub, On...GoTo Statements

Branch to one of several specified lines, depending on the value of an [expression](vbe-glossary.md).

 **Syntax**

 **On**_expression_**GoSub**_destinationlist_

 **On**_expression_**GoTo**_destinationlist_
The  **On...GoSub** and **On...GoTo** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _expression_|Required. Any [numeric expression](vbe-glossary.md) that evaluates to a whole number between 0 and 255, inclusive. If _expression_ is any number other than a whole number, it is rounded before it is evaluated.|
| _destinationlist_|Required. List of [line numbers](vbe-glossary.md) or[line labels](vbe-glossary.md) separated by commas.|
 **Remarks**
The value of  _expression_ determines which line is branched to in _destinationlist_. If the value of _expression_ is less than 1 or greater than the number of items in the list, one of the following results occurs:


|**If  _expression_ is**|**Then**|
|:-----|:-----|
|Equal to 0|Control drops to the [statement](vbe-glossary.md) following **On...GoSub** or **On...GoTo**.|
|Greater than number of items in list|Control drops to the statement following  **On...GoSub** or **On...GoTo**.|
|Negative|An error occurs.|
|Greater than 255|An error occurs.|
You can mix line numbers and line labels in the same list. You can use as many line labels and line numbers as you like with  **On...GoSub** and **On...GoTo**. However, if you use more labels or numbers than fit on a single line, you must use the[line-continuation character](vbe-glossary.md) to continue the logical line onto the next physical line.

 **Tip**   **Select Case** provides a more structured and flexible way to perform multiple branching.


## Example

This example uses the  **On...GoSub** and **On...GoTo** statements to branch to subroutines and line labels, respectively.


```vb
Sub OnGosubGotoDemo() 
Dim Number, MyString 
 Number = 2 ' Initialize variable. 
 ' Branch to Sub2. 
 On Number GoSub Sub1, Sub2 ' Execution resumes here after 
 ' On...GoSub. 
 On Number GoTo Line1, Line2 ' Branch to Line2. 
 ' Execution does not resume here after On...GoTo. 
 Exit Sub 
Sub1: 
 MyString = "In Sub1" : Return 
Sub2: 
 MyString = "In Sub2" : Return 
Line1: 
 MyString = "In Line1" 
Line2: 
 MyString = "In Line2" 
End Sub
```


