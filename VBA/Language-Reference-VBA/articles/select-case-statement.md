---
title: Select Case Statement
keywords: vblr6.chm1008810
f1_keywords:
- vblr6.chm1008810
ms.prod: office
ms.assetid: 8e885f14-c722-5217-705e-474516fa416b
ms.date: 06/08/2017
---


# Select Case Statement

Executes one of several groups of [statements](vbe-glossary.md), depending on the value of an [expression](vbe-glossary.md).

 **Syntax**

 **Select Case** _testexpression_
 [ **Case** _expressionlist-n_
 [ _statements-n_ ]]
 [ **Case Else**
 [ _elsestatements_ ]]

 **End Select**
 
The  **Select Case** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _testexpression_|Required. Any [numeric expression](vbe-glossary.md) or[string expression](vbe-glossary.md).|
| _expressionlist-n_|Required if a  **Case** appears. Delimited list of one or more of the following forms: _expression_, _expression_ **To** _expression_, **Is** _comparisonoperator_ _expression_. The **To** [keyword](vbe-glossary.md) specifies a range of values. If you use the **To** keyword, the smaller value must appear before **To**. Use the **Is** keyword with [comparison operators](vbe-glossary.md) (except **Is** and **Like**) to specify a range of values. If not supplied, the **Is** keyword is automatically inserted.|
| _statements-n_|Optional. One or more statements executed if  _testexpression_ matches any part of _expressionlist-n._|
| _elsestatements_|Optional. One or more statements executed if  _testexpression_ doesn't match any of the **Case** clause.|

**Remarks**

If  _testexpression_ matches any **Case** _expressionlist_ expression, the _statements_ following that **Case** clause are executed up to the next **Case** clause, or, for the last clause, up to **End Select**. Control then passes to the statement following **End Select**. If _testexpression_ matches an _expressionlist_ expression in more than one **Case** clause, only the statements following the first match are executed.

The  **Case Else** clause is used to indicate the _elsestatements_ to be executed if no match is found between the _testexpression_ and an _expressionlist_ in any of the other **Case** selections. Although not required, it is a good idea to have a **Case Else** statement in your **Select Case** block to handle unforeseen _testexpression_ values. If no **Case** _expressionlist_ matches _testexpression_ and there is no **Case Else** statement, execution continues at the statement following **End Select**.

You can use multiple expressions or ranges in each  **Case** clause. For example, the following line is valid:



```
Case 1 To 4, 7 To 9, 11, 13, Is > MaxNumber 

```


 **Note**  The  **Is** comparison operator is not the same as the **Is** keyword used in the **Select Case** statement.

You also can specify ranges and multiple expressions for character strings. In the following example,  **Case** matches strings that are exactly equal to `everything` , strings that fall between `nuts` and `soup` in alphabetic order, and the current value of `TestItem` :



```
Case "everything", "nuts" To "soup", TestItem 

```

 **Select Case** statements can be nested. Each nested **Select Case** statement must have a matching **End Select** statement.

## Example

This example uses the  **Select Case** statement to evaluate the value of a variable. The second **Case** clause contains the value of the variable being evaluated, and therefore only the statement associated with it is executed.


```vb
Dim Number 
Number = 8    ' Initialize variable. 
Select Case Number    ' Evaluate Number. 
Case 1 To 5    ' Number between 1 and 5, inclusive. 
    Debug.Print "Between 1 and 5" 
' The following is the only Case clause that evaluates to True. 
Case 6, 7, 8    ' Number between 6 and 8. 
    Debug.Print "Between 6 and 8" 
Case 9 To 10    ' Number is 9 or 10. 
    Debug.Print "Greater than 8" 
Case Else    ' Other values. 
    Debug.Print "Not between 1 and 10" 
End Select
```


