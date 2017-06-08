---
title: For Each...Next Statement
keywords: vblr6.chm1009275
f1_keywords:
- vblr6.chm1009275
ms.prod: office
ms.assetid: bbff57d3-3655-3426-02a1-ae6748736fb1
ms.date: 06/08/2017
---


# For Each...Next Statement

Repeats a group of [statements](vbe-glossary.md) for each element in an[array](vbe-glossary.md) or[collection](vbe-glossary.md).

 **Syntax**

 **For** **Each**_element_**In**_group_
 [ _statements_ ]
 [ **Exit For** ]
 [ _statements_ ]

 **Next** [ _element_ ]
The  **For...Each...Next** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _element_|Required. [Variable](vbe-glossary.md) used to iterate through the elements of the collection or array. For collections, _element_ can only be a[Variant](vbe-glossary.md) variable, a generic object variable, or any specific object variable. For arrays, _element_ can only be a **Variant** variable.|
| _group_|Required. Name of an object collection or array (except an array of [user-defined types](vbe-glossary.md)).|
| _statements_|Optional. One or more statements that are executed on each item in  _group_.|
 **Remarks**
The  **For…Each** block is entered if there is at least one element in _group_. Once the loop has been entered, all the statements in the loop are executed for the first element in _group_. If there are more elements in _group_, the statements in the loop continue to execute for each element. When there are no more elements in _group_, the loop is exited and execution continues with the statement following the **Next** statement.
Any number of  **Exit For** statements may be placed anywhere in the loop as an alternative way to exit. **Exit For** is often used after evaluating some condition, for example **If…Then**, and transfers control to the statement immediately following **Next**.
You can nest  **For...Each...Next** loops by placing one **For…Each…Next** loop within another. However, each loop _element_ must be unique.

 **Note**  If you omit  _element_ in a **Next** statement, execution continues as if _element_ is included. If a **Next** statement is encountered before its corresponding **For** statement, an error occurs.

You can't use the  **For...Each...Next** statement with an array of user-defined types because a **Variant** can't contain a user-defined type.

## Example

This example uses the  **For Each...Next** statement to search the **Text** property of all elements in a collection for the existence of the string "Hello". In the example, `MyObject` is a text-related object and is an element of the collection `MyCollection`. Both are generic names used for illustration purposes only.


```vb
Dim Found, MyObject, MyCollection 
Found = False    ' Initialize variable. 
For Each MyObject In MyCollection    ' Iterate through each element.  
    If MyObject.Text = "Hello" Then    ' If Text equals "Hello". 
        Found = True    ' Set Found to True. 
        Exit For    ' Exit loop. 
    End If 
Next
```


