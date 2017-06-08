---
title: Is Operator
keywords: vblr6.chm1008950
f1_keywords:
- vblr6.chm1008950
ms.prod: office
ms.assetid: c84836c1-7b21-a659-9d34-3bef8784c5a3
ms.date: 06/08/2017
---


# Is Operator



Used to compare two object reference [variables](vbe-glossary.md).
 **Syntax**
 _result_**=**_object1_**Is**_object2_
The  **Is** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric variable.|
| _object1_|Required; any object name.|
| _object2_|Required; any object name.|
 **Remarks**
If  _object1_ and _object2_ both refer to the same object, _result_ is **True**; if they do not, _result_ is **False**. Two variables can be made to refer to the same object in several ways.
In the following example, A has been set to refer to the same object as B:



```vb
Set A = B

```

The following example makes A and B refer to the same object as C:



```vb
Set A = C
Set B = C


```


## Example

This example uses the  **Is** operator to compare two object references. The object variable names are generic and used for illustration purposes only.


```vb
Dim MyObject, YourObject, ThisObject, OtherObject, ThatObject, MyCheck
Set YourObject = MyObject    ' Assign object references.
Set ThisObject = MyObject
Set ThatObject = OtherObject
MyCheck = YourObject Is ThisObject    ' Returns True.
MyCheck = ThatObject Is ThisObject    ' Returns False.
' Assume MyObject <> OtherObject
MyCheck = MyObject Is ThatObject    ' Returns False.

```


