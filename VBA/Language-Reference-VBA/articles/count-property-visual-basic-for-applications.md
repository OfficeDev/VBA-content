---
title: Count Property (Visual Basic for Applications)
keywords: vblr6.chm1014018
f1_keywords:
- vblr6.chm1014018
ms.prod: office
ms.assetid: 319907c0-9c68-0e24-cc76-c16d4386269e
ms.date: 06/08/2017
---


# Count Property (Visual Basic for Applications)



Returns a [Long](vbe-glossary.md) (long integer) containing the number of objects in a[collection](vbe-glossary.md). Read-only.

## Example

This example uses the  **Collection** object's **Count** property to specify how many iterations are required to remove all the elements of the collection called `MyClasses`. When collections are numerically indexed, the base is 1 by default. Since collections are reindexed automatically when a removal is made, the following code removes the first member on each iteration.


```vb
Dim Num, MyClasses
For Num = 1 To MyClasses. Count    ' Remove name from the collection.
    MyClasses.Remove 1    ' Default collection numeric indexes
Next    ' begin at 1.
```


