---
title: Range has no values
keywords: vblr6.chm1040109
f1_keywords:
- vblr6.chm1040109
ms.prod: office
ms.assetid: 5e74762b-e8b1-cf82-8185-227065bb5f8a
ms.date: 06/08/2017
---


# Range has no values

There are limitations on the way you can specify the number of elements in an [array](vbe-glossary.md). This error has the following cause and solution:



- You specified your array boundaries incorrectly. For example, the following ranges are invalid:
    
```vb
Dim MyArray(10 To -5)    ' Descending order not permitted. 
Dim MyArray(0 To 0)        ' No elements in the array. 

  ```


    Check to be sure your syntax is correct. For example, the following range is valid:
    


```vb
Dim MyArray(-5 To 10)
  ```


For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

