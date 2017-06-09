---
title: Invalid use of object
keywords: vblr6.chm1011117
f1_keywords:
- vblr6.chm1011117
ms.prod: office
ms.assetid: 63d3d9ba-3521-af29-9484-7c8aa6e65364
ms.date: 06/08/2017
---


# Invalid use of object

You tried to use an object in an incorrect way. This error has the following causes and solutions:


- You tried to discontinue an object reference by assigning  **Nothing** to it but omitted the **Set**[keyword](vbe-glossary.md):
    
  ```
  MyObject = Nothing 
  ```


    Use the  **Set** statement to set an object to **Nothing**. Assuming `MyObject` is an object, you must set it to **Nothing** with the **Set** statement:
    


```vb
Set MyObject = Nothing 
  ```


    Omitting the  **Set** keyword is an implicit use of **Let**, which causes an attempt to perform a value assignment, rather than a reference assignment. **Nothing** can't be used in a value assignment.
    
- You attempted to use Nothing in an [expression](vbe-glossary.md).
    
    Rewrite the expression without the  **Nothing**.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


