---
title: Out of string space (Error 14)
keywords: vblr6.chm1011244
f1_keywords:
- vblr6.chm1011244
ms.prod: office
ms.assetid: b400380a-4dda-306e-b086-af201e5f2835
ms.date: 06/08/2017
---


# Out of string space (Error 14)

Visual Basic permits you to use very large strings. However, the requirements of other programs and the way you manipulate your strings may cause this error. This error has the following causes and solutions:



- [Expressions](vbe-glossary.md) requiring that temporary strings be created for evaluation may cause this error. For example, the following code causes an `Out of string space` error on some operating systems:
    
  ```
  MyString = "Hello" 
For Count = 1 To 100 
MyString = MyString &; MyString 
Next Count 

  ```


    Assign the string to a [variable](vbe-glossary.md) of another name.
    
- Your system may have run out of memory, which prevented a string from being allocated. Remove any unnecessary applications from memory to create more space.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

