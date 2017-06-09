---
title: Invalid Next control variable reference
keywords: vblr6.chm1011226
f1_keywords:
- vblr6.chm1011226
ms.prod: office
ms.assetid: 1fd6eeda-b1e9-5c36-8100-b0e8ea3614fc
ms.date: 06/08/2017
---


# Invalid Next control variable reference

The numeric [variable](vbe-glossary.md) in the **Next** part of a **For...Next** loop must match the variable in the **For** part. This error has the following cause and solution:



- The variable in the  **Next** part of a **For...Next** loop differs from the variable in the **For** part. For example:
    
```vb
For Counter = 1 To 10 
MyVar = Counter 
Next Count 

  ```


    Check the spelling of the variable in the  **Next** part to be sure it matches the **For** part. Also, be sure you haven't inadvertently deleted parts of the enclosing loop that used the variable.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

