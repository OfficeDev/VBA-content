---
title: Nothing <keyword>
keywords: vblr6.chm1011405
f1_keywords:
- vblr6.chm1011405
ms.prod: office
ms.assetid: 9eedf4db-3aca-df26-8bc7-c3a7f7264e6b
ms.date: 06/08/2017
---


# Nothing <keyword>

The  **Nothing**[keyword](vbe-glossary.md) is used to disassociate an object[variable](vbe-glossary.md) from an actual object. Use the **Set** statement to assign **Nothing** to an object variable. For example:


```vb
Set MyObject = Nothing 

```


Several object variables can refer to the same actual object. When  **Nothing** is assigned to an object variable, that variable no longer refers to an actual object. When several object variables refer to the same object, memory and system resources associated with the object to which the variables refer are released only after all of them have been set to **Nothing**, either explicitly using **Set**, or implicitly after the last object variable set to **Nothing** goes out of[scope](vbe-glossary.md).


