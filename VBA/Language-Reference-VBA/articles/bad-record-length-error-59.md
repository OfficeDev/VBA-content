---
title: Bad record length (Error 59)
keywords: vblr6.chm1000059
f1_keywords:
- vblr6.chm1000059
ms.prod: office
ms.assetid: 4fd56cfa-a45c-1ac9-5ef0-3ccec2004d48
ms.date: 06/08/2017
---


# Bad record length (Error 59)

The length of a record [variable](vbe-glossary.md) in a **Get** or **Put** statement must be the length specified in its corresponding **Open** statement. This error has the following causes and solutions:



- The record variable's length differs from the length specified in the corresponding  **Open** statement. Make sure the sum of the sizes of fixed-length[variables](vbe-glossary.md) in the[user-defined type](vbe-glossary.md) defining the record variable's type is the same as the value stated in the **Open** statement's **Len** clause. In the following example, assume `RecVar` is a variable of the appropriate type. You can use the **Len** function to specify the length, as follows:
    
  ```
  Open MyFile As #1 Len = Len(RecVar) 

  ```


    
    
- The variable in a  **Put** statement is (or includes) a variable-length string. Because a 2-byte descriptor is always added to a variable-length string placed in a random access file with **Put**, the variable-length string must be at least 2 characters shorter than the record length specified in the **Len** clause of the **Open** statement.
    
- The variable in a  **Put** statement is (or includes) a **Variant**. Like variable-length strings,[Variant data types](vbe-glossary.md) also require a 2-byte descriptor. Variants containing variable-length strings require a 4-byte descriptor. Therefore, for variable-length strings in a **Variant**, the string must be at least 4 bytes shorter than the record length specified in the **Len** clause.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

