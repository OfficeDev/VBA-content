---
title: Member identifier already exists in object module from which this object module derives
keywords: vblr6.chm1040356
f1_keywords:
- vblr6.chm1040356
ms.prod: office
ms.assetid: 29eda9b1-b690-8c4a-7c57-fc938bbcd25a
ms.date: 06/08/2017
---


# Member identifier already exists in object module from which this object module derives

[Identifiers](vbe-glossary.md) used for object module members can't conflict with names already used in an[object module](vbe-glossary.md) from which they derive. This error has the following cause and solution:



- A [procedure](vbe-glossary.md) or data member identifier in your object module uses an identifier already used in the object module from which it derives. For example, a form has a **BackColor** property, so the following code would cause this error:
    
  ```
  ' Form already has a BackColor property. 
Dim BackColor As Integer    ' Generates the error. 
 
Function BackColor()    ' Generates the error. 
End Function
```


    Change the identifier that conflicts with the member identifier in your object module.
    
     **Note**  The following names cannot be used as property or method names because they belong to the underlying  **IUnknown** and **IDispatch** interfaces: **QueryInterface**, **AddRef**, **Release**, **GetTypeInfoCount**, **GetTypeInfo**, **GetIDsOfNames**, **Invoke**. Using these names causes a compilation error.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

