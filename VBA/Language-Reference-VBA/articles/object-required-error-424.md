---
title: Object required (Error 424)
keywords: vblr6.chm1000424
f1_keywords:
- vblr6.chm1000424
ms.prod: office
ms.assetid: 282292d7-d147-b71e-4d1e-149af7da8f7e
ms.date: 06/08/2017
---


# Object required (Error 424)

References to [properties](vbe-glossary.md) and[methods](vbe-glossary.md) often require an explicit object qualifier. This error has the following causes and solutions:



- You referred to an object property or method, but didn't provide a valid object qualifier. Specify an object qualifier if you didn't provide one. For example, although you can omit an object qualifier when referencing a form property from within the form's own [module](vbe-glossary.md), you must explicitly specify the qualifier when referencing the property from a [standard module](vbe-glossary.md).
    
- You supplied an object qualifier, but it isn't recognized as an object. Check the spelling of the object qualifier and make sure the object is visible in the part of the program in which you are referencing it. In the case of  **Collection** objects, check any occurrences of the **Add** method to be sure the syntax and spelling of all the elements are correct.
    
- You supplied a valid object qualifier, but some other portion of the call contained an error. An incorrect path as an [argument](vbe-glossary.md) to a[host application's ](vbe-glossary.md) **File Open** command could cause the error. Check arguments.
    
- You didn't use the  **Set** statement in assigning an object reference. If you assign the return value of a **CreateObject** call to a **Variant** variable, an error doesn't necessarily occur if the **Set** statement is omitted. In the following code example, an implicit instance of Microsoft Excel is created, and its default property (the string "Microsoft Excel") is returned and assigned to the **Variant** `RetVal`. A subsequent attempt to use  `RetVal` as an object reference causes this error:
    
```vb
Dim RetVal ' Implicitly a Variant. 
' Default property is assigned to Type 8 Variant RetVal. 
RetVal = CreateObject("Excel.Application") 
RetVal.Visible = True ' Error occurs here. 

  ```


    Use the  **Set** statement when assigning an object reference.
    
- In rare cases, this error occurs when you have a valid object but are attempting to perform an invalid action on the object. For example, you may receive this error if you try to assign a value to a read-only property. Check the object's documentation and make sure the action you are trying to perform is valid.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

