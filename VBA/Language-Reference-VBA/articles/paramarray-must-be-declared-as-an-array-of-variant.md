---
title: ParamArray must be declared as an array of Variant
keywords: vblr6.chm1040144
f1_keywords:
- vblr6.chm1040144
ms.prod: office
ms.assetid: d6c8fce1-590f-53c3-8379-a5324003134e
ms.date: 06/08/2017
---


# ParamArray must be declared as an array of Variant

Each [argument](vbe-glossary.md) to a **ParamArray**[parameter](vbe-glossary.md) can be of a different[data type](vbe-glossary.md). Therefore, the parameter itself must be declared as an [array](vbe-glossary.md) of **Variant** type. You can also supply any number of arguments to a **ParamArray**. When the call is made, each argument supplied in the call becomes a corresponding element of the **Variant** array. For example:


```vb
Sub MySub(ParamArray VarArg()) 
    . . . 
End Sub 
Call MySub ("First arg", 2, 3.54) 

```


This error has the following causes and solutions:



- In the definition of the [procedure](vbe-glossary.md), the  **ParamArray** parameter is defined as an array of a type other than **Variant**.
    
    Redeclare the parameter type as an array of  **Variant** elements.
    
- No data type was specified for the  **ParamArray** parameter, but the procedure definition is within the scope of a **Def**_type_ statement, so the parameter is implicitly declared as having a type other than **Variant**. Use an explicit **As Variant** clause in the specification of the **ParamArray** parameter.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

