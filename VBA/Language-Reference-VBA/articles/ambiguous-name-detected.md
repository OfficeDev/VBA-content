---
title: Ambiguous name detected
keywords: vblr6.chm1032812
f1_keywords:
- vblr6.chm1032812
ms.prod: office
ms.assetid: e2bebd51-75cc-99f6-9dcf-81c9bd34e897
ms.date: 06/08/2017
---


# Ambiguous name detected

The [identifier](vbe-glossary.md) conflicts with another identifier or requires qualification. This error has the following causes and solutions:



- More than one object in the same [scope](vbe-glossary.md) may have elements with the same name.
    
    Qualify the element name by including the object name and a period. For example:
    
     _object.property_
    
    [Module-level](vbe-glossary.md) identifiers and[project](vbe-glossary.md)-level identifiers (module names and [referenced project](vbe-glossary.md) names) may be reused in a[procedure](vbe-glossary.md), although it makes programs harder to maintain and debug. However, if you want to refer to both items in the same procedure, the item having wider scope must be qualified. For example, if  `MyID` is declared at the module level of `MyModule`, and then a [procedure-level](vbe-glossary.md)[variable](vbe-glossary.md) is declared with the same name in the module, references to the module-level variable must be appropriately qualified:
    


```vb
Dim MyID As String 
Sub MySub 
MyModule.MyID = "This is module-level variable" 
Dim MyID As String 
MyID = "This is the procedure-level variable" 
Debug.Print MyID 
Debug.Print MyModule.MyID 
End Sub
```


    
    



- An identifier declared at module-level conflicts with a procedure name. For example, this error occurs if the variable  `MyID` is declared at module level, and then a procedure is defined with the same name:
    
```vb
Public MyID 
Sub MyID 
'. . . 
End Sub 
```


     In this case, you must change one of the names because qualification with a common module name would not resolve the ambiguity. Procedure names are **Public** by default, but variable names are **Private** unless specified as **Public**.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

