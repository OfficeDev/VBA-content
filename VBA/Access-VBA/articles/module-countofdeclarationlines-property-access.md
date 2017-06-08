---
title: Module.CountOfDeclarationLines Property (Access)
keywords: vbaac10.chm12284
f1_keywords:
- vbaac10.chm12284
ms.prod: access
api_name:
- Access.Module.CountOfDeclarationLines
ms.assetid: fc0bdb0f-264c-0311-946c-c5cdc03a86f0
ms.date: 06/08/2017
---


# Module.CountOfDeclarationLines Property (Access)

The  **CountOfDeclarationLines** property returns a **Long** value indicating the number of lines of code in the Declarations section in a standard module or class module. Read-only **Long**.


## Syntax

 _expression_. **CountOfDeclarationLines**

 _expression_ A variable that represents a **Module** object.


## Remarks

Lines in a module are numbered beginning with 1.

The value of the  **CountOfDeclarationLines** property is equal to the line number of the last line of the Declarations section. You can use this property to determine where the Declarations section ends and the body of the module begins.


## Example

The following example counts the number of lines and declaration lines in each standard module in the  **Modules** collection. Note that the **Modules** collection contains only modules that are open in the module editor.


```vb
Public Sub ModuleLineTotal(ByVal strModuleName As String) 
 
 Dim mdl As Module 
 
 ' Open module to include in Modules collection. 
 DoCmd.OpenModule strModuleName 
 
 ' Return reference to Module object. 
 Set mdl = Modules(strModuleName) 
 
 ' Print number of lines in module. 
 Debug.Print "Number of lines: ", mdl.CountOfLines 
 
 ' Print number of declaration lines. 
 Debug.Print "Number of declaration lines: ", _ 
 mdl.CountOfDeclarationLines 
 
End Sub
```


## See also


#### Concepts


[Module Object](module-object-access.md)

