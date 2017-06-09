---
title: Module.CountOfLines Property (Access)
keywords: vbaac10.chm12276
f1_keywords:
- vbaac10.chm12276
ms.prod: access
api_name:
- Access.Module.CountOfLines
ms.assetid: 6c3bb4c8-15a9-6365-155d-d28dc0c5de78
ms.date: 06/08/2017
---


# Module.CountOfLines Property (Access)

The  **CountOfLines** property returns a **Long** value indicating the number of lines of code in a standard module or class module. Read-only **Long**.


## Syntax

 _expression_. **CountOfLines**

 _expression_ A variable that represents a **Module** object.


## Remarks

Lines in a module are numbered beginning with 1.

The line number of the last line in a module is the value of the  **CountOfLines** property.


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

