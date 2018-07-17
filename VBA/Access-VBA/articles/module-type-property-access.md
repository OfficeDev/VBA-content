---
title: Module.Type Property (Access)
keywords: vbaac10.chm12287
f1_keywords:
- vbaac10.chm12287
ms.prod: access
api_name:
- Access.Module.Type
ms.assetid: df30b007-5ce9-9de2-1013-747c47917288
ms.date: 06/08/2017
---


# Module.Type Property (Access)

Indicates whether a module is a standard module or a class module. Read-only  **[AcModuleType](acmoduletype-enumeration-access.md)**.


## Syntax

 _expression_. **Type**

 _expression_ A variable that represents a **Module** object.


## Example

The following example determines whether a  **Module** object represents a standard module or a class module:


```vb
Function CheckModuleType(strModuleName As String) As Integer 
 Dim mdl As Module 
 
 ' Open module to include in Modules collection. 
 DoCmd.OpenModule strModuleName 
 ' Return reference to Module object. 
 Set mdl = Modules(strModuleName) 
 ' Check Type property. 
 If mdl.Type = acClassModule Then 
 ' Insert comment. 
 mdl.InsertLines 1, "' Class module." 
 CheckModuleType = acClassModule 
 Else 
 ' Insert comment. 
 mdl.InsertLines 1, "' Standard module." 
 CheckModuleType = acStandardModule 
 End If 
End Function
```


## See also


#### Concepts


[Module Object](module-object-access.md)

