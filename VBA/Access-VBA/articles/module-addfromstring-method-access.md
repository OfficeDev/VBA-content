---
title: Module.AddFromString Method (Access)
keywords: vbaac10.chm12273
f1_keywords:
- vbaac10.chm12273
ms.prod: access
api_name:
- Access.Module.AddFromString
ms.assetid: 119db9d9-fac2-b86f-be21-c94366bda7d6
ms.date: 06/08/2017
---


# Module.AddFromString Method (Access)

The  **AddFromString** method adds a string to a **[Module](module-object-access.md)** object. The **Module** object may represent a standard module or a class module.


## Syntax

 _expression_. **AddFromString**( ** _String_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _String_|Required|**String**|The information that you want to add to the module.|

### Return Value

Nothing


## Remarks

The  **AddFromString** method places the contents of a string after the Declarations section and before the first existing procedure in the module if the module contains other procedures.

In order to add a string to a form or report module, the form or report must be open in form Design view or report Design view. In order to add a string to a standard module or class module, the module must be open.


## Example

This example creates a new form and adds a string and the contents of the Functions.txt file to its module. Run the following procedure from a standard module:


```vb
Sub AddTextToFormModule() 
 Dim frm As Form, mdl As Module 
 
 Set frm = CreateForm 
 Set mdl = frm.Module 
 mdl.AddFromString "Public intY As Integer" 
 mdl.AddFromFile "C:\My Documents\Functions.txt" 
End Sub
```


## See also


#### Concepts


[Module Object](module-object-access.md)

