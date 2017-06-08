---
title: Application.BrokenReference Property (Access)
keywords: vbaac10.chm12593
f1_keywords:
- vbaac10.chm12593
ms.prod: access
api_name:
- Access.Application.BrokenReference
ms.assetid: 20a55f4b-5fe4-9231-bbef-e90c66f88b90
ms.date: 06/08/2017
---


# Application.BrokenReference Property (Access)

Returns a  **Boolean** indicating whether the current database has any broken references to databases or type libraries. **True** if there are any broken references. Read-only.


## Syntax

 _expression_. **BrokenReference**

 _expression_ A variable that represents an **Application** object.


## Remarks

To test the validity of a specific reference, use the  **[IsBroken](reference-isbroken-property-access.md)** property of the **[Reference](reference-object-access.md)** object.


## Example

This example checks to see if there are any broken references in the current database and reports the results to the user.


```vb
' Looping variable. 
Dim refLoop As Reference 
' Output variable. 
Dim strReport As String 
 
' Test whether there are broken references. 
If Application.BrokenReference = True Then 
 strReport = "The following references are broken:" &; vbCr 
 
 ' Test validity of each reference. 
 For Each refLoop In Application.References 
 If refLoop.IsBroken = True Then 
 strReport = strReport &; " " &; refLoop.Name &; vbCr 
 End If 
 Next refLoop 
Else 
 strReport = "All references in the current database are valid." 
End If 
 
' Display results. 
MsgBox strReport
```


## See also


#### Concepts


[Application Object](application-object-access.md)

