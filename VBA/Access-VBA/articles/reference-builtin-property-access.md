---
title: Reference.BuiltIn Property (Access)
keywords: vbaac10.chm12635
f1_keywords:
- vbaac10.chm12635
ms.prod: access
api_name:
- Access.Reference.BuiltIn
ms.assetid: 2c3f8eca-55b9-aa24-1a93-c8926e9587bd
ms.date: 06/08/2017
---


# Reference.BuiltIn Property (Access)

The  **BuiltIn** property returns a **Boolean** value indicating whether a **[Reference](reference-object-access.md)** object points to a default reference that's necessary for Microsoft Access to function properly. Read-only **Boolean**.


## Syntax

 _expression_. **BuiltIn**

 _expression_ A variable that represents a **Reference** object.


## Remarks

The  **BuiltIn** property is available only by using Visual Basic and is read-only.



|**Value**|**Description**|
|:-----|:-----|
|**True** (?1)|The  **Reference** object refers to a default reference that can't be removed.|
|**False** (0)|The  **Reference** object refers to a nondefault reference that isn't necessary for Microsoft Access to function properly.|
The default references in Microsoft Access include the Microsoft Access 12.0 object library, Microsoft Office 12.0 Access Connectivity Engine, the Visual Basic For Applications library, OLE Automation library, and Microsoft ActiveX Data Objects 2.5 library.


## Example

The following example prints the default references in the  **References** collection:


```vb
Sub ReferenceBuiltInOnly() 
 Dim ref As Reference 
 
 ' Enumerate through References collection. 
 For Each ref In References 
 ' Check BuiltIn property. 
 If ref.BuiltIn = True Then 
 Debug.Print ref.Name 
 End If 
 Next ref 
End Sub
```


## See also


#### Concepts


[Reference Object](reference-object-access.md)

