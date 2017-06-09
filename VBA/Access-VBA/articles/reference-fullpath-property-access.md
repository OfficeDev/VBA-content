---
title: Reference.FullPath Property (Access)
keywords: vbaac10.chm12634
f1_keywords:
- vbaac10.chm12634
ms.prod: access
api_name:
- Access.Reference.FullPath
ms.assetid: 41e2b1b5-a0fd-79a0-27f2-71b996cc25ea
ms.date: 06/08/2017
---


# Reference.FullPath Property (Access)

The  **FullPath** property returns a string containing the path and file name of the referenced type library.


## Syntax

 _expression_. **FullPath**

 _expression_ A variable that represents a **Reference** object.


## Remarks

If the  **[IsBroken](reference-isbroken-property-access.md)** property setting of a **Reference** object is **True**, reading the **FullPath** property generates an error.


## Example

The following example prints the value of the  **FullPath**, **GUID**, **IsBroken**, **Major**, and **Minor** properties for each **Reference** object in the **References** collection:


```vb
Sub ReferenceProperties() 
 Dim ref As Reference 
 
 ' Enumerate through References collection. 
 For Each ref In References 
 ' Check IsBroken property. 
 If ref.IsBroken = False Then 
 Debug.Print "Name: ", ref.Name 
 Debug.Print "FullPath: ", ref.FullPath 
 Debug.Print "Version: ", ref.Major &; "." &; ref.Minor 
 Else 
 Debug.Print "GUIDs of broken references:" 
 Debug.Print ref.GUID 
 EndIf 
 Next ref 
End Sub
```


## See also


#### Concepts


[Reference Object](reference-object-access.md)

