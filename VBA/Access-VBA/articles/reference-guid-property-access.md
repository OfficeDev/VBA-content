---
title: Reference.Guid Property (Access)
keywords: vbaac10.chm12631
f1_keywords:
- vbaac10.chm12631
ms.prod: access
api_name:
- Access.Reference.Guid
ms.assetid: a5419b60-f113-2c56-ff74-62c9ff8cc868
ms.date: 06/08/2017
---


# Reference.Guid Property (Access)

The  **GUID** property of a **[Reference](reference-object-access.md)** object returns a GUID that identifies a type library in the Windows Registry. Read-only **String**.


## Syntax

 _expression_. **Guid**

 _expression_ A variable that represents a **Reference** object.


## Remarks

Every type library has an associated GUID which is stored in the Registry. When you set a reference to a type library, Microsoft Access uses the type library's GUID to identify the type library.

You can use the  **[AddFromGUID](references-addfromguid-method-access.md)** method to create a **Reference** object from a type library's GUID.


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

