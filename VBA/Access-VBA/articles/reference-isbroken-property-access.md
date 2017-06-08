---
title: Reference.IsBroken Property (Access)
keywords: vbaac10.chm12636
f1_keywords:
- vbaac10.chm12636
ms.prod: access
api_name:
- Access.Reference.IsBroken
ms.assetid: 7a0bce38-0362-2645-a934-ddfb92322bcd
ms.date: 06/08/2017
---


# Reference.IsBroken Property (Access)

The  **IsBroken** property returns a **Boolean** value indicating whether a **[Reference](reference-object-access.md)** object points to a valid reference in the Windows Registry. Read-only **Boolean**.


## Syntax

 _expression_. **IsBroken**

 _expression_ A variable that represents a **Reference** object.


## Remarks

The default value of the  **IsBroken** property is **False**. The **IsBroken** property returns **True** only if the **Reference** object no longer points to a valid reference in the Registry.

By evaluating the  **IsBroken** property, you can determine whether or not the file associated with a particular **Reference** object has been moved to a different directory or deleted.

If the  **IsBroken** property is **True**, Microsoft Access generates an error when you try to read the **Name** or **FullPath** properties.


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

