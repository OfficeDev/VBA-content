---
title: Plate.InUse Property (Publisher)
keywords: vbapb10.chm2883602
f1_keywords:
- vbapb10.chm2883602
ms.prod: publisher
api_name:
- Publisher.Plate.InUse
ms.assetid: 6c98ada2-ff05-30c9-0043-afbe892dab3d
ms.date: 06/08/2017
---


# Plate.InUse Property (Publisher)

Returns  **True** if the specified ink (represented by the plate) is used in the publication. Read-only **Boolean**.


## Syntax

 _expression_. **InUse**

 _expression_A variable that represents an  **Plate** object.


### Return Value

Boolean


## Remarks

This property corresponds to the  **In Use** or **Not In Use** notation listed by each ink on the **Ink** tab of the **Color Printing** dialog box.


## Example

The following example loops through the active publication's plates collection, determines which plates represent inks that are not used in the publication, and deletes them.


```vb
Sub DeleteUnusedInks() 
 
Dim intCount As Integer 
 
With ActiveDocument.Plates 
 For intCount = .Count To 1 Step -1 
 With .Item(intCount) 
 If .InUse = False Then 
 Debug.Print "Name: " &; .Name 
 .Delete 
 End If 
 End With 
 Next 
End With 
 
End Sub
```


