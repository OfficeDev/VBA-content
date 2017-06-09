---
title: OLEFormat.IconIndex Property (Word)
keywords: vbawd10.chm154337289
f1_keywords:
- vbawd10.chm154337289
ms.prod: word
api_name:
- Word.OLEFormat.IconIndex
ms.assetid: 091bd36d-75f6-b31b-ca8f-668a23f215d7
ms.date: 06/08/2017
---


# OLEFormat.IconIndex Property (Word)

Returns or sets the icon that is used when the  **[DisplayAsIcon](oleformat-displayasicon-property-word.md)** property is **True** . Read/write **Long** .


## Syntax

 _expression_ . **IconIndex**

 _expression_ An expression that returns an **[OLEFormat](oleformat-object-word.md)** object.


## Remarks

Zero (0) corresponds to the first icon, 1 corresponds to the second icon, and so on. If this argument is omitted, the first (default) icon is used.


## Example

This example returns the icon index number in a message box for the first selected shape that's displayed as an icon.


```vb
Dim olefTemp As OLEFormat 
 
If Selection.ShapeRange.Count >= 1 Then 
 Set olefTemp = Selection.ShapeRange(1).OLEFormat 
 With olefTemp 
 If .DisplayAsIcon = True Then Msgbox .IconIndex 
 End With 
End If
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

