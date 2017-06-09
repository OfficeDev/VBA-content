---
title: OLEFormat.DisplayAsIcon Property (Word)
keywords: vbawd10.chm154337283
f1_keywords:
- vbawd10.chm154337283
ms.prod: word
api_name:
- Word.OLEFormat.DisplayAsIcon
ms.assetid: eb27a24c-69f0-a94d-b2cb-0fc0ccb54a1a
ms.date: 06/08/2017
---


# OLEFormat.DisplayAsIcon Property (Word)

 **True** if the specified object is displayed as an icon. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayAsIcon**

 _expression_ A variable that represents a **[OLEFormat](oleformat-object-word.md)** object.


## Example

This example displays a message box containing the name of each floating shape that's displayed as an icon on the active document.


```vb
Dim shapeLoop As Shape 
 
For Each shapeLoop In ActiveDocument.Shapes 
 If shapeLoop.OLEFormat.DisplayAsIcon Then 
 MsgBox shapeLoop.Name &; " is displayed as an icon." 
 End If 
Next shapeLoop
```

This example inserts a Microsoft Excel worksheet as a linked OLE object on the active document and then changes the display of the object to an icon.




```vb
Dim objNew As Object 
 
Set objNew = ActiveDocument.Shapes.AddOLEObject _ 
 (FileName:="C:\Program Files\Microsoft Office" _ 
 &; "\Office\Samples\samples.xls", LinkToFile:=True) 
 
objNew.OLEFormat.DisplayAsIcon = True
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

