---
title: OLEFormat.ClassType Property (Word)
keywords: vbawd10.chm154337282
f1_keywords:
- vbawd10.chm154337282
ms.prod: word
api_name:
- Word.OLEFormat.ClassType
ms.assetid: 4c9ecec9-f7a9-f644-3a79-f88b9468200e
ms.date: 06/08/2017
---


# OLEFormat.ClassType Property (Word)

Returns or sets the class type for the specified OLE object, picture, or field. Read/write  **String** .


## Syntax

 _expression_ . **ClassType**

 _expression_ A variable that represents a **[OLEFormat](oleformat-object-word.md)** object.


## Remarks

This property is read-only for linked objects other than DDE links.

You can see a list of the available applications in the  **Object** type box on the **Create New** tab in the **Object** dialog box ( **Insert** menu). You can find the **ClassType** string by inserting an object as an inline shape and then viewing the field codes. The class type of the object follows either the word "EMBED" or the word "LINK."


## Example

This example loops through all the floating shapes on the active document and sets all linked Microsoft Excel worksheets to be updated automatically.


```vb
Dim shapeLoop As Shape 
 
For Each shapeLoop In ActiveDocument.Shapes 
 With shapeLoop 
 If .Type = msoLinkedOLEObject Then 
 If .OLEFormat.ClassType = "Excel.Sheet" Then 
 .LinkFormat.AutoUpdate = True 
 End If 
 End If 
 End With 
Next
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

