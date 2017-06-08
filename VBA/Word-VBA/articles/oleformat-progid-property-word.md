---
title: OLEFormat.ProgID Property (Word)
keywords: vbawd10.chm154337302
f1_keywords:
- vbawd10.chm154337302
ms.prod: word
api_name:
- Word.OLEFormat.ProgID
ms.assetid: f3e99411-ebea-9135-e25d-66948f53e037
ms.date: 06/08/2017
---


# OLEFormat.ProgID Property (Word)

Returns the programmatic identifier (ProgID) for the specified OLE object. Read-only  **String** .


## Syntax

 _expression_ . **ProgID**

 _expression_ Required. A variable that represents an **[OLEFormat](oleformat-object-word.md)** object.


## Remarks

The  **ProgID** and **ClassType** properties will (by default) return the same string. However, you can change the **ClassType** property for DDE links.


 **Security Note**  



For information about programmatic identifiers, see [OLE Programmatic Identifiers](http://msdn.microsoft.com/library/b68618d9-81e6-d97f-f706-f80a30d0f082%28Office.15%29.aspx).


## Example

This example loops through all the floating shapes in the active document and sets all linked Microsoft Excel worksheets to be updated automatically.


```vb
For Each s In ActiveDocument.Shapes 
 If s.Type = msoLinkedOLEObject Then 
 If s.OLEFormat.ProgID = "Excel.Sheet" Then 
 s.LinkFormat.AutoUpdate = True 
 End If 
 End If 
Next
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

