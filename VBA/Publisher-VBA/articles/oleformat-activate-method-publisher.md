---
title: OLEFormat.Activate Method (Publisher)
keywords: vbapb10.chm4456454
f1_keywords:
- vbapb10.chm4456454
ms.prod: publisher
api_name:
- Publisher.OLEFormat.Activate
ms.assetid: 43c01633-f624-c5ef-ba2c-d1ff62e91ec5
ms.date: 06/08/2017
---


# OLEFormat.Activate Method (Publisher)

Activates a window or OLE object.


## Syntax

 _expression_. **Activate**

 _expression_A variable that represents an  **OLEFormat** object.


## Remarks

Because Microsoft Publisher runs in a single window, using the  **Activate** method with a **Window** object makes Publisher the active application.


## Example

The following example makes Publisher the active application.


```vb
Application.ActiveWindow.Activate
```

The following example adds an Excel spreadsheet to the first page of the active publication and activates the spreadsheet for editing.




```vb
Dim shpSheet As Shape 
 
Set shpSheet = ActiveDocument.Pages(1).Shapes.AddOLEObject (Left:=72, Top:=72, ClassName:="Excel.Sheet") 
 
shpSheet.OLEFormat.Activate
```


