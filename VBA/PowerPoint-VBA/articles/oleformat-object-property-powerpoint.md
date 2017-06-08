---
title: OLEFormat.Object Property (PowerPoint)
keywords: vbapp10.chm562004
f1_keywords:
- vbapp10.chm562004
ms.prod: powerpoint
api_name:
- PowerPoint.OLEFormat.Object
ms.assetid: fcaef43d-590e-179f-6698-4a8c191b92f9
ms.date: 06/08/2017
---


# OLEFormat.Object Property (PowerPoint)

Returns the object that represents the specified OLE object's top-level interface. Read-only.


## Syntax

 _expression_. **Object**

 _expression_ A variable that represents an **OLEFormat** object.


### Return Value

Object


## Remarks

This property allows you to access the properties and methods of the application in which an OLE object was created.

Use the  **TypeName** function to determine the type of object this property returns for a specific OLE object.


## Example

This example displays the type of object contained in shape one on slide one in the active presentation. Shape one must contain an OLE object.


```vb
MsgBox TypeName(ActivePresentation.Slides(1) _
    .Shapes(1).OLEFormat.Object)
```

This example displays the name of the application in which each embedded OLE object on slide one in the active presentation was created.




```vb
For Each s In ActivePresentation.Slides(1).Shapes

    If s.Type = msoEmbeddedOLEObject Then

        MsgBox s.OLEFormat.Object.Application.Name

    End If

Next


```

This example adds text to cell A1 on worksheet one in the Microsoft Excel workbook contained in shape three on slide one in the active presentation.




```vb
With ActivePresentation.Slides(1).Shapes(3)

    .OLEFormat.Object.Worksheets(1).Range("A1").Value = "New text"

End With
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-powerpoint.md)

