---
title: AddIn.FullName Property (PowerPoint)
keywords: vbapp10.chm521003
f1_keywords:
- vbapp10.chm521003
ms.prod: powerpoint
api_name:
- PowerPoint.AddIn.FullName
ms.assetid: 0e442ae8-ac67-d28c-d38f-b3d7e4ba9d34
ms.date: 06/08/2017
---


# AddIn.FullName Property (PowerPoint)

Returns the name of the specified add-in or saved presentation, including the path, the current file system separator, and the file name extension. Read-only.


## Syntax

 _expression_. **FullName**

 _expression_ A variable that represents an **AddIn** object.


### Return Value

String


## Remarks

This property is equivalent to the  **Path** property, followed by the current file system separator, followed by the **Name** property.


## Example

This example displays the path and file name of every available add-in.


```vb
For Each a In Application.AddIns

    MsgBox a.FullName

Next a
```

This example displays the path and file name of the active presentation (assuming that the presentation has been saved).




```vb
MsgBox Application.ActivePresentation.FullName
```


## See also


#### Concepts


[AddIn Object](addin-object-powerpoint.md)

