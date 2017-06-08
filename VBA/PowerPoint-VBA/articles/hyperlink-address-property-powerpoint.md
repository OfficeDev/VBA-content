---
title: Hyperlink.Address Property (PowerPoint)
keywords: vbapp10.chm526004
f1_keywords:
- vbapp10.chm526004
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.Address
ms.assetid: d3d2174a-fbb2-432d-bc42-6623c91e9843
ms.date: 06/08/2017
---


# Hyperlink.Address Property (PowerPoint)

Returns or sets the Internet address (URL) to the target document. Read/write.


## Syntax

 _expression_. **Address**

 _expression_ A variable that represents an **Hyperlink** object.


### Return Value

String


## Example

This example scans all shapes on the first slide for the URL to the Microsoft Web site.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Hyperlinks

    If s.Address = "http://www.microsoft.com/" Then

        MsgBox "You have a link to the Microsoft Home Page"

    End If

Next
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-powerpoint.md)

