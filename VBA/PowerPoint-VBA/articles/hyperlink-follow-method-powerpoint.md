---
title: Hyperlink.Follow Method (PowerPoint)
keywords: vbapp10.chm526011
f1_keywords:
- vbapp10.chm526011
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.Follow
ms.assetid: d56ace43-cf92-b3a6-abb4-dd7b87bc3feb
ms.date: 06/08/2017
---


# Hyperlink.Follow Method (PowerPoint)

Displays the HTML document associated with the specified hyperlink in a new Web browser window.


## Syntax

 _expression_. **Follow**

 _expression_ A variable that represents a **Hyperlink** object.


## Example

This example loads the document associated with the first hyperlink on slide one in a new instance of the Web browser.


```vb
ActivePresentation.Slides(1).Hyperlinks(1).Follow
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-powerpoint.md)

