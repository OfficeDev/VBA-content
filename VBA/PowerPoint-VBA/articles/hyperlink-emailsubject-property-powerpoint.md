---
title: Hyperlink.EmailSubject Property (PowerPoint)
keywords: vbapp10.chm526007
f1_keywords:
- vbapp10.chm526007
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.EmailSubject
ms.assetid: 2416a620-9788-5da9-3095-432cab5cdc95
ms.date: 06/08/2017
---


# Hyperlink.EmailSubject Property (PowerPoint)

Returns or sets the text string of the hyperlink subject line. The subject line is appended to the Internet address (URL) of the hyperlink. Read/write.


## Syntax

 _expression_. **EmailSubject**

 _expression_ A variable that represents an **Hyperlink** object.


### Return Value

String


## Remarks

This property is commonly used with e-mail hyperlinks. The value of this property takes precedence over any e-mail subject specified in the  **[Address](hyperlink-address-property-powerpoint.md)** property of the same **Hyperlink** object.


## Example

This example sets the e-mail subject line of the first hyperlink on slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Hyperlinks(1) _
    .EmailSubject = "Quote Request"
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-powerpoint.md)

