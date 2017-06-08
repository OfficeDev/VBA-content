---
title: Font.NameAscii Property (PowerPoint)
keywords: vbapp10.chm575017
f1_keywords:
- vbapp10.chm575017
ms.prod: powerpoint
api_name:
- PowerPoint.Font.NameAscii
ms.assetid: 06db0f5b-71ac-704d-eef2-1be8a96fb7a8
ms.date: 06/08/2017
---


# Font.NameAscii Property (PowerPoint)

Returns or sets the font used for ASCII characters (characters with character set numbers within the range of 0 to 127). Read/write.


## Syntax

 _expression_. **NameAscii**

 _expression_ A variable that represents a **Font** object.


### Return Value

String


## Remarks

The default value of this property is Times New Roman. Use the  **[Replace](fonts-replace-method-powerpoint.md)** method to change the font that's applied to all text and that appears in the **Font** box on the **Font** tab.


## Example

This example sets the font used for ASCII characters in the title of the first slide to Century.


```vb
Application.ActivePresentation.Slides(1).Shapes.Title _
    .TextFrame.TextRange.Font.NameAscii = "Century"
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

