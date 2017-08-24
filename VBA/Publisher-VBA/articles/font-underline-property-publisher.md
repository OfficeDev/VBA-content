---
title: Font.Underline Property (Publisher)
keywords: vbapb10.chm5373987
f1_keywords:
- vbapb10.chm5373987
ms.prod: publisher
api_name:
- Publisher.Font.Underline
ms.assetid: a01a943e-274d-725e-3f78-aa76c51d5c46
ms.date: 06/08/2017
---


# Font.Underline Property (Publisher)

Returns or sets an  **PbUnderlineType** constant that indicates the type of underline for the selected characters in the specified font in a text range. Read/write.


## Syntax

 _expression_. **Underline**

 _expression_A variable that represents an  **Font** object.


### Return Value

PbUnderlineType


## Remarks

The  **Underline** property value can be one of the **[PbUnderlineType](pbunderlinetype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example formats the characters of the first story with a dashed and heavy underline.


```vb
Sub DashHeavy() 
 
 Application.ActiveDocument.Stories(1).TextRange.Font.Underline = pbUnderlineDashHeavy 
 
End Sub
```


