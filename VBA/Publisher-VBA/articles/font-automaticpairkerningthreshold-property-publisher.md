---
title: Font.AutomaticPairKerningThreshold Property (Publisher)
keywords: vbapb10.chm5373975
f1_keywords:
- vbapb10.chm5373975
ms.prod: publisher
api_name:
- Publisher.Font.AutomaticPairKerningThreshold
ms.assetid: f5f43a19-7227-b25d-9322-84a79596c525
ms.date: 06/08/2017
---


# Font.AutomaticPairKerningThreshold Property (Publisher)

Returns or sets a  **Variant** value that represents the point size above which kerning is automatically adjusted for characters in the specified text range. Read/write.


## Syntax

 _expression_. **AutomaticPairKerningThreshold**

 _expression_A variable that represents a  **Font** object.


### Return Value

Variant


## Remarks

Valid range is 0.0 points to 999.5 points. Returns -2 if the value for characters in the text range is indeterminate. Setting this property to 0.0 disables automatic pair kerning on the range.


## Example

This example sets the point size threshold to 12 points. All text in the second story above the threshold will implement auto kerning.


```vb
Sub Threshold() 
 
 Application.ActiveDocument.Stories(2).TextRange _ 
 .Font.AutomaticPairKerningThreshold = 12 
 
End Sub
```


