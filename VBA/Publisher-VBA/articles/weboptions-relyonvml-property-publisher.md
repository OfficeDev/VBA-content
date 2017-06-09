---
title: WebOptions.RelyOnVML Property (Publisher)
keywords: vbapb10.chm8257543
f1_keywords:
- vbapb10.chm8257543
ms.prod: publisher
api_name:
- Publisher.WebOptions.RelyOnVML
ms.assetid: 8cd29d64-48a6-d33e-cb9d-6b1ea094069a
ms.date: 06/08/2017
---


# WebOptions.RelyOnVML Property (Publisher)

Returns or sets a  **Boolean** value that specifies whether image files are generated from drawing objects when a Web publication is saved. If **True**, image files are not generated. If  **False**, images are generated. The default value is  **False**. Read/write.


## Syntax

 _expression_. **RelyOnVML**

 _expression_A variable that represents a  **WebOptions** object.


### Return Value

Boolean


## Remarks

File sizes can be reduced by not generating images for drawing objects. Note that a Web browser must support Vector Markup Language (VML) to render drawing objects. Microsoft Internet Explorer versions 5.0 and later support VML, so the  **RelyOnVML** property could be set to **True** if targeting those browsers. For browsers that do not support VML, a drawing object will not appear when a Web publication is saved with this property enabled.

If unsure about which browsers will be used to view the Web site, this property should be set to  **False**.


## Example

The following example assumes that end users have Microsoft Internet Explorer version 5.0, and therefore specifies that images should not be generated from drawing objects when the Web publication is saved.


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .RelyOnVML = True 
End With
```


