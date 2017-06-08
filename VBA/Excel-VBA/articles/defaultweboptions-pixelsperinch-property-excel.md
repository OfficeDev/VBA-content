---
title: DefaultWebOptions.PixelsPerInch Property (Excel)
keywords: vbaxl10.chm660084
f1_keywords:
- vbaxl10.chm660084
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.PixelsPerInch
ms.assetid: 9264ea44-cfc7-8640-4606-71a17b806a48
ms.date: 06/08/2017
---


# DefaultWebOptions.PixelsPerInch Property (Excel)

Returns or sets the density (pixels per inch) of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. The default setting is 96. Read/write  **Long** .


## Syntax

 _expression_ . **PixelsPerInch**

 _expression_ A variable that represents a **DefaultWebOptions** object.


## Remarks

This property determines the size of the images and cells on the specified Web page relative to the size of text whenever you view the saved document in a Web browser. The physical dimensions of the resulting image or cell are the result of the original dimensions (in inches) multiplied by the number of pixels per inch.

You use the  **[ScreenSize](defaultweboptions-screensize-property-excel.md)** property to set the optimum screen size for the targeted Web browsers.


## Example

This example sets the pixel density depending on the target screen size of the browser. For 800x600 pixel screens, the density is 72 pixels per inch. For 1024x768 pixel screens, the density is 96 pixels per inch. For all other cases, use a density of 120 pixels per inch.


```vb
With Application.DefaultWebOptions 
 Select Case .ScreenSize 
 Case msoScreenSize800x600 
 .PixelsPerInch = 72 
 Case msoScreenSize1024x768 
 .PixelsPerInch = 96 
 Case Else 
 .PixelsPerInch = 120 
 End Select 
End With
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-excel.md)

