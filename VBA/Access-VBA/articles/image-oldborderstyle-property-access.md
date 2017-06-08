---
title: Image.OldBorderStyle Property (Access)
keywords: vbaac10.chm10384
f1_keywords:
- vbaac10.chm10384
ms.prod: access
api_name:
- Access.Image.OldBorderStyle
ms.assetid: ab5295d3-9e24-4604-a541-ac5bba837c0b
ms.date: 06/08/2017
---


# Image.OldBorderStyle Property (Access)

You can use this property to set or returns the unedited value of the  **BorderStyle** property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write **Byte**.


## Syntax

 _expression_. **OldBorderStyle**

 _expression_ A variable that represents an **Image** object.


## Remarks

The  **OldBorderStyle** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Transparent|0|(Default only for label, chart, and subreport) Transparent|
|Solid|1|(Default) Solid line|
|Dashes|2|Dashed line|
|Short dashes|3|Dashed line with short dashes|
|Dots|4|Dotted line|
|Sparse dots|5|Dotted line with dots spaced far apart|
|Dash dot|6|Line with a dash-dot combination|
|Dash dot dot|7|Line with a dash-dot-dot combination|
|Double solid|8|Double solid lines|

 **Note**  


## Example

The following example demonstrates the effect of changing a control's  **BorderStyle** property, while leaving the **OldBorderStyle** unaffected. The example concludes with setting the **BorderStyle** property to its original unedited value.


```vb
With Forms("Order Entry").Controls("Zip Code")
    .BorderStyle = 3 ' Short dashed border. 
  
    MsgBox "BorderStyle = " &; .BorderStyle &; vbCrLf &; _ 
        "OldBorderStyle = " &; .OldBorderStyle  ' Prints 3, 1. 
 
    .BorderStyle = 2 ' Dashed border. 
  
    MsgBox "BorderStyle = " &; .BorderStyle &; vbCrLf &; _ 
        "OldBorderStyle = " &; .OldBorderStyle  ' Prints 2, 1 
  
    .BorderStyle = .OldBorderStyle ' Solid (default) border. 
         
    MsgBox "BorderStyle = " &; .BorderStyle &; vbCrLf &; _ 
        "OldBorderStyle = " &; .OldBorderStyle  ' Prints 1, 1 
End With
```


## See also


#### Concepts


[Image Object](image-object-access.md)

