---
title: GradientStops.Insert2 Method (Office)
ms.prod: office
api_name:
- Office.GradientStops.Insert2
ms.assetid: bd9ed41d-eaeb-d3aa-6a8a-e38e2bfb9a17
ms.date: 06/08/2017
---


# GradientStops.Insert2 Method (Office)

Adds a stop to a gradient and specifies the brightness, as well as the transparency, of the color.


## Syntax

 _expression_. **Insert2**( **_RGB_**, **_Position_**, **_Transparency_**, **_Index_**, **_Brightness_** )

 _expression_ An expression that returns a **GradientStops** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RGB_|Required|**MsoRGBType**|Specifies the color at the gradient stop.|
| _Position_|Required|**Single**|Specifies the position of the stop within the gradient expressed as a percent.|
| _Transparency_|Optional|**Single**|Specifies the opacity of the color at the gradient stop.|
| _Index_|Optional|**Integer**|The index number of the gradient stop.|
| _Brightness_|Optional|**Single**|Specifies the brightness of the color at the gradient stop.|

### Return Value

Nothing


## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops. 

This method differs from the [Insert](gradientstops-insert-method-office.md) method in that it allows you to specify the brightness, as well as the transparency, of the color at the gradient stop.


## See also


#### Concepts


[GradientStops Object](gradientstops-object-office.md)
#### Other resources


[GradientStops Object Members](gradientstops-members-office.md)

