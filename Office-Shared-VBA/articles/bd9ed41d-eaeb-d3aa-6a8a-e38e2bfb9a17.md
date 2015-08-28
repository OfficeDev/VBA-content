
# GradientStops.Insert2 Method (Office)

 **Last modified:** July 28, 2015

Adds a stop to a gradient and specifies the brightness, as well as the transparency, of the color.

## Syntax

 _expression_. **Insert2**( **_RGB_**,  **_Position_**,  **_Transparency_**,  **_Index_**,  **_Brightness_**)

 _expression_An expression that returns a  **GradientStops** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|RGB|Required| **MsoRGBType**|Specifies the color at the gradient stop.|
|Position|Required| **Single**|Specifies the position of the stop within the gradient expressed as a percent.|
|Transparency|Optional| **Single**|Specifies the opacity of the color at the gradient stop.|
|Index|Optional| **Integer**|The index number of the gradient stop.|
|Brightness|Optional| **Single**|Specifies the brightness of the color at the gradient stop.|

### Return Value

Nothing


## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops. 

This method differs from the  [Insert](98aec7ed-44f9-c9b4-7a1a-e5b9a1d26d95.md) method in that it allows you to specify the brightness, as well as the transparency, of the color at the gradient stop.


## See also


#### Concepts


 [GradientStops Object](365949f0-29b3-76e1-1163-2ac870f68f7a.md)
#### Other resources


 [GradientStops Object Members](9cab316d-3302-a119-b02b-54eea372acee.md)
