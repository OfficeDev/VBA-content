
# Font.TrackingPreset Property (Publisher)

Returns or sets a  **PbTrackingPresetType** constant representing the preset tracking type for characters in the specified font in a text range. Read/write.


## Syntax

 _expression_. **TrackingPreset**

 _expression_A variable that represents a  **Font** object.


### Return Value

PbTrackingPresetType


## Remarks

The  **TrackingPreset** property value can be one of these **PbTrackingPresetType** constants.



| **pbTrackingCustom**|
| **pbTrackingLoose**|
| **pbTrackingMixed**|
| **pbTrackingNormal**|
| **pbTrackingTight**|
| **pbTrackingVeryLoose**|
| **pbTrackingVeryTight**|
Loose and very loose tracking leaves ample space between characters, whereas tight and very tight tracking can produce character overlap.


## Example

This example specifies tight tracking as the preset for the characters in the second story.


```vb
Sub TrackingType() 
 
 Application.ActiveDocument.Stories(2).TextRange _ 
 .Font.TrackingPreset = pbTrackingTight 
 
End Sub
```

