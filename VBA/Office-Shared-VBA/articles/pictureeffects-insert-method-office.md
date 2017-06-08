---
title: PictureEffects.Insert Method (Office)
ms.prod: office
api_name:
- Office.PictureEffects.Insert
ms.assetid: 589c38d7-1d0a-ad87-a84c-72147b6b07cf
ms.date: 06/08/2017
---


# PictureEffects.Insert Method (Office)

Inserts a Picture Effect in a chain of composite effects.


## Syntax

 _expression_. **Insert**( **_EffectType_**, **_Position_** )

 _expression_ An expression that returns a **PictureEffects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EffectType_|Required|**MsoPictureEffectType**|An enumeration specifying the type of Picture Effect.|
| _Position_|Optional|**Integer**|The position of the effect in the composite chain of Picture Effects.|

### Return Value

PictureEffect


## Remarks

Picture Effects are processed as a chain composed of individual items which are applied in sequence to create the final composited image. An Effects chain will allow an effect to be added to the chain, reordered, or removed from the chain.


## Example

The following code sets several Picture Effect fill properties on a shape in a Microsoft PowerPoint slide.


```
Sub PictureEffectSample() 
' Setup a slide with one picture shape. 
With ActivePresentation.Slides(1).Shapes(1).Fill.PictureEffects 
 
 ' Insert a 150% Saturation effect. 
 .Insert(msoEffectSaturation).EffectParameters(1).Value = 1.5 
 
 ' Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast. 
 Dim brightnessContrast As PictureEffect 
 Set brightnessContrast = .Insert(msoEffectBrightnessContrast) 
 brightnessContrast.EffectParameters(1).Value = -0.5 
 brightnessContrast.EffectParameters(2).Value = 0.25 
 
 ' Remove all Picture effects. 
 While .Count > 0 
 .Delete (1) 
 Wend 
 
End With 
End Sub
```


## See also


#### Concepts


[PictureEffects Object](pictureeffects-object-office.md)
#### Other resources


[PictureEffects Object Members](pictureeffects-members-office.md)

