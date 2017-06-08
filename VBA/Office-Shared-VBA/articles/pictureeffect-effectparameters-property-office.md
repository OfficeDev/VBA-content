---
title: PictureEffect.EffectParameters Property (Office)
ms.prod: office
api_name:
- Office.PictureEffect.EffectParameters
ms.assetid: a0729015-14ab-e5c3-9772-678b892e4834
ms.date: 06/08/2017
---


# PictureEffect.EffectParameters Property (Office)

Returns an  **EffectParameter** object. Read-only


## Syntax

 _expression_. **EffectParameters**

 _expression_ An expression that returns a **PictureEffect** object.


## Remarks

Picture Effects are processed as a chain composed of individual items which are applied in sequence to create the final composited image. An Effects chain will allow an effect to be added to the chain, reordered, or removed from the chain. Effect Parameters specify properties of those effects.


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


[PictureEffect Object](pictureeffect-object-office.md)
#### Other resources


[PictureEffect Object Members](pictureeffect-members-office.md)

