---
title: EffectParameter.Value Property (Office)
ms.prod: office
api_name:
- Office.EffectParameter.Value
ms.assetid: 45bf51fe-c049-1c8e-cc3b-fdbd5d6d7157
ms.date: 06/08/2017
---


# EffectParameter.Value Property (Office)

Retrieves or sets the value of the  **EffectParameter** object. Read/write


## Syntax

 _expression_. **Value**

 _expression_ An expression that returns a **EffectParameter** object.


## Example

The following code sets the first parameter of the  **PictureEffect** object as color temperature.


```
Dim picEffect As PictureEffect 
 
picEffect.EffectParameters(1).Value = MsoPictureEffectType.msoEffectColorTemperature
```


## See also


#### Concepts


[EffectParameter Object](effectparameter-object-office.md)
#### Other resources


[EffectParameter Object Members](effectparameter-members-office.md)

