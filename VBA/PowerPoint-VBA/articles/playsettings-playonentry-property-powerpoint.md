---
title: PlaySettings.PlayOnEntry Property (PowerPoint)
keywords: vbapp10.chm568006
f1_keywords:
- vbapp10.chm568006
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.PlayOnEntry
ms.assetid: 63a226b9-b0f2-b739-ced2-f4e57a91b5f5
ms.date: 06/08/2017
---


# PlaySettings.PlayOnEntry Property (PowerPoint)

Determines whether the specified movie or sound is played automatically when it is animated. Read/write.


## Syntax

 _expression_. **PlayOnEntry**

 _expression_ A variable that represents a **PlaySettings** object.


### Return Value

MsoTriState


## Remarks

Setting this property to  **msoTrue** sets the **[Animate](animationsettings-animate-property-powerpoint.md)** property of the **AnimationSettings** object to **msoTrue**. Setting the **Animate** property to **msoFalse** automatically sets the **PlayOnEntry** property to **msoFalse**.

Use the  **[ActionVerb](actionsetting-actionverb-property-powerpoint.md)** property to set the verb that will be invoked when the media clip is animated.

The value of the  **PlayOnEntry** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified movie or sound is not played automatically when it is animated.|
|**msoTrue**| The specified movie or sound is played automatically when it is animated.|

## Example

This example specifies that shape three on slide one in the active presentation will be played automatically when it is animated. Shape three must be a sound or movie object.


```vb
Set OLEobj = ActivePresentation.Slides(1).Shapes(3)

OLEobj.AnimationSettings.PlaySettings.PlayOnEntry = msoTrue
```


## See also


#### Concepts


[PlaySettings Object](playsettings-object-powerpoint.md)

