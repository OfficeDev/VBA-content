---
title: Sequence.ConvertToBuildLevel Method (PowerPoint)
keywords: vbapp10.chm651008
f1_keywords:
- vbapp10.chm651008
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.ConvertToBuildLevel
ms.assetid: ee674e55-dae3-1940-cf44-5520e8e82306
ms.date: 06/08/2017
---


# Sequence.ConvertToBuildLevel Method (PowerPoint)

Changes the build level information for a specified animation effect. Returns an  **[Effect](effect-object-powerpoint.md)** object that represents the build level information.


## Syntax

 _expression_. **ConvertToBuildLevel**( **_Effect_**, **_Level_** )

 _expression_ A variable that represents a **Sequence** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Effect_|Required|**Effect**| The specified animation effect.|
| _Level_|Required|**[MsoAnimateByLevel](msoanimatebylevel-enumeration-powerpoint.md)**|The animation build level.|

### Return Value

Effect


## Remarks

Changing build level information for an effect invalidates any existing effects.


## Example

The following example changes the build level information for an animation effect, making the original effect invalid.


```vb
Sub ConvertBuildLevel()

    Dim sldFirst As Slide
    Dim shpFirst As Shape
    Dim effFirst As Effect
    Dim effConvert As Effect

    Set sldFirst = ActiveWindow.Selection.SlideRange(1)
    Set shpFirst = sldFirst.Shapes(1)
    Set effFirst = sldFirst.TimeLine.MainSequence _
        .AddEffect(Shape:=shpFirst, EffectID:=msoAnimEffectAscend)

    Set effConvert = sldFirst.TimeLine.MainSequence _
        .ConvertToBuildLevel(Effect:=effFirst, _
        Level:=msoAnimateTextByFirstLevel)

End Sub
```


## See also


#### Concepts


[Sequence Object](sequence-object-powerpoint.md)

