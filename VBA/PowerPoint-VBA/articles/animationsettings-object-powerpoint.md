---
title: AnimationSettings Object (PowerPoint)
keywords: vbapp10.chm565000
f1_keywords:
- vbapp10.chm565000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings
ms.assetid: ebbe4257-236b-35b4-bdf1-e92a1b4b417b
ms.date: 06/08/2017
---


# AnimationSettings Object (PowerPoint)

Represents the special effects applied to the animation for the specified shape during a slide show.


## Example

Use the [AnimationSettings](http://msdn.microsoft.com/library/c960d0de-afb3-55f2-b6fb-e67779cc42d2%28Office.15%29.aspx)property of the  **Shape** object to return the **AnimationSettings** object. The following example adds a slide that contains both a title and a three-item list to the active presentation, and then it sets the list to be animated by first-level paragraphs, to fly in from the left when animated, to dim to the specified color after being animated, and to animate its items in reverse order.


```
Set sObjs = ActivePresentation.Slides.Add(2, ppLayoutText).Shapes

sObjs.Title.TextFrame.TextRange.Text = "Top Three Reasons"

With sObjs.Placeholders(2)

    .TextFrame.TextRange.Text = _

        "Reason 1" &amp; VBNewLine &amp; "Reason 2" &amp; VBNewLine &amp; "Reason 3"

    With .AnimationSettings

        .TextLevelEffect = ppAnimateByFirstLevel

        .EntryEffect = ppEffectFlyFromLeft

        .AfterEffect = ppAfterEffectDim

        .DimColor.RGB = RGB(100, 120, 100)

        .AnimateTextInReverse = True

    End With

End With
```


## Properties



|**Name**|
|:-----|
|[AdvanceMode](http://msdn.microsoft.com/library/794d867f-cd7d-eeb6-0d6c-081e2be72ee5%28Office.15%29.aspx)|
|[AdvanceTime](http://msdn.microsoft.com/library/f4e5cec6-ba11-f605-3b3f-c4867fbce315%28Office.15%29.aspx)|
|[AfterEffect](http://msdn.microsoft.com/library/d8ccab29-8637-a48d-0f44-81a7fd1cca0b%28Office.15%29.aspx)|
|[Animate](http://msdn.microsoft.com/library/7434630f-3c73-4261-36f7-a26d45e9df11%28Office.15%29.aspx)|
|[AnimateBackground](http://msdn.microsoft.com/library/929ba50f-23c4-9dea-09fb-fa580715b118%28Office.15%29.aspx)|
|[AnimateTextInReverse](http://msdn.microsoft.com/library/cceba8ad-9896-10ef-5c11-7c93d370c82c%28Office.15%29.aspx)|
|[AnimationOrder](http://msdn.microsoft.com/library/0a29fb35-1cd8-4d12-184e-1132494a0864%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/caf149e6-302b-ff24-da9e-e604d4146480%28Office.15%29.aspx)|
|[ChartUnitEffect](http://msdn.microsoft.com/library/a2b66cf3-c8b9-6b9c-d184-13a828b474b2%28Office.15%29.aspx)|
|[DimColor](http://msdn.microsoft.com/library/574c24b0-45af-2e7c-6fd5-bfc17f552c83%28Office.15%29.aspx)|
|[EntryEffect](http://msdn.microsoft.com/library/de803113-6f7f-b1a2-1d52-43eeacccf666%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/73f01a7a-51c5-129f-34bf-2b7385e98ba5%28Office.15%29.aspx)|
|[PlaySettings](http://msdn.microsoft.com/library/2cfd1ed9-7ed0-0f69-4df5-43aa22e37f46%28Office.15%29.aspx)|
|[SoundEffect](http://msdn.microsoft.com/library/b357a83d-167b-5429-7d7d-94851c8735ac%28Office.15%29.aspx)|
|[TextLevelEffect](http://msdn.microsoft.com/library/008e3db2-2d22-5218-c312-663f0106adc6%28Office.15%29.aspx)|
|[TextUnitEffect](http://msdn.microsoft.com/library/6948db54-775a-39d6-9d90-99ad25f9cb80%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
