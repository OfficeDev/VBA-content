---
title: SlideShowSettings Object (PowerPoint)
keywords: vbapp10.chm514000
f1_keywords:
- vbapp10.chm514000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings
ms.assetid: d58c7c3b-a1cc-d819-b386-fd3fb7f967a2
ms.date: 06/08/2017
---


# SlideShowSettings Object (PowerPoint)

Represents the slide show setup for a presentation.


## Example

Use the [SlideShowSettings](http://msdn.microsoft.com/library/90a5a5cb-1f78-bbb2-8e4c-eb35aae13c90%28Office.15%29.aspx)property to return the  **SlideShowSettings** object. The first section in the following example sets all the slides in the active presentation to advance automatically after five seconds. The second section sets the slide show to start on slide two, end on slide four, advance slides by using the timings set in the first section, and run in a continuous loop until the user presses ESC. Finally, the example runs the slide show.


```
For Each s In ActivePresentation.Slides

    With s.SlideShowTransition

        .AdvanceOnTime = True

        .AdvanceTime = 5

    End With

Next



With ActivePresentation.SlideShowSettings

    .RangeType = ppShowSlideRange

    .StartingSlide = 2

    .EndingSlide = 4

    .AdvanceMode = ppSlideShowUseSlideTimings

    .LoopUntilStopped = True

    .Run

End With
```


## Methods



|**Name**|
|:-----|
|[Run](http://msdn.microsoft.com/library/497fae3b-b6a3-dc26-20d9-bdc8057ddc09%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AdvanceMode](http://msdn.microsoft.com/library/0fc398c3-b7e6-5301-a19d-381d8ff35155%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/ec61fee1-46bd-d385-0d50-4c2c0d82b43e%28Office.15%29.aspx)|
|[EndingSlide](http://msdn.microsoft.com/library/50489e3a-bdfe-b495-97d1-69ba1d7bf2b9%28Office.15%29.aspx)|
|[LoopUntilStopped](http://msdn.microsoft.com/library/767a5865-b50b-d7c6-6076-6786b43c6b88%28Office.15%29.aspx)|
|[NamedSlideShows](http://msdn.microsoft.com/library/8af7610f-1981-df5f-5be8-2bb04c895602%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/8ddb2bac-f057-2532-5825-3346046afe8c%28Office.15%29.aspx)|
|[PointerColor](http://msdn.microsoft.com/library/530072d6-3a2d-8236-b4ac-3ede8823e95a%28Office.15%29.aspx)|
|[RangeType](http://msdn.microsoft.com/library/63e266b6-4898-abb1-23fe-20039a6aea78%28Office.15%29.aspx)|
|[ShowMediaControls](http://msdn.microsoft.com/library/6b7a63d3-f43d-bbb2-0af2-574e19d48e3d%28Office.15%29.aspx)|
|[ShowPresenterView](http://msdn.microsoft.com/library/62ec6a39-1e8d-f6e5-0769-64a175d4d611%28Office.15%29.aspx)|
|[ShowScrollbar](http://msdn.microsoft.com/library/9f6be3f3-1099-2f8c-4c1c-b5ab1be89f4a%28Office.15%29.aspx)|
|[ShowType](http://msdn.microsoft.com/library/6537dd4c-8029-3e95-7073-7701ba12a627%28Office.15%29.aspx)|
|[ShowWithAnimation](http://msdn.microsoft.com/library/9255fc7b-50fa-c65e-5ef4-3c214dede4a4%28Office.15%29.aspx)|
|[ShowWithNarration](http://msdn.microsoft.com/library/65390c53-abeb-ca9e-0697-f68dcb455324%28Office.15%29.aspx)|
|[SlideShowName](http://msdn.microsoft.com/library/212a2851-cc73-76ad-98fa-f295ae3c89c8%28Office.15%29.aspx)|
|[StartingSlide](http://msdn.microsoft.com/library/e7afc69c-0224-b22a-fc23-bb985e710c1a%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
