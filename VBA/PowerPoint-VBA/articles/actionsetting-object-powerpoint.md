---
title: ActionSetting Object (PowerPoint)
keywords: vbapp10.chm567000
f1_keywords:
- vbapp10.chm567000
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting
ms.assetid: 21381ff0-b9ff-59d8-77e9-345905fb8617
ms.date: 06/08/2017
---


# ActionSetting Object (PowerPoint)

Contains information about how the specified shape or text range reacts to mouse actions during a slide show. 


## Remarks

The  **ActionSetting** object is a member of the **[ActionSettings](http://msdn.microsoft.com/library/8914c203-6b8d-fa80-16ad-7015595657b7%28Office.15%29.aspx)** collection. The **ActionSettings** collection contains one **ActionSetting** object that represents how the specified object reacts when the user clicks it during a slide show and one **ActionSetting** object that represents how the specified object reacts when the user moves the mouse pointer over it during a slide show.

If you've set properties of the  **ActionSetting** object that don't seem to be taking effect, make sure that you've set the[Action](http://msdn.microsoft.com/library/32ed5574-5ac0-abb7-d300-6644fc894ec1%28Office.15%29.aspx) property to the appropriate value.


## Example

Use  **ActionSettings** (index), where index is the either **ppMouseClick** or **ppMouseOver**, to return a single **ActionSetting** object. The following example sets the mouse-click action for the text in the third shape on slide one in the active presentation to an Internet link.


```
With ActivePresentation.Slides(1).Shapes(3) _ 
        .TextFrame.TextRange.ActionSettings(ppMouseClick) 
    .Action = ppActionHyperlink 
    .Hyperlink.Address = "http://www.microsoft.com" 
End With
```


## Properties



|**Name**|
|:-----|
|[Action](http://msdn.microsoft.com/library/32ed5574-5ac0-abb7-d300-6644fc894ec1%28Office.15%29.aspx)|
|[ActionVerb](http://msdn.microsoft.com/library/f7b57e12-0c70-bc62-b94d-7ae8f65f7de0%28Office.15%29.aspx)|
|[AnimateAction](http://msdn.microsoft.com/library/cf6c13e4-1fc5-8335-16b3-9a9f30c246ea%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/a8792fb6-587c-20ee-1fe7-bf0927f96803%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/8654000a-bbc5-6d23-e5a7-d689bc767b1b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/ade56ee1-5664-64a4-8936-1c80630a82fe%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/5c5bc9ee-528c-ca49-0c36-c1f343671ffd%28Office.15%29.aspx)|
|[ShowAndReturn](http://msdn.microsoft.com/library/76797234-161d-50a5-cbc3-b1a169bc6719%28Office.15%29.aspx)|
|[SlideShowName](http://msdn.microsoft.com/library/680e998d-feba-3010-d0d4-b916a9bdf722%28Office.15%29.aspx)|
|[SoundEffect](http://msdn.microsoft.com/library/ea577e7a-32be-ec68-42ab-625816534ab4%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
