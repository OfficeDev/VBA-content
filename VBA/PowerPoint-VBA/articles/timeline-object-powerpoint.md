---
title: TimeLine Object (PowerPoint)
keywords: vbapp10.chm649000
f1_keywords:
- vbapp10.chm649000
ms.prod: powerpoint
api_name:
- PowerPoint.TimeLine
ms.assetid: 0b5a8863-8329-48d0-cb0b-3b34e87acb76
ms.date: 06/08/2017
---


# TimeLine Object (PowerPoint)

Stores animation information for a  **Master**, **Slide**, or **SlideRange** object.


## Example

Use the [TimeLine](http://msdn.microsoft.com/library/f57756b5-9b13-336b-0d5c-00161590ba03%28Office.15%29.aspx)property of the  **[Master](master-object-powerpoint.md)**, **[Slide](slide-object-powerpoint.md)**, or **[SlideRange](http://msdn.microsoft.com/library/440ab59d-744a-209f-bf28-d0acd3a21e1a%28Office.15%29.aspx)** object to return a **TimeLine** object.

The  **TimeLine** object's **[MainSequence](http://msdn.microsoft.com/library/b71f83ad-6d92-cc10-9692-a7567ca0a077%28Office.15%29.aspx)** property gains access to the main animation sequence, while the **[InteractiveSequences](http://msdn.microsoft.com/library/6dbd6b26-6715-e66c-747f-12f1a16416c8%28Office.15%29.aspx)** property gains access to the collection of interactive animation sequences of a slide or slide range.

To reference a timeline object, use syntax similar to these code examples:




```
ActivePresentation.Slides(1).TimeLine.MainSequence

ActivePresentation.SlideMaster.TimeLine.InteractiveSequences

ActiveWindow.Selection.SlideRange.TimeLine.InteractiveSequences
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/ca619c2e-5a15-810f-9441-cf3b17f11ca1%28Office.15%29.aspx)|
|[InteractiveSequences](http://msdn.microsoft.com/library/6dbd6b26-6715-e66c-747f-12f1a16416c8%28Office.15%29.aspx)|
|[MainSequence](http://msdn.microsoft.com/library/b71f83ad-6d92-cc10-9692-a7567ca0a077%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/7c38d6ba-928c-5770-a6f5-9f948a1a50e9%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
