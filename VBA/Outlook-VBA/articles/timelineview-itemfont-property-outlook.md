---
title: TimelineView.ItemFont Property (Outlook)
keywords: vbaol11.chm2669
f1_keywords:
- vbaol11.chm2669
ms.prod: outlook
api_name:
- Outlook.TimelineView.ItemFont
ms.assetid: 7f01e8b1-cd9e-eb19-e481-35b98029320c
ms.date: 06/08/2017
---


# TimelineView.ItemFont Property (Outlook)

Returns a  **[ViewFont](viewfont-object-outlook.md)** object that represents the font used when displaying Outlook items in the **[TimelineView](timelineview-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **ItemFont**

 _expression_ A variable that represents a **TimelineView** object.


## Example

The following Visual Basic for Applications (VBA) sample increments the value of the  **[Size](viewfont-size-property-outlook.md)** property for the **ViewFont** object returned from the **ItemFont** property for the current **TimelineView** object.


```vb
Private Sub IncreaseItemFontSize() 
 
 Dim objTimelineView As TimelineView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTimelineView Then 
 
 
 
 ' Obtain a TimelineView object reference for the 
 
 ' current timeline view. 
 
 Set objTimelineView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Increment the Size property of the 
 
 ' ViewFont object obtained from the 
 
 ' ItemFont property, but only 
 
 ' if the font is less than 24 points 
 
 ' in size. 
 
 If objTimelineView.ItemFont.Size < 24 Then 
 
 objTimelineView.ItemFont.Size = _ 
 
 objTimelineView.ItemFont.Size + 1 
 
 
 
 ' Save the timeline view. 
 
 objTimelineView.Save 
 
 End If 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[TimelineView Object](timelineview-object-outlook.md)

