---
title: TimelineView.MaxLabelWidth Property (Outlook)
keywords: vbaol11.chm2665
f1_keywords:
- vbaol11.chm2665
ms.prod: outlook
api_name:
- Outlook.TimelineView.MaxLabelWidth
ms.assetid: b97e4104-89d8-c8a6-598e-7397cf47f320
ms.date: 06/08/2017
---


# TimelineView.MaxLabelWidth Property (Outlook)

Returns or sets a  **Long** value that represents the maximum length (in characters) for the label of an Outlook item in the **[TimelineView](timelineview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **MaxLabelWidth**

 _expression_ A variable that represents a **TimelineView** object.


## Remarks

This property can be set to a value between 0 and 132. If this property is set to a value less than 0, the property is set to 0. If this property is set to a value greater than 132, the property is set to 132. The default value for this property is 80.

If this property is set to 0, labels for Outlook items are not displayed in the view.


## Example

The following Visual Basic for Applications (VBA) example configures the current  **TimelineView** object to display Outlook items by month, with week number labels on the lower portion of the timeline scale, with labels no longer than 40 characters.


```vb
Private Sub ConfigureMonthTimelineView() 
 
 Dim objTimelineView As TimelineView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTimelineView Then 
 
 
 
 ' Obtain a TimelineView object reference for the 
 
 ' current timeline view. 
 
 Set objTimelineView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Configure the TimelineView object so that it displays 
 
 ' Outlook items by month and week, displaying labels 
 
 ' no larger than 40 characters for Outlook items 
 
 ' displayed in the view. 
 
 With objTimelineView 
 
 ' Display items by month. 
 
 .TimelineViewMode = olTimelineViewMonth 
 
 ' Display week numbers. If this value is 
 
 ' set to False when TimelineViewMode is 
 
 ' set to olTimelineViewMonth, the day 
 
 ' numbers are displayed instead. 
 
 .ShowWeekNumbers = True 
 
 ' Display labels for Outlook items 
 
 ' while TimelineViewMode is set to 
 
 ' olTimelineViewMonth. 
 
 .ShowLabelWhenViewingByMonth = True 
 
 ' Show no more than the first 40 characters 
 
 ' for each Outlook item in the view. 
 
 .MaxLabelWidth = 40 
 
 
 
 ' Save and apply the view. 
 
 .Save 
 
 .Apply 
 
 End With 
 
 End If 
 
 
 
End Sub
```


## See also


#### Concepts


[TimelineView Object](timelineview-object-outlook.md)

