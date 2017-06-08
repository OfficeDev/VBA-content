---
title: TimelineView.ShowWeekNumbers Property (Outlook)
keywords: vbaol11.chm2664
f1_keywords:
- vbaol11.chm2664
ms.prod: outlook
api_name:
- Outlook.TimelineView.ShowWeekNumbers
ms.assetid: c4c5a7e5-bc4a-e30a-90c4-89aa3d23368a
ms.date: 06/08/2017
---


# TimelineView.ShowWeekNumbers Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether week number labels are displayed in the timeline scale for the **[TimelineView](timelineview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **ShowWeekNumbers**

 _expression_ A variable that represents a **TimelineView** object.


## Remarks

If this property is set to  **True** , the location in which week number labels are displayed in the timeline scale for the **TimelineView** object depends on the value of the **[TimelineViewMode](timelineview-timelineviewmode-property-outlook.md)** property.



| **Property value**| **Label location**|
| **olTimelineViewDay**|Displayed in the upper portion of the timeline scale, prefacing the date label.|
| **olTimelineViewWeek**|Displayed in the upper portion of the timeline scale, prefacing the week label.|
| **olTimelineViewMonth**|Displayed in the lower portion of the timeline scale, replacing the day and date labels.|

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

