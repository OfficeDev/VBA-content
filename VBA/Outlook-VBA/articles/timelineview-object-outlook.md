---
title: TimelineView Object (Outlook)
keywords: vbaol11.chm3185
f1_keywords:
- vbaol11.chm3185
ms.prod: outlook
api_name:
- Outlook.TimelineView
ms.assetid: fb14c1a1-f542-fa1e-f30f-c5ee3d2f0206
ms.date: 06/08/2017
---


# TimelineView Object (Outlook)

Represents a view that displays Outlook items in a timeline.


## Remarks

The  **TimelineView** object, derived from the **[View](view-object-outlook.md)** object, allows you to create customizable views that allow you to display Outlook items within a timeline.

Outlook provides several built-in  **TimelineView** objects, and you can also create custom **TimelineView** objects. Use the **[Add](views-add-method-outlook.md)** method of the **[Views](views-object-outlook.md)** collection to add a new **TimelineView** to a **[Folder](folder-object-outlook.md)** object. Use the **[Standard](timelineview-standard-property-outlook.md)** property to determine if an existing **TimelineView** object is built-in or custom.

The  **TimelineView** object supports several different view modes, depending on the desired layout and time period in which to display Outlook items. Use the **[TimelineViewMode](timelineview-timelineviewmode-property-outlook.md)** property to set the view mode, the **[StartField](timelineview-startfield-property-outlook.md)** property to specify the Outlook item property that contains the start date, and the **[EndField](timelineview-endfield-property-outlook.md)** property to specify the Outlook item property that contains the end date for Outlook items to be displayed.

You can configure the appearance of the  **TimelineView**, depending on the view mode. Use the **[ShowWeekNumbers](timelineview-showweeknumbers-property-outlook.md)** property to indicate whether week numbers are displayed in the time scale for the view. Use the **[UpperScaleFont](timelineview-upperscalefont-property-outlook.md)** and **[LowerScaleFont](timelineview-lowerscalefont-property-outlook.md)** properties to specify the font used when displaying, respectively, the upper and lower portions of the time scale for the view.

You can also configure how Outlook items appear within the  **TimelineView** object. Use the **[ItemFont](timelineview-itemfont-property-outlook.md)** property to specify the font used to display Outlook item labels and the **[MaxLabelWidth](timelineview-maxlabelwidth-property-outlook.md)** property to specify the length of labels for Outlook items in the view. Use the **[DefaultExpandCollapseSetting](timelineview-defaultexpandcollapsesetting-property-outlook.md)** property to determine if Outlook items are expanded by default in the view. Use the **[Filter](timelineview-filter-property-outlook.md)** property to determine which Outlook items to display in the view and the **[GroupByFields](timelineview-groupbyfields-property-outlook.md)** collection to specify the Outlook item properties by which Outlook items are grouped in the view. If you set the **TimelineViewMode** to **olTimelineViewMonth**, you can use the **[ShowLabelWhenViewingByMonth](timelineview-showlabelwhenviewingbymonth-property-outlook.md)** property to determine if labels for Outlook items are displayed in the view.

The definition for each  **TimelineView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](timelineview-xml-property-outlook.md)** property to work with the XML definition for the **TimelineView** object.

Use the  **[Apply](timelineview-apply-method-outlook.md)** method to apply any changes made to the **TimelineView** object to the current view. Use the **[Save](timelineview-save-method-outlook.md)** method to persist any changes made to the **TimelineView** object. Use the **[LockUserChanges](timelineview-lockuserchanges-property-outlook.md)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **TimelineView** objects, but you cannot delete them. Use the **[Delete](timelineview-delete-method-outlook.md)** method to delete a custom **TimelineView** object. Use the **[Reset](timelineview-reset-method-outlook.md)** method to reset the properties of a built-in **TimelineView** object to their default values.


## Example

The following Visual Basic for Applications (VBA) example configures the current  **TimelineView** object to display Outlook items by month, with week number labels on the lower portion of the timeline scale, with labels no longer than 40 characters.


```
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


## Methods



|**Name**|
|:-----|
|[Apply](timelineview-apply-method-outlook.md)|
|[Copy](timelineview-copy-method-outlook.md)|
|[Delete](timelineview-delete-method-outlook.md)|
|[GoToDate](timelineview-gotodate-method-outlook.md)|
|[Reset](timelineview-reset-method-outlook.md)|
|[Save](timelineview-save-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](timelineview-application-property-outlook.md)|
|[Class](timelineview-class-property-outlook.md)|
|[DefaultExpandCollapseSetting](timelineview-defaultexpandcollapsesetting-property-outlook.md)|
|[EndField](timelineview-endfield-property-outlook.md)|
|[Filter](timelineview-filter-property-outlook.md)|
|[GroupByFields](timelineview-groupbyfields-property-outlook.md)|
|[ItemFont](timelineview-itemfont-property-outlook.md)|
|[Language](timelineview-language-property-outlook.md)|
|[LockUserChanges](timelineview-lockuserchanges-property-outlook.md)|
|[LowerScaleFont](timelineview-lowerscalefont-property-outlook.md)|
|[MaxLabelWidth](timelineview-maxlabelwidth-property-outlook.md)|
|[Name](timelineview-name-property-outlook.md)|
|[Parent](timelineview-parent-property-outlook.md)|
|[SaveOption](timelineview-saveoption-property-outlook.md)|
|[Session](timelineview-session-property-outlook.md)|
|[ShowLabelWhenViewingByMonth](timelineview-showlabelwhenviewingbymonth-property-outlook.md)|
|[ShowWeekNumbers](timelineview-showweeknumbers-property-outlook.md)|
|[Standard](timelineview-standard-property-outlook.md)|
|[StartField](timelineview-startfield-property-outlook.md)|
|[TimelineViewMode](timelineview-timelineviewmode-property-outlook.md)|
|[UpperScaleFont](timelineview-upperscalefont-property-outlook.md)|
|[ViewType](timelineview-viewtype-property-outlook.md)|
|[XML](timelineview-xml-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
