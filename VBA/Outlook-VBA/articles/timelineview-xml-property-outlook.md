---
title: TimelineView.XML Property (Outlook)
keywords: vbaol11.chm2657
f1_keywords:
- vbaol11.chm2657
ms.prod: outlook
api_name:
- Outlook.TimelineView.XML
ms.assetid: 34dee7f8-ee8f-1194-f421-e43fd7815ffe
ms.date: 06/08/2017
---


# TimelineView.XML Property (Outlook)

Returns or sets a  **String** value that specifies the XML definition of the view. Read/write.


## Syntax

 _expression_ . **XML**

 _expression_ A variable that represents a **TimelineView** object.


## Remarks

The XML definition describes the view type by using a series of tags and keywords corresponding to various properties of the view itself. When the view is created, the XML definition is parsed to render the settings for the new view.

To determine how the XML should be structured when creating views, you can create a view by using the Outlook user interface and then you can retrieve the  **XML** property for that view.


## See also


#### Concepts


[TimelineView Object](timelineview-object-outlook.md)

