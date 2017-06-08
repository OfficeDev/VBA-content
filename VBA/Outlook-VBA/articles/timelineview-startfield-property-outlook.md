---
title: TimelineView.StartField Property (Outlook)
keywords: vbaol11.chm2660
f1_keywords:
- vbaol11.chm2660
ms.prod: outlook
api_name:
- Outlook.TimelineView.StartField
ms.assetid: 2477ce1d-a5d0-ddf5-49e9-b25dcd90efbd
ms.date: 06/08/2017
---


# TimelineView.StartField Property (Outlook)

Returns or sets a  **String** value that represents the name of the property that starts the time duration for Outlook items displayed in the **[TimelineView](timelineview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **StartField**

 _expression_ A variable that represents a **TimelineView** object.


## Remarks

The values of the  **StartField** and **[EndField](timelineview-endfield-property-outlook.md)** properties indicate which Outlook item properties are used by the **TimelineView** object to represent the duration of an Outlook item. Both custom and built-in properties can be specified, but only date/time properties are allowed.


## See also


#### Concepts


[TimelineView Object](timelineview-object-outlook.md)

