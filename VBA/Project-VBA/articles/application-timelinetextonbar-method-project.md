---
title: Application.TimelineTextOnBar Method (Project)
keywords: vbapj.chm63
f1_keywords:
- vbapj.chm63
ms.prod: project-server
api_name:
- Project.Application.TimelineTextOnBar
ms.assetid: d57ec0d8-8e35-b6eb-1932-454210bc7dad
ms.date: 06/08/2017
---


# Application.TimelineTextOnBar Method (Project)

Changes the format of text to display as a callout or within the Timeline bar, for one or more selected tasks.


## Syntax

 _expression_. **TimelineTextOnBar**( ** _TextOnBar_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TextOnBar_|Optional|**Boolean**|**False** if the selected tasks should be displayed as callouts; otherwise, **True**. The default value is **True**, which makes the task text show within the Timeline bar.|

### Return Value

 **Boolean**


## Remarks

The  **TimelineTextOnBar** method is equivalent to the **Display as Bar** and **Display as Callout** commands in the **Current Selection** group on the **Format** tab on the ribbon.


## Example

The following statement changes selected tasks on the Timeline bar to display as callouts.


```
TimelineTextOnBar TextOnBar:=False
```


