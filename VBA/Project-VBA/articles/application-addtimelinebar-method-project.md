---
title: Application.AddTimelineBar Method (Project)
keywords: vbapj.chm157
f1_keywords:
- vbapj.chm157
ms.assetid: 2cb9d639-3363-79e3-ced6-73b0a574986a
ms.date: 06/08/2017
ms.prod: project-server
---


# Application.AddTimelineBar Method (Project)

Adds a  **timeline** bar to the view. Introduced in Office 2016.


## Syntax

 _expression_. **AddTimelineBar**

 _expression_ A variable that represents a **Application** object.


### Parameters


|||||
|:-----|:-----|:-----|:-----|
|**Name**|**Required/Optional**|**Data Type**|**Description**|
| _BarPosition_|Optional|Integer|Indicates where to add the timeline bar. If a number isn't specified, it is added at the bottom. The top bar is 0 and the next is 1, and so on. |
| _TimelineViewName_|Optional|String|Specifies the name of a timeline to use. The name can be the built-in timeline or an existing custom timeline such as "My Timeline". The default value is the name of the active timeline.|

### Return Value

 **BOOLEAN**


