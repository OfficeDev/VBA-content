---
title: Application.TimelineBarDateRange Method (Project)
keywords: vbapj.chm154
f1_keywords:
- vbapj.chm154
ms.assetid: a1d257f3-92b7-6719-4ce5-5b959823e702
ms.date: 06/08/2017
ms.prod: project-server
---


# Application.TimelineBarDateRange Method (Project)

Modifies the start and finish dates for a  **Timeline** bar. Introduced in Office 2016.


## Syntax

 _expression_. **TimelineBarDateRange**(  _CustomDates_,  _StartDate_,  _FinishDate_,  _BarIndex_)

 _expression_ A variable that represents a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CustomDates_|Optional|Boolean|Indicates if the timeline bar should use custom dates. If true, and the start and finish values are not specified, uses the current project's start and finish dates. If false, ignores any of the other values.|
| _StartValue_|Optional|Variant|Start date.|
| _FinishValue_|Optional|Variant|Finish date.|
| _TimelineViewName_|Optional|String|Specifies the name of a timeline to use. The name can be the built-in timeline or an existing custom timeline such as "My Timeline". The default value is the name of the active timeline.|

### Return Value

 **BOOLEAN**


