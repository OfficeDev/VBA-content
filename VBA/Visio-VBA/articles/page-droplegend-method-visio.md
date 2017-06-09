---
title: Page.DropLegend Method (Visio)
keywords: vis_sdr.chm10962175
f1_keywords:
- vis_sdr.chm10962175
ms.prod: visio
api_name:
- Visio.Page.DropLegend
ms.assetid: 8253eafd-4d87-9f1c-833c-cb553c1b73cf
ms.date: 06/08/2017
---


# Page.DropLegend Method (Visio)

Inserts a data graphics legend on a Microsoft Visio drawing page. Returns the list shape instance specified in the  _OuterList_ parameter.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DropLegend**( **_OuterList_** , **_InnerList_** , **_populateFlags_** )

 _expression_ An expression that returns a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OuterList_|Required| **[UNKNOWN]**|A  **[Master](master-object-visio.md)** or **[MasterShortcut](mastershortcut-object-visio.md)** object that represents the legend object. Corresponds to the outermost list shape.|
| _InnerContainer_|Required| **[UNKNOWN]**|A  **Master** or **MasterShortcut** object that represents the legend object. Corresponds to the inner field container shape used within the legend for each data-graphic field.|
| _populateFlags_|Required| **[VisLegendFlags](vislegendflags-enumeration-visio.md)**|A flag that specifies whether Visio should populate the legend.|

### Return Value

 **[Shape](shape-object-visio.md)**


## Remarks

The value of the  _populateFlags_ parameter must be one of the following **VisLegendFlags** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visLegendPopulate**|0|Insert the legend and populate it.|
| **visLegendNoContents**|1|Insert the legend but do not populate it.|
If you pass  **visLegendPopulate** for the _populateFlags_ parameter, Visio inserts the legend and populates it with eligible data-graphic items in use on the specified drawing page. If no such items exist, Visio returns the error EU_API_NOOP, "Operation succeeded but no action taken."

If you pass  **visLegendNoContents** for the _populateFlags_ parameter, Visio inserts a legend that consists of the outer list shape as well as a single inner container shape, which contains an inner list shape but has no heading text.


