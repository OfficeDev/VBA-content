---
title: Shape.MoveToSubprocess Method (Visio)
keywords: vis_sdr.chm11262210
f1_keywords:
- vis_sdr.chm11262210
ms.prod: visio
api_name:
- Visio.Shape.MoveToSubprocess
ms.assetid: 3688c9d5-5b28-23d7-369a-332649267ffe
ms.date: 06/08/2017
---


# Shape.MoveToSubprocess Method (Visio)

Moves the shape to the specified page and drops a replacement shape on the source page, and then links it to the target page. Returns the selection of moved shapes on the target page.

## Syntax

_expression_. **MoveToSubprocess** (**_Page_**, **_ObjectToDrop_**, **_NewShape_**)

_expression_ A variable that represents a **[Shape](shape-object-visio.md)** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Required|**[Page](page-object-visio.md)**|The sub-process page to which the shape should be moved. You cannot pass the current page.|
| _ObjectToDrop_|Required|**[UNKNOWN]**|The replacement shape to drop.|
| _NewShape_|Optional|**Shape**|Out parameter. Returns the shape that is linked to the new page.|

### Return value

 **[Selection](selection-object-visio.md)**

## Remarks

_ObjectToDrop_ is typically a Visio object, such as a **[Master](master-object-visio.md)** or **[MasterShortcut](mastershortcut-object-visio.md)** object. However, it can be any OLE object that provides an **IDataObject** interface. If _ObjectToDrop_ is null, Visio drops a default shape.
