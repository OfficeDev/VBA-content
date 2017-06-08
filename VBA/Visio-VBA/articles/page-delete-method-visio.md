---
title: Page.Delete Method (Visio)
keywords: vis_sdr.chm10951185
f1_keywords:
- vis_sdr.chm10951185
ms.prod: visio
api_name:
- Visio.Page.Delete
ms.assetid: 7adc0e81-7000-2bfa-cca5-c74c3fcbac5c
ms.date: 06/08/2017
---


# Page.Delete Method (Visio)

Deletes a  **Page** object. Can also renumber remaining pages.


## Syntax

 _expression_ . **Delete**( **_fRenumberPages_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fRenumberPages_|Required| **Integer**|1 ( **True** ) to renumber remaining pages; otherwise, 0 ( **False** ).|

### Return Value

Nothing


## Remarks

When  _fRenumberPages_ is non-zero, the remaining pages' default page names are renumbered after the page is deleted, otherwise, the pages retain their names.


