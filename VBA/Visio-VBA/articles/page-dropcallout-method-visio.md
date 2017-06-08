---
title: Page.DropCallout Method (Visio)
keywords: vis_sdr.chm10962160
f1_keywords:
- vis_sdr.chm10962160
ms.prod: visio
api_name:
- Visio.Page.DropCallout
ms.assetid: 72edbd4b-e068-6dac-0298-9f746a728892
ms.date: 06/08/2017
---


# Page.DropCallout Method (Visio)

Creates a new callout  **[Shape](shape-object-visio.md)** object on the page near the specified target shape, and associates the callout with the target shape. Returns the callout shape.


## Syntax

 _expression_ . **DropCallout**( **_ObjectToDrop_** , **_TargetShape_** )

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToDrop_|Required| **[UNKNOWN]**|The callout to add to the page. Can be a  **[Master](master-object-visio.md)** , **[MasterShortcut](mastershortcut-object-visio.md)** , **Shape** , or **IDataObject** object.|
| _TargetShape_|Required| **Shape**|The existing shape with which to associate the callout.|

### Return Value

 **Shape**


## Remarks

If the  _ObjectToDrop_ parameter is not a Microsoft Visio object, Visio returns an Invalid Parameter error. If the value you pass is a shape that does not match the context of the method, Visio returns an Invalid Source error.

If the  _TargetShape_ paremeter is null, Visio places the callout shape at the center of the page and does not associate it with any target shapes. If the specified target shapes are not top-level members of the page, Visio returns an Invalid Parameter error.

The  **DropCallout** method corresponds to the **Insert Callout** command in the Visio user interface. (On the **Insert** tab, click **Callout**.)


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **DropCallout** method to add a callout to the active page and associate it with a specific shape.


```vb
Dim vsoDocument As Visio.Document
Set vsoDocument = Application.Documents.OpenEx(Application.GetBuiltInStencilFile(visBuiltInStencilCallouts, visMSUS), visOpenHidden) 
Application.ActivePage.DropCallout vsoDocument.Masters.ItemU("Text callout"), vsoTargetShape
vsoDocument.Close
```


