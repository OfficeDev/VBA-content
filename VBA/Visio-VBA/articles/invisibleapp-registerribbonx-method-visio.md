---
title: InvisibleApp.RegisterRibbonX Method (Visio)
keywords: vis_sdr.chm17562090
f1_keywords:
- vis_sdr.chm17562090
ms.prod: visio
api_name:
- Visio.InvisibleApp.RegisterRibbonX
ms.assetid: db9f5050-0813-f805-5e1c-6fe141742dbe
ms.date: 06/08/2017
---


# InvisibleApp.RegisterRibbonX Method (Visio)

Registers the  **[IRibbonExtensibility](http://msdn.microsoft.com/library/b27a7576-b6f5-031e-e307-78ef5f8507e0%28Office.15%29.aspx)** interface that is implemented by the specified add-on to populate the custom user interface (UI).


## Syntax

 _expression_ . **RegisterRibbonX**( **_SourceAddOn_** , **_TargetDocument_** , **_TargetModes_** , **_FriendlyName_** )

 _expression_ A variable that represents an **[InvisibleApp](invisibleapp-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceAddOn_|Required| **IRibbonExtensibility**|The add-on to register.|
| _TargetDocument_|Required| **[Document](document-object-visio.md)**|The document that uses the custom UI.|
| _TargetModes_|Required| **[VisRibbonXModes](visribbonxmodes-enumeration-visio.md)**|The modes in which the custom UI should be visible. See Remarks for possible values.|
| _FriendlyName_|Required| **String**|The name to associate with the UI items and errors that originate in the add-on.|

### Return Value

 **HRESULT**


## Remarks

If  _TargetDocument_ is null, the custom UI is defined at the level of the application. Otherwise, Microsoft Visio binds the visibility of the custom UI to the specified document. The UI does not appear in conjunction with any other document.

 _TargetModes_ can be one or more of the following **VisRibbonXModes** constants.



|**Names**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRxModeNone**|0|Display the custom UI when no documents are open.|
| **visRxModeDrawing**|1|Display the custom UI in Drawing mode.|
| **visRxModeStencil**|2|Display the custom UI in Stencil mode.|
| **visRxModePrintPreview**|4|Display the custom UI in Print Preview mode.|
If  _FriendlyName_ is null, the method fails.


