---
title: WebBrowserControl.DocumentComplete Event (Access)
keywords: vbaac10.chm143141
f1_keywords:
- vbaac10.chm143141
ms.prod: access
api_name:
- Access.WebBrowserControl.DocumentComplete
ms.assetid: 8cb83f9f-b9c2-8534-8fe3-eb5c56338d6c
ms.date: 06/08/2017
---


# WebBrowserControl.DocumentComplete Event (Access)

Occurs when a document is completely loaded and initialized.


## Syntax

 _expression_. **DocumentComplete**( ** _pDisp_**, ** _URL_** )

 _expression_ A variable that represents a **WebBrowserControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pDisp_|Required|**Object**| pointer to the **IDispatch** interface of the window or frame in which the document is loaded.|
| _URL_|Required|**Variant**|Contains the URL of the loaded document.|

### Return Value

nothing


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

