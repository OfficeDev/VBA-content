---
title: WebBrowserControl.BeforeNavigate2 Event (Access)
keywords: vbaac10.chm143140
f1_keywords:
- vbaac10.chm143140
ms.prod: access
api_name:
- Access.WebBrowserControl.BeforeNavigate2
ms.assetid: 7f6c963b-604e-c350-e71f-899fd6258e46
ms.date: 06/08/2017
---


# WebBrowserControl.BeforeNavigate2 Event (Access)

Occurs before navigation occurs in the given  **WebBrowserControl**.


## Syntax

 _expression_. **BeforeNavigate2**( ** _pDisp_**, ** _URL_**, ** _flags_**, ** _TargetFrameName_**, ** _PostData_**, ** _Headers_**, ** _Cancel_** )

 _expression_ A variable that represents a **WebBrowserControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pDisp_|Required|**Object**|A pointer to the  **IDispatch** interface for the WebBrowser object that represents the window or frame.|
| _URL_|Required|**Variant**|Contains the URL to be navigated to.|
| _flags_|Required|**Variant**|Reserved. Must be set to  **NULL**.|
| _TargetFrameName_|Required|**Variant**|Contains the name of the frame in which to display the resource, or  **NULL** if no named frame is targeted for the resource.|
| _PostData_|Required|**Variant**|Contains the data to send to the server, if the HTTP POST transaction is used.|
| _Headers_|Required|**Variant**|Contains additional HTTP headers to send to the server (HTTP URLs only). The headers can specify information, such as the action required of the server, the type of data being passed to the server, or a status code.|
| _Cancel_|Required|**Boolean**|Contains the cancel flag. Set to  **True** to cancel the navigation operation.|

## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

