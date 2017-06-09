---
title: WebBrowserControl.NavigateError Event (Access)
keywords: vbaac10.chm143143
f1_keywords:
- vbaac10.chm143143
ms.prod: access
api_name:
- Access.WebBrowserControl.NavigateError
ms.assetid: 1b94a46a-b423-81e7-13df-e2d24434f0df
ms.date: 06/08/2017
---


# WebBrowserControl.NavigateError Event (Access)

Occurs when an error occurs during navigation.


## Syntax

 _expression_. **NavigateError**( ** _pDisp_**, ** _URL_**, ** _TargetFrameName_**, ** _StatusCode_**, ** _Cancel_** )

 _expression_ A variable that represents a **WebBrowserControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pDisp_|Required|**Object**|A pointer to an  **IDispatch** interface for the WebBrowser object that represents the window or frame in which the navigation error occurred.|
| _URL_|Required|**Variant**|Contains the URL for which navigation failed.|
| _TargetFrameName_|Required|**Variant**|Contains the name of the frame in which to display the resource, or  **NULL** if no named frame was targeted for the resource.|
| _StatusCode_|Required|**Variant**|Contains an error status code, if available.|
| _Cancel_|Required|**Boolean**|Specifies whether to cancel the navigation to an error page or to any further autosearch.|

### Return Value

nothing


## Remarks

This event fires before the  **WebBrowser** control displays an error page due to an error in navigation. You can stop the display of the error page by setting the _Cancel_ parameter to **True**. However, if the server contacted in the original navigation supplies its own substitute page navigation, when you set _Cancel_ to **True**, it has no effect, and the navigation to the server's alternate page proceeds. For example, assume that a navigation to http://www.www.wingtiptoys.com/BigSale.htm causes this event to fire because the page does not exist. However, the server is set to redirect the navigation to http://www.www.wingtiptoys.com/home.htm. In this case, when you set _Cancel_ to **True**, it has no effect, and navigation proceeds to http://www.www.wingtiptoys.com/home.htm.


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

