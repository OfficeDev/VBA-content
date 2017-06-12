---
title: Viewer.GetErrorMessage Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.GetErrorMessage
ms.assetid: 31ede4e5-a7ea-c2b8-784e-2e4c7e8bd9ea
ms.date: 06/08/2017
---


# Viewer.GetErrorMessage Method (Visio Viewer)

Returns a string that describes the specified error message code in Microsoft Visio Viewer.


## Syntax

 _expression_. **GetErrorMessage**( **_ErrorCode_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ErrorCode|Required| **Long**|The error message code for which you want to get a description.|

### Return Value

String


## Remarks

If you pass an error code that Visio Viewer does not recognize, the  **GetErrorMessage** method will return either a string saying so, or nothing.

If you pass the value that the  **[LastErrorCode](viewer-lasterrorcode-property-visio-viewer.md)** property returns, the **GetErrorMessage** method returns the last error code that Visio Viewer returned.


## Example

The following code shows how to use the  **GetErrorMessage** method to get a description of the last error code that Visio Viewer returned.


```vb
Debug.Print vsoViewer.GetErrorMessage(vsoViewer.LastErrorCode)
```


