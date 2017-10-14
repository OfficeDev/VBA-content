---
title: VisWebPageSettings.GetPhysicalDimensions Method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.GetPhysicalDimensions
ms.assetid: 879589f5-4b06-df98-c889-ffcf5a4d6419
ms.date: 06/08/2017
---


# VisWebPageSettings.GetPhysicalDimensions Method (Visio Save As Web)

Based on the enumerated screen-resolution value passed to the method in the eRes parameter, places real-world values for screen width and height in pixels in the pnWidth and pnHeight variables passed to the method as parameters.


## Syntax

 _expression_. **GetPhysicalDimensions**( **_eRes_**, **_pnWidth_**, **_pnHeight_**)




### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|eRes|Required| ** [VISWEB_DISP_RES](visweb_disp_res-enumeration-visio-save-as-web.md)**|A screen-resolution value.|
|pnWidth |Required| **Long**|The number of horizontal screen pixels.|
|pnHeight |Required| **Long**|The number of vertical screen pixels.|

### Return Value

 **Nothing**


## Remarks

For example, if you pass in the  **VISWEB_DISP_RES** enumerated screen-resolution value **res1024x768** for eRes, the values 1024 and 768 are returned in pnWidth and pnHeight.


## Example

The following example shows how to use the  **GetPhysicalDimensions** method to determine the screen width and height that correspond to the screen resolution passed to the method as the first parameter.


```vb
Public Sub GetPhysicalDimensions_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 Dim lngWidth As Long 
 Dim lngHeight As Long 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 vsoWebSettings.GetPhysicalDimensions res1280x1024, lngWidth, lngHeight 
 
 Debug.Print lngwidth; lngHeight 
End Sub
```


