---
title: ApplicationSettings.SetRasterExportSize Method (Visio)
keywords: vis_sdr.chm16262280
f1_keywords:
- vis_sdr.chm16262280
ms.prod: visio
ms.assetid: 763157d2-014b-0aa4-7c55-a0fb71fb5e23
ms.date: 06/08/2017
---


# ApplicationSettings.SetRasterExportSize Method (Visio)

Sets the raster export size.


## Syntax

 _expression_ . **SetRasterExportSize**( **_size_** , **_[Width]_** , **_[Height]_** , **_[sizeUnits]_** )

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _size_|Required| **VisRasterExportSize**|The raster export size. See Remarks for possible values.|
| _Width_|Optional| **Double**|The raster export size width. Must be greater than or equal to 1.|
| _Height_|Optional| **Double**|The raster export size height. Must be greater than or equal to 1.|
| _sizeUnits_|Optional| **VisRasterExportSizeUnits**|The units used to specify size. See Remarks for possible values.|
| _size_|Required|VISRASTEREXPORTSIZE||
| _Width_|Optional|DOUBLE||
| _Height_|Optional|DOUBLE||
| _sizeUnits_|Optional|VISRASTEREXPORTSIZEUNITS||
|Name|Required/Optional|Data type|Description|

### Return Value

Nothing


## Remarks

The  _size_ parameter must be one of the following **VisRasterExportSize** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterFitToScreenSize**|0|Use screen size.|
| **visRasterFitToPrinterSize**|1|Use printer size.|
| **visRasterFitToSourceSize**|2|Use source size.|
| **visRasterFitToCustomSize**|3|Use custom size.|
If  _size_ is anything other than **visRasterFitToCustomSize** , **SetRasterExportSize** ignores all other parameters.

If  _size_ is **visRasterFitToCustomSize** , **SetRasterExportSize** accepts values for all parameters, if they meet the noted constraints. If they do not meet these constraints, **SetRasterExportSize** returns an Invalid Parameter error.

The  _sizeUnits_ parameter must be one of the following **VisRasterExportSizeUnits** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterPixel**|0|Pixels|
| **visRasterCm**|1|Centimeters|
| **visRasterInch**|2|Inches|
When the  **SetRasterExportSize** method runs successfully, the resulting settings will remain in effect until you either run the method again or change the settings in the user interface.


