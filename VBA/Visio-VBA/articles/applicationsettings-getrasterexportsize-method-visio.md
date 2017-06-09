---
title: ApplicationSettings.GetRasterExportSize Method (Visio)
keywords: vis_sdr.chm16262285
f1_keywords:
- vis_sdr.chm16262285
ms.prod: visio
ms.assetid: 70591d2c-ac80-5637-996e-3ebef6be0c51
ms.date: 06/08/2017
---


# ApplicationSettings.GetRasterExportSize Method (Visio)

Gets the raster export size.


## Syntax

 _expression_ . **GetRasterExportSize**( **_pSize_** , **_pWidth_** , **_pHeight_** , **_pSizeUnits_** )

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pSize_|Required| **VisRasterExportSize**|Out parameter. The raster export size. See Remarks for possible values.|
| _pWidth_|Required| **Double**|Out parameter. The raster export size width. |
| _pHeight_|Required| **Double**|Out parameter. The raster export size height.|
| _pSizeUnits_|Required| **VisRasterExportSizeUnits**|Out parameter. The units used to specify size. See Remarks for possible values.|
| _pSize_|Required|VISRASTEREXPORTSIZE||
| _pWidth_|Required|DOUBLE||
| _pHeight_|Required|DOUBLE||
| _pSizeUnits_|Required|VISRASTEREXPORTSIZEUNITS||
|Name|Required/Optional|Data type|Description|

### Return Value

Nothing


## Remarks

The  _pSize_ parameter must be one of the following **VisRasterExportSize** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterFitToScreenSize**|0|Use screen size.|
| **visRasterFitToPrinterSize**|1|Use printer size.|
| **visRasterFitToSourceSize**|2|Use source size.|
| **visRasterFitToCustomSize**|3|Use custom size.|
If  _pSize_ is a constant other than **visRasterFitToCustomSize** , **GetRasterExportSize** returns null for all other parameters. If _pSize_ is **visRasterFitToCustomSize** , **GetRasterExportSize** returns non-null values for all parameters.

The  _pSizeUnits_ parameter must be one of the following **VisRasterExportSizeUnits** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterPixel**|0|Pixels|
| **visRasterCm**|1|Centimeters|
| **visRasterInch**|2|Inches|

