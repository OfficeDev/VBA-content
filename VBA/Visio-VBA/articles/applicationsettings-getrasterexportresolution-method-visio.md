---
title: ApplicationSettings.GetRasterExportResolution Method (Visio)
keywords: vis_sdr.chm16262270
f1_keywords:
- vis_sdr.chm16262270
ms.prod: visio
ms.assetid: 526d2970-006b-6596-bfef-49446dd58610
ms.date: 06/08/2017
---


# ApplicationSettings.GetRasterExportResolution Method (Visio)

Returns the raster export resolution settings.


## Syntax

 _expression_ . **GetRasterExportResolution**( **_pResolution_** , **_pWidth_** , **_pHeight_** , **_pResolutionUnits_** )

 _expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _pResolution_|Required| **VisRasterExportResolution**|Out parameter. The raster export resolution. See Remarks for possible values.|
| _pWidth_|Required| **Double**|Out parameter. The raster export resolution width.|
| _pHeight_|Required| **Double**|Out parameter. The raster export resolution height.|
| _pResolutionUnits_|Required| **VisRasterExportResolutionUnits**|Out parameter. The units used to specify resolution. See Remarks for possible values.|
| _pResolution_|Required|VISRASTEREXPORTRESOLUTION||
| _pWidth_|Required|DOUBLE||
| _pHeight_|Required|DOUBLE||
| _pResolutionUnits_|Required|VISRASTEREXPORTRESOLUTIONUNITS||
|Name|Required/Optional|Data type|Description|

### Return Value

Nothing


## Remarks

The  _pResolution_ parameter must be one of the following **VisRasterExportResolution** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterUseScreenResolution**|0|Use screen resolution.|
| **visRasterUsePrinterResolution**|1|Use printer resolution.|
| **visRasterUseSourceResolution**|2|Use source resolution.|
| **visRasterUseCustomResolution**|3|Use custom resolution.|
If  _pResolution_ is a constant other than **visRasterUseCustomResolution** , **GetRasterExportResolution** returns null for all other parameters. If _pResolution_ is **visRasterUseCustomResolution** , **GetRasterExportResolution** returns non-null values for all parameters.

The  _pResolutionUnits_ parameter must be one of the following **VisRasterExportResolutionUnits** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRasterPixelsPerInch**|0|Pixels per inch.|
| **visRasterPixelsPerCm**|1|Pixels per centimeter.|

