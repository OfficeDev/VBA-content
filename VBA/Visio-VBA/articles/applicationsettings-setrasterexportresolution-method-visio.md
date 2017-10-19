---
title: ApplicationSettings.SetRasterExportResolution Method (Visio)
keywords: vis_sdr.chm16262265
f1_keywords:
- vis_sdr.chm16262265
ms.prod: visio
ms.assetid: 18b28fe1-4460-940c-0de7-566a608a8f04
ms.date: 06/08/2017
---


# ApplicationSettings.SetRasterExportResolution Method (Visio)

Specifies the raster export resolution settings.

## Syntax

_expression_. **SetRasterExportResolution** (**_resolution_**, **_Width_**, **_Height_**, **_resolutionUnits_**)

_expression_ An expression that returns an **[ApplicationSettings](applicationsettings-object-visio.md)** object.

### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _resolution_|Required|**VisRasterExportResolution**|The raster export resolution. For possible values, see [Remarks](#remarks).|
| _Width_|Optional|**Double**|The raster export resolution width. Must be greater than or equal to 1.|
| _Height_|Optional|**Double**|The raster export resolution height. Must be greater than or equal to 1.|
| _resolutionUnits_|Optional|**VisRasterExportResolutionUnits**|The units used to specify resolution. For possible values, see [Remarks](#remarks).|

<br/>

### Return value

Nothing

## Remarks

The _resolution_ parameter must be one of the following **VisRasterExportResolution** constants.

|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**visRasterUseScreenResolution**|0|Use screen resolution.|
|**visRasterUsePrinterResolution**|1|Use printer resolution.|
|**visRasterUseSourceResolution**|2|Use source resolution.|
|**visRasterUseCustomResolution**|3|Use custom resolution.|

<br/>

If _resolution_ is a constant other than **visRasterUseCustomResolution**, **SetRasterExportResolution** ignores all other parameters.

If _resolution_ is **visRasterUseCustomResolution**, **SetRasterExportResolution** accepts values for all parameters if they meet the noted constraints. If they do not meet these constraints, **SetRasterExportResolution** returns an Invalid Parameter error.

The _resolutionUnits_ parameter must be one of the following **VisRasterExportResolutionUnits** constants.

|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**visRasterPixelsPerInch**|0|Pixels per inch.|
|**visRasterPixelsPerCm**|1|Pixels per centimeter.|

<br/>

When the **SetRasterExportResolution** method runs successfully, the resulting settings will remain in effect until you either run the method again or change the settings in the user interface.
