---
title: XlFileValidationPivotMode Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlFileValidationPivotMode
ms.assetid: 8ca2047c-be0f-5ecd-3762-f5c294b9542c
ms.date: 06/08/2017
---


# XlFileValidationPivotMode Enumeration (Excel)

Specifies how to validate the data caches for PivotTable reports.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlFileValidationPivotDefault**|0|Validate the contents of data caches as specified by the  **PivotOptions** registry setting (default).|
| **xlFileValidationPivotRun**|1|Validate the contents of all data caches regardless of the registry setting.|
| **xlFileValidationPivotSkip**|2|Do not validate the contents of data caches.|

## Remarks

This enumeration is used to specify the setting of the  **[FileValidationPivot](application-filevalidationpivot-property-excel.md)** property of the **[Application](application-object-excel.md)** object.

The effect of the  **xlFileValidationPivotDefault** setting is controlled by the `PivotOptions` registry value, which is set in the following registry subkey: `HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Excel\Security\FileValidation`. The  `PivotOptions` value is a **DWORD** value that can be set as listed in the following table.



|**PivotOptions value**|**Description**|
|:-----|:-----|
|0|Never validate PivotTable report data caches. (Not recommended)|
|1|Validate PivotTable report data caches in the following cases (Default setting):<ul><li><p>The file was opened from the Internet.</p></li><li><p>The file is an e-mail attachment.</p></li><li><p>The file was opened by using the <span class="ui">Open in Protected View</span> command of the <span class="ui">Open</span> dialog box.</p></li><li><p>The file was opened from a known unsafe location where Internet content is cached locally, or from a user-defined untrusted location. </p></li><li><p>The data cache is parsed on load when the file is opened.</p></li></ul>|
|2|Validate all PivotTable report data caches.|

