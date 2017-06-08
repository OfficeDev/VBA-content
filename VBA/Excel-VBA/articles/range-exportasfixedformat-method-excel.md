---
title: Range.ExportAsFixedFormat Method (Excel)
keywords: vbaxl10.chm144246
f1_keywords:
- vbaxl10.chm144246
ms.prod: excel
api_name:
- Excel.Range.ExportAsFixedFormat
ms.assetid: 9786c633-e9bd-3ce3-0246-7bcb3c4b4ce1
ms.date: 06/08/2017
---


# Range.ExportAsFixedFormat Method (Excel)

Exports to a file of the specified format.


## Syntax

 _expression_ . **ExportAsFixedFormat**( **_Type_** , **_Filename_** , **_Quality_** , **_IncludeDocProperties_** , **_IgnorePrintAreas_** , **_From_** , **_To_** , **_OpenAfterPublish_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **XlFixedFormatType**|The type of file format to export to.|
| _Filename_|Optional| **Variant**|The file name of the file to be saved. You can include a full path, or Excel saves the file in the current folder.|
| _Quality_|Optional| **Variant**|Optional  **[XlFixedFormatQuality](xlfixedformatquality-enumeration-excel.md)** . Specifies the quality of the published file.|
| _IncludeDocProperties_|Optional| **Variant**| **True** to include the document properties; otherwise **False** .|
| _IgnorePrintAreas_|Optional| **Variant**| **True** to ignore any print areas set when publishing; otherwise **False** .|
| _From_|Optional| **Variant**|The number of the page at which to start publishing. If this argument is omitted, publishing starts at the beginning.|
| _To_|Optional| **Variant**|The number of the last page to publish. If this argument is omitted, publishing ends with the last page.|
| _OpenAfterPublish_|Optional| **Variant**| **True** to display the file in the viewer after it is published; otherwise **False** .|
| _FixedFormatExtClassPtr_|Optional| **Variant**|Pointer to the  **FixedFormatExt** class.|

## Remarks

 This method also supports initializing an add-in to export a file to a fixed-format file. For example, Excel will perform file format conversion if the converters are present. The conversion is usually initiated by the user.


## See also


#### Concepts


[Range Object](range-object-excel.md)

