---
title: PivotTable.CreateCubeFile Method (Excel)
keywords: vbaxl10.chm235152
f1_keywords:
- vbaxl10.chm235152
ms.prod: excel
api_name:
- Excel.PivotTable.CreateCubeFile
ms.assetid: 585641a1-c708-75fd-4789-f7a254830b57
ms.date: 06/08/2017
---


# PivotTable.CreateCubeFile Method (Excel)

Creates a cube file from a PivotTable report connected to an Online Analytical Processing (OLAP) data source.


## Syntax

 _expression_ . **CreateCubeFile**( **_File_** , **_Measures_** , **_Levels_** , **_Members_** , **_Properties_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _File_|Required| **String**|The name of the cube file to be created. It will overwrite the file if it already exists.|
| _Measures_|Optional| **Variant**|An array of unique names of measures that are to be part of the slice.|
| _Levels_|Optional| **Variant**|An array of strings. Each array item is a unique level name. It represents the lowest level of a hierarchy that is in the slice.|
| _Members_|Optional| **Variant**|An array of string arrays. The elements correspond, in order, to the hierarchies represented in the  _Levels_ array. Each element is an array of string arrays that consists of the unique names of the top level members in the dimension that are to be included in the slice.|
| _Properties_|Optional| **Variant**| **False** results in no member properties being included in the slice. The default value is **True** .|

### Return Value

String


## Example

This example creates a cube file titled "CustomCubeFile" on drive C:\ with no member properties to be included in the slice. With the  _Measures_,  _Levels_, and  _Members_ arguments omitted from this example, the cube file will end up matching the view of the PivotTable report. This example assumes a PivotTable report connected to an OLAP data source exists on the active worksheet.


```vb
Sub UseCreateCubeFile() 
 
 ActiveSheet.PivotTables(1).CreateCubeFile _ 
 File:="C:\CustomCubeFile", Properties:=False 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

