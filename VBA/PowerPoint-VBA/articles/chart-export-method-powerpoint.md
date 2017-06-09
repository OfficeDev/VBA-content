---
title: Chart.Export Method (PowerPoint)
keywords: vbapp10.chm684028
f1_keywords:
- vbapp10.chm684028
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Export
ms.assetid: 19b95f24-c262-902e-7e96-c488affeb88d
ms.date: 06/08/2017
---


# Chart.Export Method (PowerPoint)

Exports the chart in a graphic format.


## Syntax

 _expression_. **Export**( **_FileName_**, **_FilterName_**, **_Interactive_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the exported file.|
| _FilterName_|Optional|**Variant**|The language-independent name of the graphic filter as it appears in the registry.|
| _Interactive_|Optional|**Variant**|**True** to display the dialog box that contains the filter-specific options. **False** to indicate that Word should use the default values for the filter. The default is **False**.|

### Return Value

A  **Boolean** value that indicates whether the export was successful.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example exports the first chart in the active document as a GIF file.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Export _
            FileName:="current_sales.gif", FilterName:="GIF"
    End If
End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

