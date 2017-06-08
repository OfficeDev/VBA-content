---
title: Presentation.Export Method (PowerPoint)
keywords: vbapp10.chm583038
f1_keywords:
- vbapp10.chm583038
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Export
ms.assetid: e114d86d-0400-35d3-fc89-d93748993874
ms.date: 06/08/2017
---


# Presentation.Export Method (PowerPoint)

Exports each slide in the presentation, using the specified graphics filter, and saves the exported files in the specified folder.


## Syntax

 _expression_. **Export**( **_Path_**, **_FilterName_**, **_ScaleWidth_**, **_ScaleHeight_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The path of the folder where you want to save the exported slides. You can include a full path; if you don't do this, Microsoft PowerPoint creates a subfolder in the current folder for the exported slides.|
| _FilterName_|Required|**String**|The graphics format in which you want to export slides. The specified graphics format must have an export filter registered in the Windows registry. You can specify either the registered extension or the registered filter name. PowerPoint will first search for a matching extension in the registry. If no extension that matches the specified string is found, PowerPoint will look for a filter name that matches.|
| _ScaleWidth_|Optional|**Long**|The width in pixels of an exported slide.|
| _ScaleHeight_|Optional|**Long**|The height in pixels of an exported slide.|

## Remarks

Exporting a presentation doesn't set the  **[Saved](presentation-saved-property-powerpoint.md)** property of a presentation to **True**.

PowerPoint uses the specified graphics filter to save each individual slide in the presentation. The names of the slides exported and saved to disk are determined by PowerPoint. They're typically saved by using names such as Slide1.wmf, Slide2.wmf. The path of the saved files is specified in the Path argument.


## Example

This example saves the active presentation as a Microsoft PowerPoint presentation and then exports each slide in the presentation as a Portable Network Graphics (PNG) file that will be saved in the Current Work folder. The example also exports each slide with a height of 100 pixels and a width of 100 pixels.


```vb
With ActivePresentation
    .SaveAs FileName:="c:\Current Work\Annual Sales", _
        FileFormat:=ppSaveAsPresentation
    .Export Path:="c:\Current Work", FilterName:="png", _
        ScaleWidth:=100, ScaleHeight:=100
End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

