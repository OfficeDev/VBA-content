---
title: Application.TimelineExport Method (Project)
keywords: vbapj.chm66
f1_keywords:
- vbapj.chm66
ms.prod: project-server
api_name:
- Project.Application.TimelineExport
ms.assetid: a2829e86-5b83-0076-33a3-4c10040ffc17
ms.date: 06/08/2017
---


# Application.TimelineExport Method (Project)

Copies an image of the active Timeline view to the Clipboard, for pasting into other applications.


## Syntax

 _expression_. **TimelineExport**( ** _SelectionOnly_**, ** _ExportWidth_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SelectionOnly_|Optional|**Boolean**|**True** if the exported timeline includes only the selected items; otherwise, **False**. The default is **False**, where the entire visible timeline is exported.|
| _ExportWidth_|Optional|**Long**|Specifies the approximate width in pixels of the exported timeline. Values can effectively range from 100 to 10,000. The default value is the actual width of the Timeline pane in Project.|

### Return Value

 **Boolean**


## Remarks


 **Note**  The Timeline view must be selected.

Selecting a task in the Gantt chart does not select the same task on the timeline. To select items for export, click or control-click them on the timeline.

The  **TimelineExport** method can duplicate commands in the **Copy Timeline** drop-down menu on the **Format** tab on the ribbon, when the Timeline pane is selected. If the ExportWidth argument is specified, the size of the copied image is based on the value of ExportWidth, not on the size of the Project window or the Timeline pane.

Values of ExportWidth are limited to a range of 100 to 10000. Values outside that range are changed to 100 or 10000. The actual width of the image is less than ExportWidth. For example, if the value of ExportWidth is 10000, the actual width is 9957 pixels.


## Example

The following statement corresponds to the  **Full Size** command in the **Copy Timeline** drop-down menu. The actual width of the exported image is the width of the Timeline pane.


```
TimelineExport
```

The following statement corresponds to the  **For Presentation** command in the **Copy Timeline** drop-down menu. The actual width of the exported image is 891 pixels.




```
TimelineExport ExportWidth:=916
```

The following statement corresponds to the  **For E-mail** command in the **Copy Timeline** drop-down menu. The actual width of the exported image is 554 pixels.




```
TimelineExport ExportWidth:=600
```


