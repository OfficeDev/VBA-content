---
title: RecentFiles Object (Excel)
keywords: vbaxl10.chm171072
f1_keywords:
- vbaxl10.chm171072
ms.prod: excel
api_name:
- Excel.RecentFiles
ms.assetid: e33ae942-0444-0631-be08-386366b6ebdb
ms.date: 06/08/2017
---


# RecentFiles Object (Excel)

Represents the list of recently used files.


## Remarks

 Each file is represented by a **[RecentFile](recentfile-object-excel.md)** object.


## Example

Use the  **[RecentFiles](application-recentfiles-property-excel.md)** property to return the **RecentFiles** collection. The following example sets the maximum number of files in the list of recently used files.


```vb
Application.RecentFiles.Maximum = 6
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


