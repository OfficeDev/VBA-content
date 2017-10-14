---
title: Global.RecentFiles Property (Word)
keywords: vbawd10.chm163119111
f1_keywords:
- vbawd10.chm163119111
ms.prod: word
api_name:
- Word.Global.RecentFiles
ms.assetid: e1004877-5fe4-8945-6b7d-8f5279201362
ms.date: 06/08/2017
---


# Global.RecentFiles Property (Word)

Returns a  **RecentFiles** collection that represents the most recently accessed files.


## Syntax

 _expression_ . **RecentFiles**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example opens the first item in the  **RecentFiles** collection (the first document name listed on the **File** menu).


```vb
If RecentFiles.Count >= 1 Then RecentFiles(1).Open
```

This example displays the name of each file in the  **RecentFiles** collection.




```vb
For Each rFile In RecentFiles 
 MsgBox rFile.Name 
Next rFile
```


## See also


#### Concepts


[Global Object](global-object-word.md)

