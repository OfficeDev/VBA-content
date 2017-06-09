---
title: PublishObjects Object (Excel)
keywords: vbaxl10.chm649072
f1_keywords:
- vbaxl10.chm649072
ms.prod: excel
api_name:
- Excel.PublishObjects
ms.assetid: 33ad393e-5ab6-2531-5e5b-42930fc596c0
ms.date: 06/08/2017
---


# PublishObjects Object (Excel)

A collection of all  **[PublishObject](publishobject-object-excel.md)** objects in the workbook.


## Remarks

 Each **PublishObject** object represents an item in a workbook that has been saved to a Web page and can be refreshed according to values specified by the properties and methods of the object.


## Example

Use the  **[PublishObjects](workbook-publishobjects-property-excel.md)** property to return the **[PublishObjects](publishobjects-object-excel.md)** collection. The following example saves all static **PublishObject** objects in the active workbook to the Web page.


```vb
Set objPObjs = ActiveWorkbook.PublishObjects 
For Each objPO in objPObjs 
 If objPO.HtmlType = xlHTMLStatic Then 
 objPO.Publish 
 End If 
Next objPO
```

Use  **PublishObjects** ( _index_ ), where _index_ is the index number of the specified item in the workbook, to return a single **PublishObject** object. The following example sets the location where the first item in workbook three is saved.




```vb
Workbooks(3).PublishObjects(1).FileName = _ 
 "\\myserver\public\finacct\statemnt.htm"
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

