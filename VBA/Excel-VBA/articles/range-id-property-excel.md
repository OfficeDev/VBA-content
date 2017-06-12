---
title: Range.ID Property (Excel)
keywords: vbaxl10.chm144231
f1_keywords:
- vbaxl10.chm144231
ms.prod: excel
api_name:
- Excel.Range.ID
ms.assetid: 0ff7f261-8829-2858-5097-a638c01e5f3c
ms.date: 06/08/2017
---


# Range.ID Property (Excel)

Returns or sets a  **String** value that represents the identifying label for the specified cell when the page is saved as a Web page.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents a **Range** object.


## Remarks

You can use an ID label as a hyperlink reference in other HTML documents or on the same Web page.


## Example

This example sets the ID of cell A1 on the active worksheet to "target".


```vb
ActiveSheet.Range("A1").ID = "target"
```

Later, the document is saved as a Web page, and the following line of HTML is added to the Web page. When the user then views the page in a Web browser and clicks the hyperlink, the browser displays the cell.




```
<A HREF="#target">Quarterly earnings</A>
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

