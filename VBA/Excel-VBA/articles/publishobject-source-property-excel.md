---
title: PublishObject.Source Property (Excel)
keywords: vbaxl10.chm652078
f1_keywords:
- vbaxl10.chm652078
ms.prod: excel
api_name:
- Excel.PublishObject.Source
ms.assetid: 2f8ca565-91f1-9636-d0c2-f5988c176ddb
ms.date: 06/08/2017
---


# PublishObject.Source Property (Excel)

Returns a  **Variant** value that represents the unique name that identifies items that have a **[SourceType](publishobject-sourcetype-property-excel.md)** property value of **xlSourceRange** , **xlSourceChart** , **xlSourcePrintArea** , **xlSourceAutoFilter** , **xlSourcePivotTable** , or **xlSourceQuery** .


## Syntax

 _expression_ . **Source**

 _expression_ A variable that represents a **PublishObject** object.


## Remarks

If the  **SourceType** property is set to **xlSourceRange** , this property returns a range, which can be a defined name. If the **SourceType** property is set to **xlSourceChart** , **xlSourcePivotTable** , or **xlSourceQuery** , this property returns the name of the object, such as a chart name, a PivotTable report name, or a query table name.


## Example

This example determines the unique name of the first chart (in the first workbook) saved as a Web page, and then it sets the  **Boolean** variable `blnChartFound` to **True** . If no items in the document have been saved as Chart components, `blnChartFound` is **False** .


```vb
blnChartFound = False 
For Each objPO In Workbooks(1).PublishObjects 
 If objPO.SourceType = xlSourceChart Then 
 strFirstPO = objPO.Source 
 blnChartFound = True 
 Exit For 
 End If 
Next objPO
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-excel.md)

