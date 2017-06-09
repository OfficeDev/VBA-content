---
title: PublishObject.SourceType Property (Excel)
keywords: vbaxl10.chm652077
f1_keywords:
- vbaxl10.chm652077
ms.prod: excel
api_name:
- Excel.PublishObject.SourceType
ms.assetid: 4d22915d-c5a3-c06f-85dc-3c6394644cec
ms.date: 06/08/2017
---


# PublishObject.SourceType Property (Excel)

Returns a  **[XlSourceType](xlsourcetype-enumeration-excel.md)** value that represents the type of item being published.


## Syntax

 _expression_ . **SourceType**

 _expression_ A variable that represents a **PublishObject** object.


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

