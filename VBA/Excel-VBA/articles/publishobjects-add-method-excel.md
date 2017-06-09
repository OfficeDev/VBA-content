---
title: PublishObjects.Add Method (Excel)
keywords: vbaxl10.chm650073
f1_keywords:
- vbaxl10.chm650073
ms.prod: excel
api_name:
- Excel.PublishObjects.Add
ms.assetid: 74629499-04d1-11d5-20b8-02b72bb110ee
ms.date: 06/08/2017
---


# PublishObjects.Add Method (Excel)

Creates an object that represents an item in a document saved to a Web page. Such objects facilitate subsequent updates to the Web page while automated changes are being made to the document in Microsoft Excel. Returns a  **PublishObject** object.


## Syntax

 _expression_ . **Add**( **_SourceType_** , **_Filename_** , **_Sheet_** , **_Source_** , **_HtmlType_** , **_DivID_** , **_Title_** )

 _expression_ A variable that represents a **PublishObjects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceType_|Required| **[XlSourceType](xlsourcetype-enumeration-excel.md)**|The source type.|
| _Filename_|Optional| **Variant**| **String** . The URL (on the intranet or the Web) or path (local or network) to which the source object was saved.|
| _Sheet_|Optional| **Variant**|The name of the worksheet that was saved as a Web page.|
| _Source_|Optional| **Variant**|A unique name used to identify items that have one of the following constants as their  _SourceType_ argument: **xlSourceAutoFilter** , **xlSourceChart** , **xlSourcePivotTable** , **xlSourcePrintArea** , **xlSourceQuery** , or **xlSourceRange** . If _SourceType_ is **xlSourceRange** , **Source** specifies a range, which can be a defined name. If _SourceType_ is **xlSourceChart** , **xlSourcePivotTable** , or **xlSourceQuery** , _Source_ specifies the name of a chart, PivotTable report, or query table.|
| _HtmlType_|Optional| **Variant**|Specifies whether the item is saved as an interactive Microsoft Office Web component or as static text and images. Can be one of the  **[XlHTMLType](xlhtmltype-enumeration-excel.md)** constants: **xlHtmlCalc** , **xlHtmlChart** , **xlHtmlList** , or **xlHtmlStatic** .|
| _DivID_|Optional| **Variant**|The unique identifier used in the HTML DIV tag to identify the item on the Web page.|
| _Title_|Optional| **Variant**|The title of the Web page.|

### Return Value

A  **[PublishObject](publishobject-object-excel.md)** object that represents the new item.


## Example

This example saves the range D5:D9 on the First Quarter worksheet in the active workbook to a Web page called  `Stockreport.htm`.


```vb
With ActiveWorkbook.PublishObjects.Add(SourceType:=xlSourceRange, _ 
    Filename:="\\Server\Stockreport.htm", Sheet:="First Quarter", Source:="$G$3:$H$6", _ 
    HtmlType:=xlHtmlStatic, DivID:="Book1_4170") 
        .Publish (True) 
        .AutoRepublish = False 
End With
```


## See also


#### Concepts


[PublishObjects Object](publishobjects-object-excel.md)

