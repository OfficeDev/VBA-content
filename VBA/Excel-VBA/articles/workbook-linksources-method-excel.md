---
title: Workbook.LinkSources Method (Excel)
keywords: vbaxl10.chm199109
f1_keywords:
- vbaxl10.chm199109
ms.prod: excel
api_name:
- Excel.Workbook.LinkSources
ms.assetid: 6466bea0-5af8-7af0-e9d7-7595133073ae
ms.date: 06/08/2017
---


# Workbook.LinkSources Method (Excel)

Returns an array of links in the workbook. The names in the array are the names of the linked documents, editions, or DDE or OLE servers. Returns  **Empty** if there are no links.


## Syntax

 _expression_ . **LinkSources**( **_Type_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|One of the constants of  **[XlLink](xllink-enumeration-excel.md)** which specifies the type of link to return.|

### Return Value

Variant


## Remarks

The format of the array is a one-dimensional array for all types but publisher and subscriber. The returned strings contain the name of the link source, in the appropriate notation for the link type. For example, DDE links use the "Server|Document!Item" syntax.

For publisher and subscriber links, the returned array is two-dimensional. The first column of the array contains the names of the edition, and the second column contains the references of the editions as text.


## Example

This example displays a list of OLE and DDE links in the active workbook. The example should be run on a workbook that contains one or more linked Word objects.


```vb
aLinks = ActiveWorkbook.LinkSources(xlOLELinks) 
If Not IsEmpty(aLinks) Then 
 For i = 1 To UBound(aLinks) 
 MsgBox "Link " &; i &; ":" &; Chr(13) &; aLinks(i) 
 Next i 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

