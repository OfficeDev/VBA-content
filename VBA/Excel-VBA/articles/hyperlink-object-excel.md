---
title: Hyperlink Object (Excel)
keywords: vbaxl10.chm535072
f1_keywords:
- vbaxl10.chm535072
ms.prod: excel
api_name:
- Excel.Hyperlink
ms.assetid: 8bdd2c2f-e6eb-a2f2-78c8-b597aa80ec05
ms.date: 06/08/2017
---


# Hyperlink Object (Excel)

Represents a hyperlink.


## Remarks

 The **Hyperlink** object is a member of the **[Hyperlinks](hyperlinks-object-excel.md)** collection.


## Example

Use the  **[Hyperlink](shape-hyperlink-property-excel.md)** property to return the hyperlink for a shape (a shape can have only one hyperlink). The following example activates the hyperlink for shape one.


```
Worksheets(1).Shapes(1).Hyperlink.Follow NewWindow:=True
```

A range or worksheet can have more than one hyperlink. Use  **[Hyperlinks](worksheet-hyperlinks-property-excel.md)** ( _index_ ), where _index_ is the hyperlink number, to return a single **Hyperlink** object. The folllowing example activates hyperlink two in the range A1:B2.




```
Worksheets(1).Range("A1:B2").Hyperlinks(2).Follow
```


## Methods



|**Name**|
|:-----|
|[AddToFavorites](hyperlink-addtofavorites-method-excel.md)|
|[CreateNewDocument](hyperlink-createnewdocument-method-excel.md)|
|[Delete](hyperlink-delete-method-excel.md)|
|[Follow](hyperlink-follow-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Address](hyperlink-address-property-excel.md)|
|[Application](hyperlink-application-property-excel.md)|
|[Creator](hyperlink-creator-property-excel.md)|
|[EmailSubject](hyperlink-emailsubject-property-excel.md)|
|[Name](hyperlink-name-property-excel.md)|
|[Parent](hyperlink-parent-property-excel.md)|
|[Range](hyperlink-range-property-excel.md)|
|[ScreenTip](hyperlink-screentip-property-excel.md)|
|[Shape](hyperlink-shape-property-excel.md)|
|[SubAddress](hyperlink-subaddress-property-excel.md)|
|[TextToDisplay](hyperlink-texttodisplay-property-excel.md)|
|[Type](hyperlink-type-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
