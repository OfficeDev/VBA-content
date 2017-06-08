---
title: Hyperlinks Object (Excel)
keywords: vbaxl10.chm533072
f1_keywords:
- vbaxl10.chm533072
ms.prod: excel
api_name:
- Excel.Hyperlinks
ms.assetid: de28e0af-7a4c-56c3-5fe5-ac47d1654628
ms.date: 06/08/2017
---


# Hyperlinks Object (Excel)

Represents the collection of hyperlinks for a worksheet or range.


## Remarks

 Each hyperlink is represented by a **[Hyperlink](hyperlink-object-excel.md)** object.


## Example

Use the  **[Hyperlinks](worksheet-hyperlinks-property-excel.md)** property to return the **Hyperlinks** collection. The following example checks the hyperlinks on worksheet one for a link that contains the word Microsoft.


```
For Each h in Worksheets(1).Hyperlinks 
 If Instr(h.Name, "Microsoft") <> 0 Then h.Follow 
Next
```

Use the  **[Add](hyperlinks-add-method-excel.md)** method to create a hyperlink and add it to the **Hyperlinks** collection. The following example creates a new hyperlink for cell E5.




```
With Worksheets(1) 
 .Hyperlinks.Add .Range("E5"), "http://example.microsoft.com" 
End With
```


## Methods



|**Name**|
|:-----|
|[Add](hyperlinks-add-method-excel.md)|
|[Delete](hyperlinks-delete-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](hyperlinks-application-property-excel.md)|
|[Count](hyperlinks-count-property-excel.md)|
|[Creator](hyperlinks-creator-property-excel.md)|
|[Item](hyperlinks-item-property-excel.md)|
|[Parent](hyperlinks-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
