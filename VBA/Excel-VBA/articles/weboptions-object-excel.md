---
title: WebOptions Object (Excel)
keywords: vbaxl10.chm661072
f1_keywords:
- vbaxl10.chm661072
ms.prod: excel
api_name:
- Excel.WebOptions
ms.assetid: d573637f-1891-4602-c961-091795e47356
ms.date: 06/08/2017
---


# WebOptions Object (Excel)

Contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.


## Remarks

 You can return or set attributes either at the application (global) level or at the workbook level. (Note that attribute values can be different from one workbook to another, depending on the attribute value at the time the workbook was saved.) Workbook-level attribute settings override application-level attribute settings. Application-level attributes are contained in the **[DefaultWebOptions](defaultweboptions-object-excel.md)** object.


## Example

Use the  **[WebOptions](workbook-weboptions-property-excel.md)** property to return the **WebOptions** object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and then sets the `strImageFileType` variable accordingly.


```vb
Set objAppWebOptions = Workbooks(1).WebOptions 
With objAppWebOptions 
 If .AllowPNG = True Then 
 strImageFileType = "PNG" 
 Else 
 strImageFileType = "JPG" 
 End If 
End With
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

