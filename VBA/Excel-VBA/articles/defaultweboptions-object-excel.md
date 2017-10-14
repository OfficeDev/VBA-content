---
title: DefaultWebOptions Object (Excel)
keywords: vbaxl10.chm659072
f1_keywords:
- vbaxl10.chm659072
ms.prod: excel
api_name:
- Excel.DefaultWebOptions
ms.assetid: 5bd1d870-e8d9-cac1-d7a7-3aeaf7c4c3cd
ms.date: 06/08/2017
---


# DefaultWebOptions Object (Excel)

Contains global application-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page. You can return or set attributes either at the application (global) level or at the workbook level.


## Remarks

 Workbook-level attribute settings override application-level attribute settings. Workbook-level attributes are contained in the **[WebOptions](weboptions-object-excel.md)** object.


 **Note**  Attribute values can be different from one workbook to another, depending on the attribute value at the time the workbook was saved.


## Example

Use the  **[DefaultWebOptions](application-defaultweboptions-property-excel.md)** property to return the **DefaultWebOptions** object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and sets the _strImageFileType_ variable accordingly.


```vb
Set objAppWebOptions = Application.DefaultWebOptions 
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

