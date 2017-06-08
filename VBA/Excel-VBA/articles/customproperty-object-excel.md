---
title: CustomProperty Object (Excel)
keywords: vbaxl10.chm681072
f1_keywords:
- vbaxl10.chm681072
ms.prod: excel
api_name:
- Excel.CustomProperty
ms.assetid: df8b58d8-ccfd-00bb-723a-a9c328f0b38b
ms.date: 06/08/2017
---


# CustomProperty Object (Excel)

Represents identifier information. Identifier information can be used as metadata for XML.


## Remarks

Use the  **[Add](customproperties-add-method-excel.md)** method or the **[Item](customproperties-item-property-excel.md)** property of the **[CustomProperties](customproperties-object-excel.md)** collection to return a **CustomProperty** object.

Once a  **CustomProperty** object is returned, you can add metadata to worksheets using the **[CustomProperties](worksheet-customproperties-property-excel.md)** property with the **Add** method.


## Example

In this example, Microsoft Excel adds identifier information to the active worksheet and returns the name and value to the user.


```vb
Sub CheckCustomProperties() 
 
 Dim wksSheet1 As Worksheet 
 
 Set wksSheet1 = Application.ActiveSheet 
 
 ' Add metadata to worksheet. 
 wksSheet1.CustomProperties.Add _ 
 Name:="Market", Value:="Nasdaq" 
 
 ' Display metadata. 
 With wksSheet1.CustomProperties.Item(1) 
 MsgBox .Name &; vbTab &; .Value 
 End With 
 
End Sub
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

