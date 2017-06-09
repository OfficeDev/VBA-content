---
title: ListDataFormat.Choices Property (Excel)
keywords: vbaxl10.chm758074
f1_keywords:
- vbaxl10.chm758074
ms.prod: excel
api_name:
- Excel.ListDataFormat.Choices
ms.assetid: c4a809e6-7977-28a1-1070-286e7df99409
ms.date: 06/08/2017
---


# ListDataFormat.Choices Property (Excel)

 Returns an **Array** of **String** values that contains the choices offered to the user by the **ListLookUp** , **ChoiceMulti** , and **Choice** data types of the **[DefaultValue](listdataformat-defaultvalue-property-excel.md)** property. Read-only **Variant** .


## Syntax

 _expression_ . **Choices**

 _expression_ A variable that represents a **ListDataFormat** object.


## Remarks

In Microsoft Excel, you cannot set any of the properties associated with the  **ListDataFormat** object. You can set these properties, however, by modifying the list on the server that is running Microsoft SharePoint Foundation.


## Example

The following example displays the setting of the  **Choice** property for the third column in a list that is linked to a SharePoint list. In this example, it is assumed that the **DefaultValue** property has been set to the **Choice** , **ChoiceMulti** , or **ListLookup** data type.


```vb
Sub PrintChoices() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.Choices 
End Sub
```


## See also


#### Concepts


[ListDataFormat Object](listdataformat-object-excel.md)

