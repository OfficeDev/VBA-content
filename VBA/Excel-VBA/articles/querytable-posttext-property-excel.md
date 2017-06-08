---
title: QueryTable.PostText Property (Excel)
keywords: vbaxl10.chm518089
f1_keywords:
- vbaxl10.chm518089
ms.prod: excel
api_name:
- Excel.QueryTable.PostText
ms.assetid: f89c21bb-2b51-49b2-b986-8c3aca2038c1
ms.date: 06/08/2017
---


# QueryTable.PostText Property (Excel)

Returns or sets the string used with the post method of inputting data into a Web server to return data from a Web query. Read/write  **String** .


## Syntax

 _expression_ . **PostText**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Microsoft Excel includes sample Web queries that you can modify by changing the HTML code by using WordPad or another text editor. You can find these samples in the Queries folder where you installed Microsoft Office.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **PostText** property applies only to **QueryTable** objects.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

