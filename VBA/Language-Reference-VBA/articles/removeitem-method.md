---
title: RemoveItem Method
keywords: fm20.chm5224968
f1_keywords:
- fm20.chm5224968
ms.prod: office
api_name:
- Office.RemoveItem
ms.assetid: b895775c-7b77-6f2b-b368-998d7114aa7a
ms.date: 06/08/2017
---


# RemoveItem Method



Removes a row from the list in a list box or combo box.
 **Syntax**
 _Boolean_ = _object_. **RemoveItem**_index_
The  **RemoveItem** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. Specifies the row to delete. The number of the first row is 0; the number of the second row is 1, and so on.|
This method will not remove a row from the list if the  **ListBox** is data[bound](glossary-vba.md) (that is, when the **RowSource** property specifies a[data source](glossary-vba.md) for the **ListBox** ).

