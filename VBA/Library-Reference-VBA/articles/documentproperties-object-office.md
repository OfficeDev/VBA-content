---
title: DocumentProperties Object (Office)
keywords: vbaof11.chm250010
f1_keywords:
- vbaof11.chm250010
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.DocumentProperties
ms.assetid: 90d42786-7d9a-b604-dbdf-88db41cbe69b
---


# DocumentProperties Object (Office)

A collection of  **DocumentProperty** objects. Each **DocumentProperty** object represents a built-in or custom property of a container document.


## Remarks

Use the ** Add** method to create a new custom property and add it to the **DocumentProperties** collection. You cannot use the **Add** method to create a built-in document property.

Use  **BuiltinDocumentProperties(index)**, where _index_ is the index number of the built-in document property, to return a single **DocumentProperty** object that represents a specific built-in document property. Use **CustomDocumentProperties(index)**, where _index_ is the number of the custom document property, to return a **DocumentProperty** object that represents a specific custom document property.


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/80738562-8b0b-33f1-3dfa-0d66b1844ef7%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/b91998a4-f933-d584-8293-e63ad82447e2%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/8f4367bd-d30a-ba45-3ec2-3c5b94ede4d8%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/47ba7f73-b72e-2990-d35d-cd73b08b91cd%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/33649822-adc5-5efd-7e05-87735b30b19f%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e1239ffa-b89e-e78f-4009-d576c473d477%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[DocumentProperties Object Members](http://msdn.microsoft.com/library/bb388713-3029-796e-3328-6193eb14d1bf%28Office.15%29.aspx)
