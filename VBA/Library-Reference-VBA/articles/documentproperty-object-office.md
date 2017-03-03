---
title: DocumentProperty Object (Office)
keywords: vbaof11.chm250002
f1_keywords:
- vbaof11.chm250002
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.DocumentProperty
ms.assetid: dd54ca3c-e0e2-4816-539a-17c5b4a928b1
---


# DocumentProperty Object (Office)

Represents a custom or built-in document property of a container document. The  **DocumentProperty** object is a member of the **DocumentProperties** collection.


## Remarks

Use the Microsoft Word  **Document.BuiltinDocumentProperties**( _index_ ) property, where _index_ is the name or index number of the built-in document property, to return a single **DocumentProperty** object that represents a specific built-in document property. Use the Microsoft Word **Document.CustomDocumentProperties**( _index_ ) property, where _index_ is the name or index number of the custom document property, to return a **DocumentProperty** object that represents a specific custom document property. The following list contains the names of all the available built-in document properties:


 **Note**  Properties of type  **msoPropertyTypeString** are limited in length to 255 characters.


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/2a9ac097-0156-007f-2b4b-62a34b240f71%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/7ab10408-c796-92de-8603-ce67c5f0af34%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/ebe1203f-7aed-266e-0701-00da74da7066%28Office.15%29.aspx)|
|[LinkSource](http://msdn.microsoft.com/library/3e3a6ebc-615a-298e-c40f-cbb6d5cf63e3%28Office.15%29.aspx)|
|[LinkToContent](http://msdn.microsoft.com/library/062df6df-cdee-81fc-3244-e229dacaa64e%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/b609c38e-71ca-e019-9852-fc7811dc798f%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/4d6e4c41-09d2-7e0b-c35b-fde629c53c46%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/a6a18498-7a71-b2fb-f037-195bddd70573%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/2d66f8f7-0dfd-e3df-168f-1ca0dfbb0e70%28Office.15%29.aspx)|

## See also


#### Other resources


[DocumentProperty Object Members](http://msdn.microsoft.com/library/568da0ff-fa90-150a-06ec-611de886334e%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
