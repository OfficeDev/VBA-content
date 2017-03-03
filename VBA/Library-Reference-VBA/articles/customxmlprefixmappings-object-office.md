---
title: CustomXMLPrefixMappings Object (Office)
keywords: vbaof11.chm290000
f1_keywords:
- vbaof11.chm290000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CustomXMLPrefixMappings
ms.assetid: 7da5e1df-a436-ab54-4ea0-270f3edaf240
---


# CustomXMLPrefixMappings Object (Office)

Represents a collection of  **CustomXMLPrefixMapping** objects.


## Example

The following example creates a  **CustomXMLPrefixMapping** object by adding a namespace and prefix to the **CustomXMLPrefixMapping** collection.


```vb
Dim objNamespace As CustomXMLPrefixMapping 
 
objNamespace = CustomXMLPrefixMappings.AddNamespace("xs", "urn:invoice:namespace")
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

