---
title: CustomXMLParts Object (Office)
keywords: vbaof11.chm300000
f1_keywords:
- vbaof11.chm300000
ms.prod: office
api_name:
- Office.CustomXMLParts
ms.assetid: 98c1c58e-a08d-6304-8626-1e6705917da3
ms.date: 06/08/2017
---


# CustomXMLParts Object (Office)

Represents a collection of  **CustomXMLPart** objects.


## Remarks

There are three default parts that are always created with a document. These are 'Cover pages', 'Doc properties' and 'App properties'. The last two were in previous versions of Microsoft Word but are now provided in XML form in the  **CustomXMLParts** object collection


## Example

The following example adds a node to a  **CustomXMLPart** object that is part of the **CustomXMLParts** object collection.


```
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[CustomXMLParts Object Members](customxmlparts-members-office.md)

