---
title: CustomXMLNodes Object (Office)
keywords: vbaof11.chm293000
f1_keywords:
- vbaof11.chm293000
ms.prod: office
api_name:
- Office.CustomXMLNodes
ms.assetid: 7aa5b7ae-7d4e-4b57-23b5-b027f39e5ff6
ms.date: 06/08/2017
---


# CustomXMLNodes Object (Office)

Contains a collection of  **CustomXMLNodes** objects representing the XML nodes in a document.


## Remarks

The  **Attributes** and the **ChildNodes** properties return collections of nodes of this type.


## Example

The following example selects one or more matching the XPath expression.


```
Sub CustomXmlNodes() 
    Dim cxp1 As CustomXMLPart 
    Dim cxns As CustomXMLNodes 
 
    With ActiveDocument 
  
        ' Returns the first custom xml part with the given root namespace. 
        Set cxp1 = .CustomXMLParts("urn:invoice:namespace")  
         
        ' Get custom xml nodes using XPath.                              
        Set cxns = cxp1.SelectNodes("//*[@unitPrice > 20]")  
                      
    End With 
     
End Sub 

```


## Properties



|**Name**|
|:-----|
|[Application](customxmlnodes-application-property-office.md)|
|[Count](customxmlnodes-count-property-office.md)|
|[Creator](customxmlnodes-creator-property-office.md)|
|[Item](customxmlnodes-item-property-office.md)|
|[Parent](customxmlnodes-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
