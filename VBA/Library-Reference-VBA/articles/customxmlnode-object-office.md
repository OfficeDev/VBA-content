---
title: CustomXMLNode Object (Office)
keywords: vbaof11.chm294000
f1_keywords:
- vbaof11.chm294000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CustomXMLNode
ms.assetid: e90213f5-6d62-52d8-3043-2399eaa5aaba
---


# CustomXMLNode Object (Office)

Represents an XML node in a tree in a document. The  **CustomXMLNode** object is a member of the **CustomXMLNodes** collection.


## Remarks

The  **CustomXMLNode** object is designed to have functional parity with the **IXMLDOMNode** interface. In addition, it contains an **XPath** property, which is a great improvement over the objects provided by MSXML.


## Example

The following example selects a single node from a  **CustomXMLPart** object by using an XPath expression and assigns it to a **CustomXMLNode** object.


```
Sub CustomXmlNodes()  
    Dim cxp1 As CustomXMLPart 
    Dim cxn As CustomXMLNode 
 
    With ActiveDocument 
 
        ' Returns the first custom xml part with the given root namespace. 
        Set cxp1 = .CustomXMLParts("urn:invoice:namespace")  
         
        ' Get the first node matching the XPath expression.                              
        Set cxn = cxp1.SelectSingleNode("//*[@quantity < 4]") 
                 
    End With 
     
End Sub
```


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AppendChildNode](http://msdn.microsoft.com/library/3fbe1c76-b60a-e365-4988-4a94a52e1fe0%28Office.15%29.aspx)||
|[AppendChildSubtree](http://msdn.microsoft.com/library/67899ba9-7e5a-e40e-2e33-b02ff1fff4b4%28Office.15%29.aspx)||
|[Delete](http://msdn.microsoft.com/library/e240dea8-3045-634d-1ac8-782facf85d4e%28Office.15%29.aspx)||
|[HasChildNodes](http://msdn.microsoft.com/library/9afc3116-372c-7efa-8cdd-04f87d903cc2%28Office.15%29.aspx)||
|[InsertNodeBefore](http://msdn.microsoft.com/library/b2805906-16b7-aebd-ccde-ded736a1b69b%28Office.15%29.aspx)||
|[InsertSubtreeBefore](http://msdn.microsoft.com/library/5d9e9303-e427-a092-3960-eee90a53970d%28Office.15%29.aspx)||
|[RemoveChild](http://msdn.microsoft.com/library/dc6c380a-6cfd-870a-9a31-d92aed1ae3e1%28Office.15%29.aspx)||
|[ReplaceChildNode](http://msdn.microsoft.com/library/72d571f4-8a54-b250-ce5d-22d595ef09f4%28Office.15%29.aspx)||
|[ReplaceChildSubtree](http://msdn.microsoft.com/library/955ec2ab-c6c9-242c-5e05-3ff03b00b120%28Office.15%29.aspx)||
|[SelectNodes](http://msdn.microsoft.com/library/443592af-a684-ee5e-98af-3e157f0f135e%28Office.15%29.aspx)||
|[SelectSingleNode](http://msdn.microsoft.com/library/630751f0-fe41-8f91-32d0-e266b3214cbf%28Office.15%29.aspx)||

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/2cf465cc-fda8-7599-7cd3-f8ff72746fa3%28Office.15%29.aspx)|
|[Attributes](http://msdn.microsoft.com/library/406847e4-25e4-77c6-883c-9cc85f781c73%28Office.15%29.aspx)|
|[BaseName](http://msdn.microsoft.com/library/7b5a6266-4020-6cab-3b4b-b3bbb59a0daa%28Office.15%29.aspx)|
|[ChildNodes](http://msdn.microsoft.com/library/6b0dcfde-8811-ff56-8f56-24db20bc1750%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/e2dc5b81-6bfe-abc9-f5e5-a3de4d1348ff%28Office.15%29.aspx)|
|[FirstChild](http://msdn.microsoft.com/library/8aa38a63-32a3-e798-83de-9797143dd1b9%28Office.15%29.aspx)|
|[LastChild](http://msdn.microsoft.com/library/b9172003-4cad-eee2-8ca6-48e120f7781a%28Office.15%29.aspx)|
|[NamespaceURI](http://msdn.microsoft.com/library/4bb671fd-b2e5-0259-40cf-5499ae0c747e%28Office.15%29.aspx)|
|[NextSibling](http://msdn.microsoft.com/library/75dff508-f657-f94e-fbff-8bab0f4e5192%28Office.15%29.aspx)|
|[NodeType](http://msdn.microsoft.com/library/e656ecb6-091e-bd1a-11ee-6c3860530215%28Office.15%29.aspx)|
|[NodeValue](http://msdn.microsoft.com/library/66be9dfe-0a8f-9522-7974-e00497ac9118%28Office.15%29.aspx)|
|[OwnerDocument](http://msdn.microsoft.com/library/7f604384-76d0-d532-9d32-18c39e1eddab%28Office.15%29.aspx)|
|[OwnerPart](http://msdn.microsoft.com/library/e0db2121-1488-b44f-d68f-7118a844fd5b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/cc87761a-a004-2d58-6f19-631ec2b7b7e1%28Office.15%29.aspx)|
|[ParentNode](http://msdn.microsoft.com/library/f9cfaf3e-1a86-e3ef-e1a1-d52e58d5b1ea%28Office.15%29.aspx)|
|[PreviousSibling](http://msdn.microsoft.com/library/511e6dfd-7027-220a-9d3e-e998a43e7239%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/9d5acd94-2f18-dbff-88f7-cb72b062ddc3%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/28a95285-f751-e0da-f6ce-f16082430176%28Office.15%29.aspx)|
|[XPath](http://msdn.microsoft.com/library/28159c24-79b2-a3ee-589e-de080dd67a82%28Office.15%29.aspx)|

## See also


#### Other resources


[CustomXMLNode Object Members](http://msdn.microsoft.com/library/fbf957c8-40b8-2f75-fcc8-db0ed6e18438%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
