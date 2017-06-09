---
title: CustomXMLNode.Delete Method (Office)
keywords: vbaof11.chm294021
f1_keywords:
- vbaof11.chm294021
ms.prod: office
api_name:
- Office.CustomXMLNode.Delete
ms.assetid: e240dea8-3045-634d-1ac8-782facf85d4e
ms.date: 06/08/2017
---


# CustomXMLNode.Delete Method (Office)

Deletes the current node from the tree (including all of its children, if any exist).


## Syntax

 _expression_. **Delete**

 _expression_ An expression that returns a **CustomXMLNode** object.


## Remarks

If the operation would result in an invalid tree structure, the deletion is not performed and an error message is displayed.


## Example

The following example demonstrates using various methods to add custom XML parts, select parts and nodes with different criteria, append child subtrees, and delete parts and nodes.


```
Sub ShowCustomXmlParts() 
    On Error GoTo Err 
 
    Dim cxps As CustomXMLParts 
    Dim cxp1 As CustomXMLPart 
    Dim cxp2 As CustomXMLPart 
    Dim cxn As CustomXMLNode 
    Dim cxns As CustomXMLNodes 
    Dim strXml As String 
    Dim strUri As String 
 
    With ActiveDocument 
        ' Example written for Word. 
 
        ' Adding a custom XML part. 
        .CustomXMLParts.Add "<custXMLPart />" 
         
        ' Add and then load from a file. 
        Set cxp1 = .CustomXMLParts.Add 
        cxp1.Load "c:\invoice.xml" 
         
        ' Returns the first custom XML part with the given root namespace. 
        Set cxp2 = .CustomXMLParts("urn:invoice:namespace")     '  
         
        ' Access all with a given root namespace; returns the entire collection. 
        Set cxps = .CustomXMLParts.SelectByNamespace("urn:invoice:namespace") 
         
        ' DOM-type operations. 
        ' Get the XML. 
        strXml = cxp2.XML 
        ' Get the root namespace 
        strUri = cxp2.NamespaceURI  
        ' Get nodes using XPath.                              
        Set cxn = cxp2.SelectSingleNode("//*[@quantity < 4]")  
        Set cxns = cxp2.SelectNodes("//*[@unitPrice > 20]") 
        ' Append a child subtree to the single node selected previously. 
        cxn.AppendChildSubtree("<discounts><discount>0.10</discount></discounts>")          
         
        ' Delete custom XML part and node and its children. 
        cxp2.Delete 
        cxn.Delete 
 
                 
    End With 
     
    Exit Sub 
                 
' Exception handling. Show the message and resume. 
Err: 
        MsgBox (Err.Description) 
        Resume Next 
End Sub
```


## See also


#### Concepts


[CustomXMLNode Object](customxmlnode-object-office.md)
#### Other resources


[CustomXMLNode Object Members](customxmlnode-members-office.md)

