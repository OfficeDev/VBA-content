---
title: CustomXMLNode.HasChildNodes Method (Office)
keywords: vbaof11.chm294022
f1_keywords:
- vbaof11.chm294022
ms.prod: office
api_name:
- Office.CustomXMLNode.HasChildNodes
ms.assetid: 9afc3116-372c-7efa-8cdd-04f87d903cc2
ms.date: 06/08/2017
---


# CustomXMLNode.HasChildNodes Method (Office)

Gets  **True** if the current element node has child element nodes.


## Syntax

 _expression_. **HasChildNodes**

 _expression_ An expression that returns a **CustomXMLNode** object.


### Return Value

Boolean


## Remarks

This method will always return  **False** when **CustomXMLNode** isn't of node type **msoCustomXMLNodeElement**.


## Example

The following example demonstrates using various methods to add custom XML parts, select parts and nodes with different criteria, append child subtrees, tests whether the subtree was successfully added, and delete parts and nodes.


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
        ' Get the root namespace. 
        strUri = cxp2.NamespaceURI  
        ' Get nodes using XPath.                              
        Set cxn = cxp2.SelectSingleNode("//*[@quantity < 4]")  
        Set cxns = cxp2.SelectNodes("//*[@unitPrice > 20]") 
        ' Append a child subtree to the single node selected previously. 
        cxn.AppendChildSubtree("<discounts><discount>0.10</discount></discounts>")   
 
         ' Tell user that the child subtree was added. 
         If  cxn.HasChildNodes then 
            MsgBox("The 'Discounts' nodes were added.")  
         Else 
            MsgBox("The 'Discounts' nodes were not added.")  
         End If          
         
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

