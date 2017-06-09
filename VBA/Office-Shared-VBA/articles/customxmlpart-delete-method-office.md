---
title: CustomXMLPart.Delete Method (Office)
keywords: vbaof11.chm295009
f1_keywords:
- vbaof11.chm295009
ms.prod: office
api_name:
- Office.CustomXMLPart.Delete
ms.assetid: 2f5b0556-9807-8224-8b3a-e202163fc3e5
ms.date: 06/08/2017
---


# CustomXMLPart.Delete Method (Office)

Deletes the current  **CustomXMLPart** from the data store ( **IXMLDataStore** interface).


## Syntax

 _expression_. **Delete**

 _expression_ An expression that returns a **CustomXMLPart** object.


## Remarks

If you attempt to delete the part containing the core properties, the operation is not performed and an error message is displayed. 


## Example

The following example adds a custom XML part, select a node with a criteria, and delete the part and node.


```
Sub ShowCustomXmlParts() 
    On Error GoTo Err 
 
    Dim cxp1 As CustomXMLPart 
    Dim cxn As CustomXMLNode 
 
    With ActiveDocument 
        ' Example written for Word. 
 
        ' Add and then load from a file. 
        Set cxp1 = .CustomXMLParts.Add 
        cxp1.Load "c:\invoice.xml" 
 
        Set cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")  
        ' Insert a subtree before the single node selected previously. 
        cxn.InsertSubTreeBefore("<discounts><discount>0.10</discount></discounts>")   
               
        ' Delete custom XML part. 
        cxp1.Delete 
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


[CustomXMLPart Object](customxmlpart-object-office.md)
#### Other resources


[CustomXMLPart Object Members](customxmlpart-members-office.md)

