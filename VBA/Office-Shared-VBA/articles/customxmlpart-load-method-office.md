---
title: CustomXMLPart.Load Method (Office)
keywords: vbaof11.chm295010
f1_keywords:
- vbaof11.chm295010
ms.prod: office
api_name:
- Office.CustomXMLPart.Load
ms.assetid: f4d50c05-15bd-ccce-6198-9d6be401b29b
ms.date: 06/08/2017
---


# CustomXMLPart.Load Method (Office)

Allows the template author to populate a  **CustomXMLPart** from an existing file. Returns **True** if the load was successful.


## Syntax

 _expression_. **Load**( **_FilePath_** )

 _expression_ An expression that returns a **CustomXMLPart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FilePath_|Required|**String**|Points to the file on the user's computer or on a network containing the XML to be loaded.|

### Return Value

Boolean


## Example

The following example adds a custom XML part, populates the custom XML part with XML from a file, and then manipulates the part's nodes.


```
Sub ShowCustomXmlParts() 
    On Error GoTo Err 
 
    Dim cxp1 As CustomXMLPart 
 
    With ActiveDocument 
        ' Example written for Word. 
 
        ' Add a custom XML part and then load the XML from a file. 
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

