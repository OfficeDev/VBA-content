---
title: CustomXMLParts.PartAfterAdd Event (Office)
keywords: vbaof11.chm299001
f1_keywords:
- vbaof11.chm299001
ms.prod: office
api_name:
- Office.CustomXMLParts.PartAfterAdd
ms.assetid: c1a263a5-94cb-f563-145b-151a52a31d52
ms.date: 06/08/2017
---


# CustomXMLParts.PartAfterAdd Event (Office)

Occurs just after a  **CustomXMLPart** object is added to the **CustomXMLParts** collection.


## Syntax

 _expression_. **PartAfterAdd**( **_NewPart_**, )

 An expression that returns a **CustomXMLParts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewPart_|Required|**CustomXMLPart**|The part that was added.|

## Example

The following example displays the XML contents of a part after it has been added to a  **CustomXMLParts** collection.


```
Sub CustomXMLParts_PartAfterAdd(ByVal objPart As CustomXMLPart) 
Dim strPartXML As String 
strPartXML = objPart.XML 
   MsgBox ("The part's contents are: " &amp; vbCrLf &amp; strPartXML) 
End Sub
```


## See also


#### Concepts


[CustomXMLParts Object](customxmlparts-object-office.md)
#### Other resources


[CustomXMLParts Object Members](customxmlparts-members-office.md)

