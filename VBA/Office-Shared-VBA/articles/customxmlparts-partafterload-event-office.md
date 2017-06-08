---
title: CustomXMLParts.PartAfterLoad Event (Office)
keywords: vbaof11.chm299003
f1_keywords:
- vbaof11.chm299003
ms.prod: office
api_name:
- Office.CustomXMLParts.PartAfterLoad
ms.assetid: d59fe837-27b5-300f-133f-ffb01f5f95b9
ms.date: 06/08/2017
---


# CustomXMLParts.PartAfterLoad Event (Office)

Occurs just after a  **CustomXMLPart** object is loaded.


## Syntax

 _expression_. **PartAfterLoad**( **_Part_**, )

 _expression_ An expression that returns a **CustomXMLParts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Part_|Required|**CustomXMLPart**|The part that was loaded.|

## Example

The following example adds XML to a part after it is loaded.


```
Sub CustomXMLParts_PartAfterLoad(ByVal objPart As CustomXMLPart) 
   objPart.XML ("<root xmlns='http://www.w3c.org/XMLSchema'>text</root>") 
End Sub
```


## See also


#### Concepts


[CustomXMLParts Object](customxmlparts-object-office.md)
#### Other resources


[CustomXMLParts Object Members](customxmlparts-members-office.md)

