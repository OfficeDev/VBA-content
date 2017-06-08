---
title: Application.IsValidObject Method (Publisher)
keywords: vbapb10.chm131126
f1_keywords:
- vbapb10.chm131126
ms.prod: publisher
api_name:
- Publisher.Application.IsValidObject
ms.assetid: 56b2bc3a-3e8e-058c-046a-146f0fbb294a
ms.date: 06/08/2017
---


# Application.IsValidObject Method (Publisher)

Determines whether the specified object variable references a valid object and returns a  **Boolean** value: **True** if the specified variable that references an object is valid, **False** if the object referenced by the variable has been deleted.


## Syntax

 _expression_. **IsValidObject**( **_Object_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Object|Required| **Object**|A variable that references an object.|

### Return Value

Boolean


## Example

This example formats the line of a valid object.


```vb
Sub ValidShape(shpObject As Shape) 
 
 If Application.IsValidObject object:=shpObject) = True Then 
 With shpObject.Line 
 .DashStyle = msoLineRoundDot 
 .ForeColor.RGB = RGB(Red:=158, Green:=50, Blue:=208) 
 .Weight = 5 
 End With 
 End If 
 
End Sub
```

Use the following subroutine to call the above subroutine.




```vb
Sub CallValidShape() 
 Call ValidShape(shpObject:=ActiveDocument.Pages(1).Shapes(2)) 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

