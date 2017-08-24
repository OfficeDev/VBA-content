---
title: Shape.WebNavigationBarSetName Property (Publisher)
keywords: vbapb10.chm5308677
f1_keywords:
- vbapb10.chm5308677
ms.prod: publisher
api_name:
- Publisher.Shape.WebNavigationBarSetName
ms.assetid: 0d9abe17-6936-562b-9210-5f092d13f215
ms.date: 06/08/2017
---


# Shape.WebNavigationBarSetName Property (Publisher)

Returns a  **String** that represents the name of the Web navigation bar set the specified shape is an instance of. Read-only.


## Syntax

 _expression_. **WebNavigationBarSetName**

 _expression_A variable that represents a  **Shape** object.


### Return Value

String


## Remarks

This property is only accessible for shapes that represent an instance of a Web navigation bar set. Use the  **[Type](shape-type-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object to determine if a shape represents an instance of a Web navigation bar set.

Use the  **WebNavigationBarSetName** property to return the name of a **[WebNavigationBarSet](webnavigationbarset-object-publisher.md)** object. Multiple pages in a Web publication can each have a shape representing an instance of the same Web navigation bar set. Changes made to a **WebNavigationBarSet** object are reflected in all the shapes representing instances of that Web navigation bar set.


## Example

The following example tests to determine which shapes on the first page of the active document represent instances of Web navigation bars. For each such shape found, the Web navigation bar it represents an instance of is set to auto update.


```vb
Sub SetWebBarsToAutoUpdate() 
 
Dim shpLoop As Shape 
Dim strWebNavBarName As String 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = pbWebNavigationBar Then 
 
 strWebNavBarName = shpLoop.WebNavigationBarSetName 
 With ActiveDocument.WebNavigationBarSets(strWebNavBarName) 
 .AutoUpdate = True 
 End With 
 
 End If 
Next 
 
End Sub
```


