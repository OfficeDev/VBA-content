---
title: Shape.Data3 Property (Visio)
keywords: vis_sdr.chm11213375
f1_keywords:
- vis_sdr.chm11213375
ms.prod: visio
api_name:
- Visio.Shape.Data3
ms.assetid: 0d02964d-0296-5142-e7c3-e319ea80c224
ms.date: 06/08/2017
---


# Shape.Data3 Property (Visio)

Gets or sets the value of the  **Data3** field for a **Shape** object. Read/write.


## Syntax

 _expression_ . **Data3**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

Use the  **Data3** property to supply additional information about a shape. The property can contain up to 64 KB of characters. Text controls should be used with care with a string that is greater than 3,000 characters. Setting the **Data3** property is equivalent to entering information in the **Data 3** box in the **Special** dialog box (click **Shape Name** in the **Shape Design** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx)tab).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to set a shape's  **Data1** , **Data2** , and **Data3** properties. It prints the values of these properties in the **Immediate** window. You can also verify that these values have been set by opening the **Special** dialog box.


```vb
Public Sub Data123_Example() 
 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 
 Set vsoPage = Documents.Add("").Pages(1) 
 Set vsoShape = vsoPage.DrawRectangle(3, 3, 5, 5) 
 
 'Use the Data1, Data2, and Data3 properties to set 
 'the shape's Data fields. 
 vsoShape.Data1 = "Data1_String" 
 vsoShape.Data2 = "Data2_String" 
 vsoShape.Data3 = "Data3_String" 
 
 'Use the Data1, Data2, and Data3 properties to verify 
 'the shape's Data field values. 
 Debug.Print vsoShape.Data1 
 Debug.Print vsoShape.Data2 
 Debug.Print vsoShape.Data3 
 
End Sub
```


