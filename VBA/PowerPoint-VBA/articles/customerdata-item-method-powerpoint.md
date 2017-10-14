---
title: CustomerData.Item Method (PowerPoint)
keywords: vbapp10.chm675003
f1_keywords:
- vbapp10.chm675003
ms.prod: powerpoint
api_name:
- PowerPoint.CustomerData.Item
ms.assetid: 4ccbd7b2-3fd5-fc13-42b6-060fc88f1465
ms.date: 06/08/2017
---


# CustomerData.Item Method (PowerPoint)

Returns the specified  **[CustomXMLPart](http://msdn.microsoft.com/library/a4f90bac-01d6-bba4-f64b-a64e2b122cfd%28Office.15%29.aspx)** object from the **[CustomerData](customerdata-object-powerpoint.md)** collection. Read-only.


## Syntax

 _expression_. **Item**( **_Id_** )

 _expression_ An expression that returns a **CustomerData** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Id_|Required|**String**|The ID of the  **CustomXMLPart** object.|

### Return Value

CustomXMLPart


## Remarks

Individual  **CustomXMLPart** objects in the **CustomerData** collection are represented by GUIDs (globally unique identifiers). Pass the GUID that represents the custom XML part that you want to get to the Id parameter of the **Item** method as a **String**. You can get the ID of a particular custom XML part by getting the value of the **Id** property of the **CustomXMLPart** object.


## Example

The following example shows how to use the Item method to get a custom XML part by its ID string.


```vb
Public Sub Item_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
     
    Dim pptShape As Shape 
    Set pptShape = pptSlide.Shapes(1) 
     
    ' Get the CustomerData collection of the shape 
    Dim pptCustomerData As customerData 
    Set pptCustomerData = pptShape.customerData 
     
    ' Add a new CustomXMLPart object to the 
    ' CustomerData collection for this shape 
    Dim pptCustomXMLPart As CustomXMLPart 
    Set pptCustomXMLPart = pptCustomerData.Add 
            
    ' Get the ID of the new part 
    Dim myString As String 
    myString = pptCustomXMLPart.Id 
    Debug.Print myString 
     
    ' Get the new part from the collection by its Id 
    ' and load XML into the part 
    pptCustomerData.Item(myString).LoadXML ("<text>This is XML data.</text>") 
    Debug.Print pptCustomXMLPart.xml 
 
End Sub
```


## See also


#### Concepts


[CustomerData Collection](customerdata-object-powerpoint.md)

