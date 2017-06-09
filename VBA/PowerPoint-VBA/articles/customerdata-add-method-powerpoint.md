---
title: CustomerData.Add Method (PowerPoint)
keywords: vbapp10.chm675004
f1_keywords:
- vbapp10.chm675004
ms.prod: powerpoint
api_name:
- PowerPoint.CustomerData.Add
ms.assetid: f39bc83a-4c3b-6803-12d1-9ae72e601b49
ms.date: 06/08/2017
---


# CustomerData.Add Method (PowerPoint)

 Adds a **[CustomXMLPart](http://msdn.microsoft.com/library/a4f90bac-01d6-bba4-f64b-a64e2b122cfd%28Office.15%29.aspx)** to the **[CustomerData](customerdata-object-powerpoint.md)** collection of a **[CustomLayout](customlayout-object-powerpoint.md)**, **[Master](master-object-powerpoint.md)**, **[Presentation](presentation-object-powerpoint.md)**, **[Shape](shape-object-powerpoint.md)**, or **[Slide](slide-object-powerpoint.md)** object and returns the **CustomXMLPart** object created.


## Syntax

 _expression_. **Add**

 _expression_ An expression that returns a **CustomerData** object.


### Return Value

CustomXMLPart


## Remarks

You can add one or more items of customer data (custom XML parts) to any of the objects listed above that can contain customer data.


## Example




```vb
Public Sub Add_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
     
    Dim pptShape As Shape 
    For Each pptShape In pptSlide.Shapes 
         
        ' Get the CustomerData collection of the shape 
        Dim pptCustomerData As customerData 
        Set pptCustomerData = pptShape.customerData 
         
        ' Add a new CustomXMLPart object to the CustomerData collection for this shape 
        Dim pptCustomXMLPart As CustomXMLPart 
        Set pptCustomXMLPart = pptCustomerData.Add 
         
        ' Add data to the CustomXMLPart 
        pptCustomXMLPart.LoadXML ("<ShapeData><DataItem>This has to be valid XML</DataItem></ShapeData>") 
         
        ' Print the ID (a GUID) of the CustomXMLPart 
        Debug.Print (pptCustomXMLPart.Id) 
         
    Next 
 
End Sub
```


## See also


#### Concepts


[CustomerData Collection](customerdata-object-powerpoint.md)

