---
title: CustomerData Object (PowerPoint)
keywords: vbapp10.chm675000
f1_keywords:
- vbapp10.chm675000
ms.prod: powerpoint
api_name:
- PowerPoint.CustomerData
ms.assetid: 1d658369-ea6c-6959-cd00-230dc111f765
ms.date: 06/08/2017
---


# CustomerData Object (PowerPoint)

Stores information about a customer (such as name, address, telephone number, and so on) or other information in XML form, as a collection of  **[CustomXMLPart](http://msdn.microsoft.com/library/a4f90bac-01d6-bba4-f64b-a64e2b122cfd%28Office.15%29.aspx)** objects associated with a Microsoft PowerPoint object.


## Remarks

You can store customer data in  **[CustomLayout](customlayout-object-powerpoint.md)**, **[Master](master-object-powerpoint.md)**, **[Presentation](presentation-object-powerpoint.md)**, **[Shape](shape-object-powerpoint.md)**, and **[Slide](slide-object-powerpoint.md)** objects. You can associate one or more **CustomXMLPart** objects with the same object.




- Customer data persists from one instance to the next in a PowerPoint document only when you save the document in XML file format, as a PowerPoint XML presentation. Customer data does not persist in documents saved in .ppt, .htm, or .mht formats.
    
- There is no user interface associated with customer data in PowerPoint. The only way that you can assign and manipulate customer data is programmatically.
    


Use the  **[Add](customerdata-add-method-powerpoint.md)** method to add a new **CustomXMLPart** object to the **CustomerData** collection.

Use the  **[Delete](customerdata-delete-method-powerpoint.md)** method to delete a **CustomXMLPart** object from the **CustomerData** collection.

Use the  **[Item](customerdata-item-method-powerpoint.md)** method to get a specific **CustomXMLPart** object from the collection. Individual items in the collection are represented by GUIDs (globally unique identifiers).

You can use customer data in the same way that you used  **[Tags](tags-object-powerpoint.md)** objects in versions of PowerPoint previous to Microsoft Office PowerPoint 2007--that is, to associate data with objects. Customer data is more powerful than tags, however, because you can store the data as XML instead of as a simple string.

You can associate customer data in PowerPoint with external data by storing the IDs of custom XML parts in a spreadsheet or database along with the external data.

When you copy an object that contains customer data, the customer data is copied to the new object. PowerPoint creates a new  **CustomXMLPart** object to hold the copied data, because two **CustomLayout**, **Master**, **Presentation**, **Shape**, or **Slide** objects can never be associated with the same **CustomXMLPart** object.


## Example

The following example shows how to add a  **CustomXMLPart** object to the **CustomerData** collection of the first shape on the first slide of the active presentation, and how to load an XML string into the custom XML part. It prints the ID of the custom XML part and the XML string in the Immediate window.


```vb
Public Sub CustomerData_Example() 
 
    Dim pptCustomXMLPart As CustomXMLPart 
     
    Set pptCustomXMLPart = ActivePresentation.Slides(1).Shapes(1).customerData.Add 
     
    Debug.Print pptCustomXMLPart.Id 
     
    pptCustomXMLPart.LoadXML ("<Customer><CustomerID>Customer #1</CustomerID></Customer>") 
     
    Debug.Print pptCustomXMLPart.xml 
 
End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

