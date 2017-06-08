---
title: Shapes.Item Property (Visio)
keywords: vis_sdr.chm11313765
f1_keywords:
- vis_sdr.chm11313765
ms.prod: visio
api_name:
- Visio.Shapes.Item
ms.assetid: c7f9eb37-b7dc-a938-bece-d70e215fe48e
ms.date: 06/08/2017
---


# Shapes.Item Property (Visio)

Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_NameUIDOrIndex_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameUIDOrIndex_|Required| **Variant**|Contains the name, unique ID, or index of the object to retrieve.|

### Return Value

Shape


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statements are equivalent to the syntax example given above:


```
objRet = object(index)  
objRet = object(stringExpression)
```

You can retrieve an object in an  **Addons** , **Documents** , **Fonts** , **Hyperlinks** , **Layers** , **Masters** , **MasterShortcuts** , **OLEObjects** , **Pages** , **Shapes** , or **Styles** collection by passing the object's name as a string expression in a **Variant** .

If you retrieve a  **Shape** object by name, the **Item** property searches all shapes in the **Shapes** collection's containing page or containing master, in addition to the collection's containing shape. Therefore, the **Shape** object returned by the **Item** property can be a shape that is not in the **Shapes** collection.

You can also pass the unique ID string of a  **Master** or **Shape** object to the **Item** property. For example:




```
objRet = vsoShapes.Item("{2287DC42-B167-11CE-88E9-0020AFDDD917}")
```

If such a string is passed to the  **Item** property of a **Shapes** collection, all the shapes contained in the collection are searched. Shapes within the group shapes in the containing shape are not searched.

To search all shapes in the collection, plus the shapes inside groups and the containing shape of the collection, prefix the unique ID string with an asterisk (*). For example: 




```
objRet = vsoShapes.Item("*{2287DC42-B167-11CE-88E9-0020AFDDD917}")
```

For more information about passing ID strings to the  **Item** property, see the topic for the **[UniqueID](shape-uniqueid-property-visio.md)** property in this Automation Reference.


 **Note**  


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Item** property to get a **Page** object from the **Pages** collection of the active document, and all the **Shape** objects in the **Shapes** collection of the **Page** object. It prints the names of all shapes on Page1 in the Immediate window.

Before running this macro, make sure that the active document has shapes on Page1.




```vb
Public Sub Item_Example() 
  
    Dim intCounter As Integer 
    Dim intShapeCount As Integer 
    Dim vsoShapes As Visio.Shapes  
 
    Set vsoShapes = ActiveDocument.Pages.Item(1).Shapes  
 
    Debug.Print "Shape Name List For..." 
    Debug.Print "Document: "; ActiveDocument.Name  
    Debug.Print "Page: "; ActiveDocument.Pages.Item(1).Name  
 
    intShapeCount = vsoShapes.Count  
 
    If intShapeCount > 0 Then 
        For intCounter = 1 To intShapeCount  
            Debug.Print " "; vsoShapes.Item(intCounter).Name  
        Next intCounter  
    Else 
        Debug.Print " No Shapes On Page"  
    End If   
 
End Sub
```


