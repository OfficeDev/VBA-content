---
title: Layers.Item Property (Visio)
keywords: vis_sdr.chm11913765
f1_keywords:
- vis_sdr.chm11913765
ms.prod: visio
api_name:
- Visio.Layers.Item
ms.assetid: 4f52c429-447c-e240-f6e6-8cad4dcd3322
ms.date: 06/08/2017
---


# Layers.Item Property (Visio)

Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_NameOrIndex_** )

 _expression_ A variable that represents a **Layers** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameOrIndex_|Required| **Variant**|Contains the name, unique ID, or index of the object to retrieve.|

### Return Value

Layer


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statements are equivalent to the syntax example given above:


```
objRet = object(index) 
objRet = object(stringExpression) 

```

You can retrieve an object in an  **Addons** , **Documents** , **Fonts** , **Hyperlinks** , **Layers** , **Masters** , **MasterShortcuts** , **OLEObjects** , **Pages** , **Shapes** , or **Styles** collection by passing the object's name as a string expression in a **Variant** .

For more information about passing ID strings to the  **Item** property, see the topic for the **[UniqueID](shape-uniqueid-property-visio.md)** property in this Automation Reference.


 **Note**  

Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Item** property to access an object in the **Masters** , **Pages** , **Shapes** , **Styles** , **Layers** , or **MasterShortcuts** collection by using its local name. Use the **ItemU** property to access an object from one of these collections by using the object's universal name.


