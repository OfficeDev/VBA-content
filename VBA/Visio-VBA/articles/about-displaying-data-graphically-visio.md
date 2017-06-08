---
title: About Displaying Data Graphically (Visio)
ms.prod: visio
ms.assetid: 48acb79c-44b8-c63b-f7fb-409b5aa9b0cd
ms.date: 06/08/2017
---


# About Displaying Data Graphically (Visio)

 **Note**  Data-connectivity features are available only to licensed users of Microsoft Visio Professional 2013.

There are four aspects of data connectivity in Visio:

- Connecting to a data source
    
- Linking shapes to data
    
- Displaying linked data graphically
    
- Refreshing linked data that has changed in the data source, updating linked shapes, and resolving any subsequent conflicts that may arise
    
Typically, you approach these aspects in the order in which they are listed; that is, you first connect your Visio drawing to a data source, then link shapes in your drawing to data in the data source, display the data in linked shapes graphically, and refresh the linked data when necessary. 
Each of these aspects has new objects and members associated with it in the Visio object model. This topic deals with the third and fourth of these aspects, displaying linked data graphically in shapes in your Visio, and refreshing data. For more information about the other aspects of data connectivity, see the following topics: 

-  [About Connecting to Data in Visio](about-connecting-to-data-in-visio.md)
    
-  [About Linking Shapes to Data](about-linking-shapes-to-data.md)
    
To display linked data programmatically, you can use the Visio API for data display, which includes the following objects and their associated members:

-  **[GraphicItems](graphicitems-object-visio.md)** collection
    
-  **[GraphicItem](graphicitem-object-visio.md)** object
    
After you  [link shapes in your Visio drawing to rows in a data recordset](about-linking-shapes-to-data.md), you can graphically display the linked data programmatically. For example, say that your drawing contains several data-linked shapes, each of which represents a project at a particular stage of completion. You may associate a progress bar with a particular item of shape data, such as the percentage of project completion. You could then apply the progress bar to a selection of project shapes and show each project's progress toward completion visually.

## Overview of Data Graphics and Graphic Items

To make it easier to display data graphically, Visio introduces the concept of data graphics and a type of  **[Master](master-object-visio.md)** object called a _data graphic master_, which is represented in the  **[VisMasterTypes](vismastertypes-enumeration-visio.md)** enumeration by the value **visTypeDataGraphic**. To add a  **Master** object of type **visTypeDataGraphic** to the **Masters** collection, you must use the **[Masters.AddEx](masters-addex-method-visio.md)** method.

Visio includes several types of masters, including shape masters. When you create an instance of a shape master, it becomes a shape. Visio also includes fill pattern, line pattern, and line-end masters, for which you cannot create instances. You apply these masters to shapes to impart the master pattern to the shape. Data graphic masters are more like pattern masters, because you do not create instances of them. Instead, you apply them to shapes as you do line pattern and fill pattern masters.

Data graphic masters correspond to the data graphics that appear in the  **Data Graphics** task pane in the Visio UI. A data graphic master consists of one or more _graphic items_. Graphic items are Visio shapes designed to be ready-made visual components that you can associate with shape data to display that data graphically, based on rules that you define, and at a position relative to the shape that you specify.

Visio provides graphic items of the following types: 


-  **Text** Displays data as text in a callout, at a specified position relative to the shape.
    
-  **Color by Value** Changes the color of the shape based on a comparison between shape data and a particular value or range of values.
    
-  **Data Bar** Uses bar charts and graphs to display data, at a specified position relative to the shape.
    
-  **Icon Set** Displays one of a set of icons that represents a data value or condition, at a specified position relative to the shape.
    
Visio provides a variety of standard data graphics that are already populated with graphic items. If you want to apply a data graphic to your shapes that has a different combination of graphic items, you can create a custom data graphic. We recommend that you use the Visio UI to create a data graphic and add graphic items to it.


## To create data graphics in the UI:


1. On the  **Data** tab, click **Data Graphics**.
    
2. Click  **Create New Data Graphic**, and then in the  **New Data Graphic** dialog box, click **New Item**.
    
3. In the dialog box that opens, customize the item, and then use the same method to add custom items.
    
You can also create data graphic masters and populate them with existing graphic items programmatically. You cannot create graphic items programmatically, but you can customize the behavior of existing data graphics. In addition, you can use code to change the behavior and position of graphic items, as well as the rules, called  _expressions_, that define how individual graphic items display data. Expressions can be ShapeSheet formulas or any other legal ShapeSheet expressions, or shape-data (custom property) labels. To set an expression that is a shape-data label, you must enclose the label in curly braces ({}) and then pass it as the second ( _Expression_) parameter of the  **[GraphicItem.SetExpression](graphicitem-setexpression-method-visio.md)** method.

After you create a data graphic that contains a custom combination of graphic items and define the behavior of those graphic items, you can apply the data graphic to data-linked shapes programmatically.


## Data Graphics Objects and Members

Besides the  **Master** objects of type **visTypeDataGraphic** described in the previous section, Visio provides the following objects and their associated members in the data graphics API:


-  **[GraphicItems](graphicitems-object-visio.md)** collection
    
-  **[GraphicItem](graphicitem-object-visio.md)** object
    
In addition to these specifically data-graphic-related objects and their members, several members of other, more conventional Visio objects constitute part of the data graphics API. For example, the  **[Shape.DataGraphic](shape-datagraphic-property-visio.md)** and **[Selection.DataGraphic](selection-datagraphic-property-visio.md)** properties allow you to apply data graphics to shapes and selections respectively. The read-only **[Shape.IsDataGraphicCallout](shape-isdatagraphiccallout-property-visio.md)** property indicates whether a specific shape is functioning as a data graphic item in your drawing.


## Applying Data Graphics to Data-linked Shapes

The following example shows how to use the  **Selection.DataGraphic** property to apply an existing custom data graphic that you create in the UI to a selection of shapes in your drawing. For this code to work, the existing custom data graphic must be named "MyCustomDataGraphic." Alternatively, you can substitute the name of an existing data graphic in your drawing for "MyCustomDataGraphic" in the code.


```vb
Public Sub ApplyDataGraphic() 
    Dim vsoSelection As Visio.Selection 
    ActiveWindow.SelectAll 
    Set vsoSelection = ActiveWindow.Selection 
    Set vsoSelection.DataGraphic = ActiveDocument.Masters("MyCustomDataGraphic") 
End Sub
```


## Customizing the Behavior of Data Graphic Masters

You can use the  **[Master.DataGraphicHidden](master-datagraphichidden-property-visio.md)** and **[Master.DataGraphicHidesText](master-datagraphichidestext-property-visio.md)** properties to customize certain aspects of the behavior of data graphic masters.

The  **DataGraphicHidden** property determines whether a data graphic master appears in the **Data Graphics** gallery in the Visio UI. When you set the value of this property to **True** for a given master, the master does not appear in the list of data graphics in the gallery. The default value of the property is **False**.

The  **DataGraphicsHidesText** property determines whether applying a data graphic master hides the text of the shape to which it is applied (the primary shape in the case of a group shape.) The default value of this property also is **False**.

The  **[GraphicItem.UseDataGraphicPosition](graphicitem-usedatagraphicposition-property-visio.md)** property determines whether to use the current default callout position for graphic items of the data graphic master to whose **GraphicItems** collection a graphic item belongs. The default callout position for graphic items in the **GraphicItems** collection of a **Master** object of type **visTypeDataGraphic** is specified by the settings of the **[Master.DataGraphicVerticalPosition](master-datagraphicverticalposition-property-visio.md)** and **[Master.DataGraphicHorizontalPosition](master-datagraphichorizontalposition-property-visio.md)** properties. If **UseDataGraphicPosition** is **True**, the graphic item is positioned according to the default setting. If  **UseDataGraphicPosition** is **False**, its position is determined by the settings of the  **[Graphic Item.VerticalPosition](graphicitem-verticalposition-property-visio.md)** and **[GraphicItem.HorizontalPosition](graphicitem-horizontalposition-property-visio.md)** properties.

In addition, if the  **HorizontalPosition** and **VerticalPosition** property values of a graphic item are equal to the **DataGraphicHorizontalPosition** and **DataGraphicVerticalPosition** property values, the value of the **UseDataGraphicPosition** property for that graphic item is automatically set to **True**.

Note, however, that you can manually re-position a data graphic that has been applied to a shape by using the control handle of the data graphic. A position set in this manner takes precedence over the position specified by property settings.

The  **[Master.DataGraphicShowBorder](master-datagraphicshowborder-property-visio.md)** property determines whether a border is displayed around graphic items that are in default positions relative to the shape to which a data graphic is applied. By default, the border is hidden.


## Assembling Data Graphics Programmatically

The following example shows how to create a data graphic master, add an existing graphic item to it, and then modify the graphic item. This example uses the  **Masters.AddEx** method to add a new data graphic master to the **Masters** collection of the current document.

Next, it uses the  **Master.Open** method to get a copy of an existing data graphic master to edit. For more information about why it is necessary to edit a copy of a master, instead of the master itself, see **Open** Method. Next, it uses the **GraphicItems.AddCopy** method to add a copy of an existing graphic item to the **GraphicItems** collection of the new master, and the **GraphicItem.SetExpression** method to modify the data field that the graphic item represents. It also sets the **GraphicItem.PositionHorizontal** property to modify the horizontal position of the graphic item relative to the shape to which it is applied.

Finally, it sets the  **Master.DataGraphicHidesText** property to **True** to hide the text of the shape, and closes the copy of the master, which applies the changes to existing shapes to which this data graphic master is applied. You can then apply the new data graphic master to additional shapes.




```vb
Public Sub AddNewDataGraphicMaster() 
 
    Dim vsoMaster As Visio.Master 
    Dim vsoMasterCopy As Visio.Master 
    Dim vsoMaster_Old As Visio.Master 
    Dim vsoGraphicItem As GraphicItem 
    Dim vsoGraphicItem_Old As Visio.GraphicItem 
 
    Set vsoMaster = ActiveDocument.Masters.AddEx(visTypeDataGraphic) 
    Set vsoMasterCopy = vsoMaster.Open 
    Set vsoMaster_Old = ActiveDocument.Masters("old_master_name") 
    Set vsoGraphicItem_Old = vsoMaster_Old.GraphicItems(1) 
    Set vsoGraphicItem = vsoMasterCopy.GraphicItems.AddCopy(vsoGraphicItem_Old) 
 
    vsoGraphicItem.SetExpression visGraphicExpression, "new_data_field_name" 
    vsoGraphicItem.PositionHorizontal = visGraphicLeft 
    vsoMasterCopy.DataGraphicHidesText = True; 
    vsoMasterCopy.Close 
 
End Sub
```

The preceding code sample assumes that you know the name of the existing data graphic master that contains one or more graphic items you want to add to the new master, as well as the IDs of one or more graphic items you want to add to the master. You can determine the name of an existing data graphic master by moving your mouse over the master in the  **Data Graphics** task pane. You can also determine master names and IDs by iterating through the **Masters** collection in the current document, as shown in the following code.




```vb
For intCounter = 1 To ActiveDocument.Masters.Count 
        If ActiveDocument.Masters(intCounter).Type = visTypeDataGraphic Then 
            Debug.Print ActiveDocument.Masters(intCounter).Name, ActiveDocument.Masters(intCounter).ID 
        End If 
    Next
```

Similarly, you can iterate through the  **GraphicItems** collection of a master to determine the values of the **[ID](graphicitem-id-property-visio.md)** and **[Tag](graphicitem-tag-property-visio.md)** properties of an existing graphic item, as shown in the following example. The **Tag** property is a string that Visio does not use. It is empty by default. However, you can set its value to make it easier to identify individual graphic items programmatically.




```vb
For intCounter = 1 To (vsoMaster_Old.GraphicItems.Count) 
        Debug.Print vsoMaster_Old.GraphicItems(intCounter).ID, oldMaster.GraphicItems(intCounter).Tag 
    Next
```

To see a code sample that shows how to customize data graphics programmatically, download the Visio SDK and refer to the Code Samples Library.


