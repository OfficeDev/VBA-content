---
title: Selection.AutomaticLink Method (Visio)
keywords: vis_sdr.chm11160210
f1_keywords:
- vis_sdr.chm11160210
ms.prod: visio
api_name:
- Visio.Selection.AutomaticLink
ms.assetid: 6943b2b1-269a-7759-d981-a3749cfbeaee
ms.date: 06/08/2017
---


# Selection.AutomaticLink Method (Visio)

Links selected shapes to data rows in the specified data recordset automatically, without requiring you to specify the exact correspondence of all shapes and data rows. Returns the number of shapes linked.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **AutomaticLink**( **_DataRecordsetID_** , **_ColumnNames()_** , **_AutoLinkFieldTypes()_** , **_FieldNames()_** , **_AutoLinkBehavior_** , **_ShapeIDs()_** )

 _expression_ An expression that returns a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordsetID_|Required| **Long**|The value of the  **ID** property of the **DataRecordset** object that contains the data rows to link to.|
| _ColumnNames()_|Required| **String**|An array of strings that correspond to names of columns in the data recordset. |
| _AutoLinkFieldTypes()_|Required| **Long**|An array of  **Long** values from the **[VisAutoLinkFieldTypes](visautolinkfieldtypes-enumeration-visio.md)** enumeration, consisting of shape attribute types. At least one position in the array must have a value that corresponds to the values in the same position in the ColumnNames and FieldNames arrays.|
| _FieldNames()_|Required| **String**|An array of strings that represent shape values.|
| _AutoLinkBehavior_|Required| **Long**|A combination of one or more constants from the  **VisAutoLinkBehavior** enumeration that specify how the linking will occur. See Remarks for possible values.|
| _ShapeIDs()_|Required| **Long**|Out parameter. An array of IDs of shapes (of type  **Long** ) that were linked by the method.|

### Return Value

Long


## Remarks

For the ColumnNames() parameter, pass a string array consisting of names of columns in the database. At least one position in the array must have a value that corresponds to the values in the same position in the AutoLinkFieldTypes() and FieldNames() arrays.

For the AutoLinkFieldTypes() parameter, pass an array of  **Long** values from the **[VisAutoLinkFieldTypes](visautolinkfieldtypes-enumeration-visio.md)** enumeration, consisting of shape attribute types. Among the shape attributes enumerated are height, width, text, and the name of the master the shape was derived from. At least one position in the array must have a value that corresponds to the values in the same position in the ColumnNames() and FieldNames() arrays.

For the FieldNames() parameter, pass an array of strings that represent shape values. At least one position in the FieldNames() array must have a value that corresponds to the values in the same position in the ColumnNames() and AutoLinkFieldTypes() arrays.

For most values of AutoLinkFieldTypes(), for example, for  **visAutoLinkShapeText** , it is not necessary to specify the FieldNames() value; you can pass an empty string instead. However, when you pass the **visAutoLinkCustPropsLabel** , **visAutoLinkUserRowName** , **visAutoLinkPropRowNameU** , or **visAutoLinkUserRowNameU** values of AutoLinkFieldTypes, you must pass a value for FieldNames() that fully specifies the shape data item (called custom property value in some previous versions of Visio) to compare to the data column name.

For the optional AutoLinkBehavior parameter, you can pass a combination of one or more values from the  **VisAutoLinkBehaviors** enumeration that specify how the linking will occur. The following table shows possible values.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visAutoLinkDontReplaceExistingLinks**|16|Do not replace existing links.|
| **visAutoLinkGenericProgressBar**|2|Show generic progress bar instead of more detailed one.|
| **visAutoLinkIncludeHiddenProps**|64|Include hidden properties.|
| **visAutoLinkNoApplyDataGraphic**|4|Do not apply the default data graphic to linked shapes.|
| **visAutoLinkNullMatchesNoFormula**|32|Allow null database values to map to "No Formula" in the Visio ShapeSheet spreadsheet.|
| **visAutoLinkReplaceExistingLinks**|8|Replace existing links.|
| **visAutoLinkSelectedShapesOnly**|1|Link selected shapes only, not sub-shapes of selected shapes.|
You cannot pass a value that includes both  **visAutoLinkDontReplaceExistingLinks** and **visAutoLinkReplaceExistingLinks** . The method returns an error if you attempt to do so.

If you pass a value for AutoLinkBehavior, it modifies the default behavior, which is as follows:


- Use the data recordset's  **LinkReplaceBehavior** setting to determine whether to break existing links. If the setting is **visLinkReplacePrompt** , it is treated as if it were **visLinkReplaceAlways** .
    
- Link selected shapes and their subshapes.
    
- Do not replace the detailed progress bar with the generic progress bar.
    
- Apply data graphics.
    
For the ShapeIDs() parameter, pass an empty, dimensionless array of type  **Long** . Visio will return the array filled with the IDs of the shapes that were linked to data by the method.

To provide Visio with enough information to create the links, you must supply at least one set of matching data: the name of a column in the data recordset, a shape attribute type, and, if necessary, a shape value, all at the same index position of the corresponding arrays you pass to the method. The shape attribute type indicates the attribute of the shape to base the matching upon. The attribute can be the value of a shape data item, shape text, or another of the values specified in the  **VisAutoLinkFieldTypes** enumeration.

For example, say that your drawing contains a selection of shapes representing different employees and that the shape text, which in this case displays the respective employee's names, identifies the shapes. As shown in the example in this topic, you would pass the method the following parameters:


- For the ColumnNames() parameter, an array that contains the "EmployeeName" column name at array position 0.
    
- For the AutoLinkFieldTypes() parameter, enumeration value  **visAutoLinkShapeText** at array position 0.
    
- For the FieldNames() parameter, an empty string (''") at array position 0, because, when AutoLinkFieldTypes() is  **visAutoLinkShapeText** , it is not necessary to specify the FieldNames() value.
    

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AutomaticLink** method to automatically link shapes in a drawing to data in a data recordset. It links data for employees contained in a data recordset to shapes in the drawing whose shape text corresponds to the employee names.

Before running this macro, create a data recordset that contains a column named "EmployeeName" that lists employee names, as well as any other columns you want to include, and then assign those employee names as shape text to corresponding shapes in your Visio drawing. Use the  **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method to add the data recordset to the **DataRecordsets** collection of the active document. Make sure that the data recordset is the one that you added most recently to the collection.




```vb
Public Sub AutomaticLink_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim vsoSelection As Visio.Selection 
    Dim astrColumnNames(1) As String 
    Dim alngFieldTypes(1) As Long 
    Dim astrFieldNames(1) As String 
    Dim alngShapesLinked() As Long 
    Dim intCount As Integer 
     
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount ) 
 
    astrColumnNames(0) = "EmployeeName" 
    alngFieldTypes(0) = Visio.VisAutoLinkFieldTypes.visAutoLinkShapeText 
    astrFieldNames(0) = "" 
 
    ActiveWindow.DeselectAll 
    ActiveWindow.SelectAll 
 
    Set vsoSelection = ActiveWindow.Selection 
    vsoSelection.AutomaticLink vsoDataRecordset.ID, _ 
                    astrColumnNames, _ 
                    alngFieldTypes, _ 
                    astrFieldNames, 0, alngShapesLinked 
 
End Sub
```


