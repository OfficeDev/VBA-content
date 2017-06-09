---
title: Shapes.AddOLEObject Method (Publisher)
keywords: vbapb10.chm2162709
f1_keywords:
- vbapb10.chm2162709
ms.prod: publisher
api_name:
- Publisher.Shapes.AddOLEObject
ms.assetid: c454f9cb-2005-5e55-80a7-6dfbe9c109e5
ms.date: 06/08/2017
---


# Shapes.AddOLEObject Method (Publisher)

Adds a new  **[Shape](shape-object-publisher.md)** object representing an OLE object to the specified **[Shapes](shapes-object-publisher.md)** collection.


## Syntax

 _expression_. **AddOLEObject**( **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**,  **_ClassName_**,  **_Filename_**,  **_Link_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Left|Required| **Variant**|The position of the left edge of the shape representing the OLE object.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the OLE object.|
|Width|Optional| **Variant**|The width of the shape representing the OLE object. Default is -1, meaning that the width of the shape is automatically set based on the object's data.|
|Height|Optional| **Variant**|The height of the shape representing the OLE object. Default is -1, meaning that the width of the shape is automatically set based on the object's data.|
|ClassName|Optional| **String**|The class name of the OLE object to be added.|
|Filename|Optional| **String**|The file name of the OLE object to be added. If the path is not specified, the current working folder is used.|
|Link|Optional| **MsoTriState**|Determines whether the OLE object is linked to or embedded in the publication.|

### Return Value

Shape


## Remarks

For the Left, Top, Width, and Height arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

You must specify either a ClassName or FileName. If neither argument is specified, or if both are specified, an error occurs.

The Link parameter can be one of the  **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The OLE object is embedded.|
| **msoTrue**|The OLE object is linked. The default.|

## Example

The following example adds a Microsoft Office Excel worksheet to the first page of the active publication and activates the worksheet for editing.


```vb
Dim shpSheet As Shape 
 
Set shpSheet = ActiveDocument.Pages(1).Shapes.AddOLEObject _ 
 (Left:=72, Top:=72, ClassName:="Excel.Sheet") 
 
shpSheet.OLEFormat.Activate
```


