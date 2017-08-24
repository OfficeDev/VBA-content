---
title: Shapes.AddPicture Method (Publisher)
keywords: vbapb10.chm2162710
f1_keywords:
- vbapb10.chm2162710
ms.prod: publisher
api_name:
- Publisher.Shapes.AddPicture
ms.assetid: a5305bd0-295f-46f6-7823-46dab750243b
ms.date: 06/08/2017
---


# Shapes.AddPicture Method (Publisher)

Adds a new  **Shape** object representing a picture to the specified **Shapes** collection.


## Syntax

 _expression_. **AddPicture**( **_Filename_**,  **_LinkToFile_**,  **_SaveWithDocument_**,  **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Filename|Required| **String**|The name of the picture file to insert into the shape. The path can be absolute or relative.|
|LinkToFile|Required| **MsoTriState**|Determines whether the picture is linked to or embedded in the publication.|
|SaveWithDocument|Required| **MsoTriState**|Determines whether the picture is saved as a separate file with the publication.|
|Left|Required| **Variant**|The position of the left edge of the shape representing the picture.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the picture.|
|Width|Optional| **Variant**|The width of the shape representing the picture. Default is -1, meaning that the width of the shape is automatically set based on the object's data.|
|Height|Optional| **Variant**|The height of the shape representing the picture. Default is -1, meaning that the width of the shape is automatically set based on the object's data.|

### Return Value

Shape


## Remarks

If the SaveWithDocument argument is  **msoTrue**, Microsoft Publisher saves a new copy of the picture file specified by the FileName argument in the same directory as the publication.

The LinkToFile and SaveWithDocument arguments cannot have the same value, or else an error occurs. If either argument is  **msoTrue**, the other must be  **msoFalse**.

For the Left, Top, Width, and Height arguments, numeric values are evaluated in points; strings can be in any units supported by Publisher (for example, "2.5 in").

The LinkToFile parameter can be one of the  **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The picture is to be embedded in the publication.|
| **msoTrue**|The picture is to be linked to the publication.|

## Example

The following example adds a picture based on an existing file to the active publication; the picture in the publication is linked to a copy of the original file. (Note that  _PathToFile_ must be replaced with a valid file path for this example to work.)


```vb
Dim shpPicture As Shape 
 
Set shpPicture = ActiveDocument.Pages(1).Shapes.AddPicture _ 
 (FileName:="PathToFile", _ 
 LinkToFile:=msoTrue, _ 
 SaveWithDocument:=msoTrue 
 Left:=72, Top:=72)
```


