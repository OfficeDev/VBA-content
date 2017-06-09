---
title: View.PasteSpecial Method (PowerPoint)
keywords: vbapp10.chm512010
f1_keywords:
- vbapp10.chm512010
ms.prod: powerpoint
api_name:
- PowerPoint.View.PasteSpecial
ms.assetid: 074fb28f-19c6-3c0f-21ae-75012614485e
ms.date: 06/08/2017
---


# View.PasteSpecial Method (PowerPoint)

Pastes the current contents of the Clipboard into the view represented by the  **View** object.


## Syntax

 _expression_. **PasteSpecial**( **_DataType_**, **_DisplayAsIcon_**, **_IconFileName_**, **_IconIndex_**, **_IconLabel_**, **_Link_** )

 _expression_ A variable that represents a **View** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataType_|Optional|**PpPasteDataType**|A format for the Clipboard contents when they're inserted into the document. The default value varies, depending on the contents in the Clipboard. An error occurs if the specified data type in the DataType argument is not supported by the clipboard contents.|
| _DisplayAsIcon_|Optional|**MsoTriState**|**msoTrue** to display the embedded object (or link) as an icon.|
| _IconFileName_|Optional|**String**|If DisplayAsIcon is set to  **msoTrue**, this argument is the path and file name for the file in which the icon to be displayed is stored. If DisplayAsIcon is set to **msoFalse**, this argument is ignored.|
| _IconIndex_|Optional|**Long**|If DisplayAsIcon is set to  **msoTrue**, this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. Icons appear in the **Change Icon** dialog box, accessed from the **Insert** tab (click **Object**, select  **Display as icon**, click  **Change Icon**): 0 (zero) corresponds to the first icon, 1 corresponds to the second icon. If this argument is omitted, the first (default) icon is used. If DisplayAsIcon is set to  **msoFalse**, this argument is ignored. If IconIndex is outside the valid range, the default icon (index 0) is used.|
| _IconLabel_|Optional|**String**|If DisplayAsIcon is set to  **msoTrue**, this argument is the text that appears below the icon. If this label is missing, Microsoft PowerPoint generates an icon label based on the Clipboard contents. If DisplayAsIcon is set to **msoFalse**, this argument is ignored.|
| _Link_|Optional|**MsoTriState**|Determines whether to create a link to the source file of the Clipboard contents. An error occurs if the Clipboard contents do not support a link.|

## Remarks

An error occurs if there is no data on the Clipboard when the  **PasteSpecial** method is called.

 Valid views for the **PasteSpecial** method are the same as those for the **Paste** method. If the data type can?t be pasted into the view (for example, if you try to paste a picture into **Slide Sorter View**), an error occurs. 

The DataType parameter can be one of these  **PpPasteDataType** constants


||
|:-----|
|**ppPasteBitmap**|
|**ppPasteDefault** default|
|**ppPasteEnhancedMetafile**|
|**ppPasteGIF**|
|**ppPasteHTML**|
|**ppPasteJPG**|
|**ppPasteMetafilePicture**|
|**ppPasteOLEObject**|
|**ppPastePNG**|
|**ppPasteRTF**|
|**ppPasteShape**|
|**ppPasteText**|
The DisplayAsIcon parameter can be one of these  **MsoTriState** constants.


||
|:-----|
|**msoFalse** The default. Does not display the embedded object (or link) as an icon.|
|**msoTrue** Displays the embedded object (or link) as an icon.|
The Link paramter can be one of these  **MsoTriState** constants.


||
|:-----|
|**msoFalse** The default. Does not create a link to the source file of the Clipboard contents.|
|**msoTrue** Creates a link to the source file of the Clipboard contents.|

## Example

The following example pastes a bitmap image as an icon into another window. This example assumes that there are two open windows and that a bitmap image in the first window is currently selected.


```vb
Sub PasteOLEObject()

    Windows(1).Selection.Copy
    Windows(2).View.PasteSpecial DataType:=ppPasteOLEObject, _
        DisplayAsIcon:=msoTrue, IconLabel:="New Bitmap Image"

End Sub
```


## See also


#### Concepts


[View Object](view-object-powerpoint.md)

