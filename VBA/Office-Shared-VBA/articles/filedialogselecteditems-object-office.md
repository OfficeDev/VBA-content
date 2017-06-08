---
title: FileDialogSelectedItems Object (Office)
keywords: vbaof11.chm253000
f1_keywords:
- vbaof11.chm253000
ms.prod: office
api_name:
- Office.FileDialogSelectedItems
ms.assetid: a72b1d99-8881-0a5f-9814-3e1b8360d011
ms.date: 06/08/2017
---


# FileDialogSelectedItems Object (Office)

A collection of  **String** values that correspond to the paths of the files or folders that a user has selected from a file dialog box displayed through the **FileDialog** object.


## Example

Use the  **SelectedItems** property with the **FileDialog** object to return a **FileDialogSelectedItems** collection. The following example displays a **File Picker** dialog box and displays each selected file in a message box.


```
Sub Main() 
 
 'Declare a variable as a FileDialog object. 
 Dim fd As FileDialog 
 
 'Create a FileDialog object as a File Picker dialog box. 
 Set fd = Application.FileDialog(msoFileDialogFilePicker) 
 
 'Declare a variable to contain the path 
 'of each selected item. Even though the path is aString, 
 'the variable must be a Variant because For Each...Next 
 'routines only work with Variants and Objects. 
 Dim vrtSelectedItem As Variant 
 
 'Use a With...End With block to reference the FileDialog object. 
 With fd 
 
 'Allow the selection of multiple file. 
 .AllowMultiSelect = True 
 
 'Use the Show method to display the File Picker dialog box and return the user's action. 
 'The user pressed the button. 
 If .Show = -1 Then 
 
 'Step through each string in the FileDialogSelectedItems collection 
 For Each vrtSelectedItem In .SelectedItems 
 
 'vrtSelectedItem is aString that contains the path of each selected item. 
 'You can use any file I/O functions that you want to work with this path. 
 'This example displays the path in a message box. 
 MsgBox "Selected item's path: " &amp; vrtSelectedItem 
 
 Next vrtSelectedItem 
 'The user pressed Cancel. 
 Else 
 End If 
 End With 
 
 'Set the object variable to Nothing. 
 Set fd = Nothing 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Item](filedialogselecteditems-item-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](filedialogselecteditems-application-property-office.md)|
|[Count](filedialogselecteditems-count-property-office.md)|
|[Creator](filedialogselecteditems-creator-property-office.md)|
|[Parent](filedialogselecteditems-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
