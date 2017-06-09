---
title: FileDialogSelectedItems.Count Property (Office)
keywords: vbaof11.chm253002
f1_keywords:
- vbaof11.chm253002
ms.prod: office
api_name:
- Office.FileDialogSelectedItems.Count
ms.assetid: c571c03e-02de-f0a3-0e3f-1fdf9f0d221c
ms.date: 06/08/2017
---


# FileDialogSelectedItems.Count Property (Office)

Gets a  **Long** indicating the number of items in the **FileDialogSelectedItem** s collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **FileDialogSelectedItems** object.


### Return Value

Long


## Example

The following example displays a  **File Picker** dialog box and displays each selected file in a message box.


```
Sub Main() 
 
 'Declare a variable as a FileDialog object. 
 Dim fd As FileDialog 
 Dim cnt As Integer 
 
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
 For cnt = 0 To .SelectedItems.Count - 1 
 
 'vrtSelectedItem is aString that contains the path of each selected item. 
 'You can use any file I/O functions that you want to work with this path. 
 'This example displays the path in a message box. 
 MsgBox "Selected item's path: " &amp; vrtSelectedItem(cnt) 
 
 Next 
 'The user pressed Cancel. 
 Else 
 End If 
 End With 
 
 'Set the object variable to Nothing. 
 Set fd = Nothing 
 
End Sub 

```


## See also


#### Concepts


[FileDialogSelectedItems Object](filedialogselecteditems-object-office.md)
#### Other resources


[FileDialogSelectedItems Object Members](filedialogselecteditems-members-office.md)

