---
title: FileDialog.AllowMultiSelect Property (Office)
keywords: vbaof11.chm256006
f1_keywords:
- vbaof11.chm256006
ms.prod: office
api_name:
- Office.FileDialog.AllowMultiSelect
ms.assetid: b109b0b5-1a94-c93f-a1c0-43728d7b9f30
ms.date: 06/08/2017
---


# FileDialog.AllowMultiSelect Property (Office)

Is  **True** if the user is allowed to select multiple files from a file dialog box. Read/write.


## Syntax

 _expression_. **AllowMultiSelect**

 _expression_ A variable that represents a **FileDialog** object.


## Remarks

This property has no effect on  **Folder Picker** dialog boxes or **SaveAs** dialog boxes because users should never be able to select multiple files in these types of file dialog boxes.


## Example

The following example displays a  **File Picker** dialog box using the **FileDialog** object and displays each selected file in a message box.


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
 
 'Allow the selection of multiple files. 
 .AllowMultiSelect = True 
 
 'Use the Show method to display the file picker dialog and return the user's action. 
 'If the user presses the button... 
 If .Show = -1 Then 
 
 'Step through each string in the FileDialogSelectedItems collection. 
 For Each vrtSelectedItem In .SelectedItems 
 
 'vrtSelectedItem is aString that contains the path of each selected item. 
 'You can use any file I/O functions that you want to work with this path. 
 'This example displays the path in a message box. 
 MsgBox "Selected item's path: " &amp; vrtSelectedItem 
 
 Next 
 'If the user presses Cancel... 
 Else 
 End If 
 End With 
 
 'Set the object variable to Nothing. 
 Set fd = Nothing 
 
End Sub
```


## See also


#### Concepts


[FileDialog Object](filedialog-object-office.md)
#### Other resources


[FileDialog Object Members](filedialog-members-office.md)

