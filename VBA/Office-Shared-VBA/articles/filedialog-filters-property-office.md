---
title: FileDialog.Filters Property (Office)
keywords: vbaof11.chm256002
f1_keywords:
- vbaof11.chm256002
ms.prod: office
api_name:
- Office.FileDialog.Filters
ms.assetid: 0aef7760-a618-c20c-0816-98be1b93e564
ms.date: 06/08/2017
---


# FileDialog.Filters Property (Office)

Gets a  **FileDialogFilters** collection. Read-only.


## Syntax

 _expression_. **Filters**

 _expression_ A variable that represents a **FileDialog** object.


### Return Value

FileDialogFilters


## Example

The following example displays a  **File Picker** dialog box using the **FileDialog** object and displays each selected file in a message box. The example also adds a new file filter called "Images."


```
Sub Main() 
 
 'Declare a variable as a FileDialog object. 
 Dim fd As FileDialog 
 
 'Create a FileDialog object as a File Picker dialog. 
 Set fd = Application.FileDialog(msoFileDialogFilePicker) 
 
 'Declare a variable to contain the path 
 'of each selected item. Even though the path is aString, 
 'the variable must be a Variant because For Each...Next 
 'routines only work with Variants and Objects. 
 Dim vrtSelectedItem As Variant 
 
 'Use a With...End With block to reference the FileDialog object. 
 With fd 
 
 'Add a filter that includes GIF and JPEG images and make it the first item in the list. 
 .Filters.Add "Images", "*.gif; *.jpg; *.jpeg", 1 
 
 'Use the Show method to display the File Picker dialog box and return the user's action. 
 'If the user presses the button... 
 If .Show = -1 Then 
 
 'Step through each string in the FileDialogSelectedItems collection. 
 For Each vrtSelectedItem In .SelectedItems 
 
 'vrtSelectedItem is aString that contains the path of each selected item. 
 'You can use any file I/O functions that you want to work with this path. 
 'This example displays the path in a message box. 
 MsgBox "Selected item's path: " &amp; vrtSelectedItem 
 
 Next vrtSelectedItem 
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

