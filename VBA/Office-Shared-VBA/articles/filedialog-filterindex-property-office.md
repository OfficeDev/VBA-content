---
title: FileDialog.FilterIndex Property (Office)
keywords: vbaof11.chm256003
f1_keywords:
- vbaof11.chm256003
ms.prod: office
api_name:
- Office.FileDialog.FilterIndex
ms.assetid: 102d3266-caab-1101-2234-68d975e11348
ms.date: 06/08/2017
---


# FileDialog.FilterIndex Property (Office)

Gets or sets a  **Long** indicating the default file filter of a file dialog box. The default filter determines which types of files are displayed when the file dialog box is first opened. Read/write.


## Syntax

 _expression_. **FilterIndex**

 _expression_ A variable that represents a **FileDialog** object.


## Remarks

If you try to set this property to a number greater than the number of filters, the last available filter will be selected.


## Example

The following example displays a  **File Picker** dialog box using the **FileDialog** object and displays each selected file in a message box. This example also demonstrates how to add a new file filter and how to make it the default filter.


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
 
 'Add a filter that includes GIF and JPEG images and make it the second item in the list. 
 .Filters.Add "Images", "*.gif; *.jpg; *.jpeg", 2 
 
 'Sets the initial file filter to number 2. 
 .FilterIndex = 2 
 
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

