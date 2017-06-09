---
title: FileDialog.InitialFileName Property (Office)
keywords: vbaof11.chm256008
f1_keywords:
- vbaof11.chm256008
ms.prod: office
api_name:
- Office.FileDialog.InitialFileName
ms.assetid: 900970fe-1331-9b0a-3182-953cb6b583ce
ms.date: 06/08/2017
---


# FileDialog.InitialFileName Property (Office)

Set or returns a  **String** representing the path or file name that is initially displayed in a file dialog box. Read/write.


## Syntax

 _expression_. **InitialFileName**

 _expression_ A variable that represents a **FileDialog** object.


## Remarks

You can use the  **'*'** and **'?'** wildcard characters when specifying the file name but not when specifying the path. The **'*'** symbol represents any number of consecutive characters and the **'?'** represents a single character. For example, **.InitialFileName = "c:\c*s.txt"** returns both "charts.txt" and "checkregister.txt."

If you specify a path and no file name, then all files that are allowed by the file filter appear in the dialog box.

If you specify a file that exists in the initial folder, then only that file appears in the dialog box.

If you specify a file name that does not exist in the initial folder, then the dialog box contains no files. The type of file that you specify in the  **InitialFileName** property overrides the file filter settings.

If you specify an invalid path, the last-used path is used. A message warns users when an invalid path is used.

Setting this property to a string longer than 256 characters causes a run-time error.


## Example

The following example displays a  **File Picker** dialog box using the **FileDialog** object and displays each selected file in a message box.


```
Sub Main() 
 
 'Declare a variable as a FileDialog object 
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
 
 'Set the initial path to the C:\ drive. 
 .InitialFileName = "C:\" 
 
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

