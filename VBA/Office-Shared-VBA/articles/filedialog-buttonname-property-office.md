---
title: FileDialog.ButtonName Property (Office)
keywords: vbaof11.chm256005
f1_keywords:
- vbaof11.chm256005
ms.prod: office
api_name:
- Office.FileDialog.ButtonName
ms.assetid: 9f9a4f26-bd96-6e8d-099d-df15ed5e585f
ms.date: 06/08/2017
---


# FileDialog.ButtonName Property (Office)

Sets or gets a  **String** representing the text that is displayed on the action button of a file dialog box. Read/write.


## Syntax

 _expression_. **ButtonName**

 _expression_ A variable that represents a **FileDialog** object.


## Remarks

By default, this property is set to the standard text for the type of file dialog box. For example, in the case of the  **Open** dialog box, the property is set to "Open" by default. This string is limited to fifty-one characters.


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
 
 'Change the text on the button. 
 .ButtonName = "Archive" 
 
 'Use the Show method to display the File Picker dialog box and return the user's action. 
 'If the user presses the button... 
 If .Show = -1 Then 
 
 'Step through eachString in the FileDialogSelectedItems collection. 
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

