---
title: FileDialog Object (Office)
keywords: vbaof11.chm256000
f1_keywords:
- vbaof11.chm256000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.FileDialog
ms.assetid: 71a030f2-3b02-21e1-c156-0514ff5eddb7
---


# FileDialog Object (Office)

Provides file dialog box functionality similar to the functionality of the standard  **Open** and **Save** dialog boxes found in Microsoft Office applications.


## Remarks

Use the  **FileDialog** property to return a **FileDialog** object. The **FileDialog** property is located in each individual Office application's **Application** object. The property takes a single argument, _DialogType_, that determines the type of **FileDialog** object that the property returns. There are four types of **FileDialog** object:


-  **Open** dialog box - lets users select one or more files that you can then open in the host application using the **Execute** method.
    
-  **SaveAs** dialog box - lets users select a single file that you can then save the current file as using the **Execute** method.
    
-  **File Picker** dialog box - lets users select one or more files. The file paths that the user selects are captured in the **FileDialogSelectedItems** collection.
    
-  **Folder Picker** dialog box - lets users select a path. The path that the user selects is captured in the **FileDialogSelectedItems** collection.
    
Each host application can only create a single instance of the  **FileDialog** object. Therefore, many of the properties of the **FileDialog** object persist even when you create multiple **FileDialog** objects. Therefore, make sure that you set all of the properties appropriately for your purpose before you display the dialog box.


## Example

To display a file dialog box using the  **FileDialog** object, you must use the **Show** method. Once a dialog box is displayed, no code executes until the user dismisses the dialog box. The following example creates and displays a **File Picker** dialog box and then displays each selected file in a message box.


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
 
 'Use the Show method to display the File Picker dialog box and return the user's action. 
 'The user pressed the button. 
 If .Show = -1 Then 
 
 'Step through each string in the FileDialogSelectedItems collection. 
 For Each vrtSelectedItem In .SelectedItems 
 
 'vrtSelectedItem is aString that contains the path of each selected item. 
 'You can use any file I/O functions that you want to work with this path. 
 'This example displays the path in a message box. 
 MsgBox "The path is: " &amp; vrtSelectedItem 
 
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
|[Execute](http://msdn.microsoft.com/library/63899b0e-51d4-f20a-b114-c713d8743527%28Office.15%29.aspx)|
|[Show](http://msdn.microsoft.com/library/e67f7fc3-326d-12d0-fe44-e20048ff6abf%28Office.15%29.aspx)|

## Properties


<<<<<<< HEAD

|**Name**|
|:-----|
|[AllowMultiSelect](http://msdn.microsoft.com/library/b109b0b5-1a94-c93f-a1c0-43728d7b9f30%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/797e82c7-0737-03ae-7df3-7178bc6ff328%28Office.15%29.aspx)|
|[ButtonName](http://msdn.microsoft.com/library/9f9a4f26-bd96-6e8d-099d-df15ed5e585f%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/682d031d-8107-8a89-4cb1-6cbe8524fc95%28Office.15%29.aspx)|
|[DialogType](http://msdn.microsoft.com/library/c589fe49-6527-7cdc-b7cb-55ac71013f3c%28Office.15%29.aspx)|
|[FilterIndex](http://msdn.microsoft.com/library/102d3266-caab-1101-2234-68d975e11348%28Office.15%29.aspx)|
|[Filters](http://msdn.microsoft.com/library/0aef7760-a618-c20c-0816-98be1b93e564%28Office.15%29.aspx)|
|[InitialFileName](http://msdn.microsoft.com/library/900970fe-1331-9b0a-3182-953cb6b583ce%28Office.15%29.aspx)|
|[InitialView](http://msdn.microsoft.com/library/17950503-6511-8159-7f9f-406dd22e4fca%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/e29dab4e-4226-32bf-f4c2-3afaeb0e3616%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/c305bcd3-dc42-f84e-abc2-1ee4a1092ef8%28Office.15%29.aspx)|
|[SelectedItems](http://msdn.microsoft.com/library/af45013a-c745-3f14-9c12-64a1c2b50279%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/a2d43a1d-78ce-3f8f-7763-7324e5af183d%28Office.15%29.aspx)|

## See also


#### Other resources

[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
=======
[Object Model Reference](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

