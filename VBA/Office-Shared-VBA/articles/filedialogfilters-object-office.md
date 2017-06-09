---
title: FileDialogFilters Object (Office)
keywords: vbaof11.chm255000
f1_keywords:
- vbaof11.chm255000
ms.prod: office
api_name:
- Office.FileDialogFilters
ms.assetid: a74663cf-ad63-e41a-8d5e-e51e8a20c173
ms.date: 06/08/2017
---


# FileDialogFilters Object (Office)

A collection of  **FileDialogFilter** objects that represent the types of files that can be selected in a file dialog box that is displayed using the **FileDialog** object.


## Example

Use the  **Filters** property of the **FileDialog** object to return a **FileDialogFilters** collection. The following code returns the **FileDialogFilters** collection for the **File Open** dialog box.


```
Application.FileDialog(msoFileDialogOpen).Filters
```

Use the  **Add** method to add **FileDialogFilter** objects to the **FileDialogFilters** collection. The following example uses the **Clear** method to clear the collection and then adds filters to the collection. The **Clear** method completely empties the collection; however, if you do not add any filters to the collection after you clear it, the "All files (*.*)" filter is added automatically.




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
 
        'Change the contents of the Files of Type list. 
        'Empty the list by clearing the FileDialogFilters collection. 
        .Filters.Clear 
 
        'Add a filter that includes all files. 
        .Filters.Add "All files", "*.*" 
 
        'Add a filter that includes GIF and JPEG images and make it the first item in the list. 
        .Filters.Add "Images", "*.gif; *.jpg; *.jpeg", 1 
 
        'Use the Show method to display the File Picker dialog box and return the user's action. 
        'The user pressed the button. 
        If .Show = -1 Then 
 
            'Step through eachString in the FileDialogSelectedItems collection. 
            For Each vrtSelectedItem In .SelectedItems 
 
                'vrtSelectedItem is aString that contains the path of each selected item. 
                'You can use any file I/O functions that you want to work with this path. 
                'This example displays the path in a message box. 
                MsgBox "Path name: " &amp; vrtSelectedItem 
 
            Next vrtSelectedItem 
        'The user pressed Cancel. 
        Else 
        End If 
    End With 
 
    'Set the object variable to Nothing. 
    Set fd = Nothing 
 
End Sub
```

When changing the  **FileDialogFilters** collection, remember that each application can only create an instance of a single **FileDialog** object. This means that the **FileDialogFilters** collection resets to its default filters whenever you call the **FileDialog** method with a new dialog box type. The following example iterates through the default filters of the **SaveAs** dialog box and displays the description of each filter that includes a Microsoft Excel file.




```
Sub Main() 
 
    'Declare a variable as a FileDialogFilters collection. 
    Dim fdfs As FileDialogFilters 
 
    'Declare a variable as a FileDialogFilter object. 
    Dim fdf As FileDialogFilter 
 
    'Set the FileDialogFilters collection variable to 
    'the FileDialogFilters collection of the SaveAs dialog box. 
    Set fdfs = Application.FileDialog(msoFileDialogSaveAs).Filters 
 
    'Iterate through the description and extensions of each 
    'default filter in the SaveAs dialog box. 
    For Each fdf In fdfs 
 
        'Display the description of filters that include 
        'Microsoft Excel files 
        If InStr(1, fdf.Extensions, "xls", vbTextCompare) > 0 Then 
            MsgBox "Description of filter: " &amp; fdf.Description 
        End If 
    Next fdf 
 
End Sub
```


 **Note**  A run-time error will occur if the  **Filters** property is used in conjunction with the **Clear**, **Add**, or **Delete** methods when applied to a Save As **FileDiaog** object. For example, `Application.FileDialog(msoFileDialogSaveAs).Filters.Clear` will result in a run-time error.


## Methods



|**Name**|
|:-----|
|[Add](filedialogfilters-add-method-office.md)|
|[Clear](filedialogfilters-clear-method-office.md)|
|[Delete](filedialogfilters-delete-method-office.md)|
|[Item](filedialogfilters-item-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](filedialogfilters-application-property-office.md)|
|[Count](filedialogfilters-count-property-office.md)|
|[Creator](filedialogfilters-creator-property-office.md)|
|[Parent](filedialogfilters-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
