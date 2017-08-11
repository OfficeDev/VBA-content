---
title: Display and Use the File Dialog Box
ms.prod: access
ms.assetid: e4690a8b-f976-3be9-70b0-2d8c2377a19a
ms.date: 06/08/2017
---


# Display and Use the File Dialog Box

The  **[FileDialog](application-filedialog-property-access.md)** object allows you to display the file dialog box used by Access and to determine what files were selected by the user. The **[SelectedItems](../../Office-Shared-VBA/articles/filedialog-selecteditems-property-office.md)** property of the **FileDialog** object contains the paths to the files selected by the user. By using a **For...Each** loop, you can enumerate this collection and display each file; the SelectedItems property constructs and returns an iterator over the collection.

The following example adds the files selected by the user to a list box named FileList.

 **Note**  This example requires a reference to the Microsoft Office 12.0 Object Library.




```vb
Private Sub cmdFileDialog_Click() 
 
'Requires reference to Microsoft Office 12.0 Object Library. 
 
   Dim fDialog As Office.FileDialog 
   Dim varFile As Variant 
 
   'Clear listbox contents. 
   Me.FileList.RowSource = "" 
 
   'Set up the File Dialog. 
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker) 
   With fDialog 
      'Allow user to make multiple selections in dialog box. 
      .AllowMultiSelect = True 
             
      'Set the title of the dialog box. 
      .Title = "Please select one or more files" 
 
      'Clear out the current filters, and add our own. 
      .Filters.Clear 
      .Filters.Add "Access Databases", "*.MDB; *.ACCDB" 
      .Filters.Add "Access Projects", "*.ADP" 
      .Filters.Add "All Files", "*.*" 
 
      'Show the dialog box. If the .Show method returns True, the 
      'user picked at least one file. If the .Show method returns 
      'False, the user clicked Cancel. 
      If .Show = True Then 
         'Loop through each file selected and add it to the list box. 
         For Each varFile In .SelectedItems 
            Me.FileList.AddItem varFile 
         Next 
      Else 
         MsgBox "You clicked Cancel in the file dialog box." 
      End If 
   End With 
End Sub
```


