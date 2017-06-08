---
title: Access the Values of a Multi-valued Property in a Table
ms.prod: outlook
ms.assetid: e914b32b-d290-705b-d4fc-fecfba85fd8b
ms.date: 06/08/2017
---


# Access the Values of a Multi-valued Property in a Table

Generally, if a multi-valued property is added to a  **[Table](table-object-outlook.md)** using its explicit built-in name, the format of the values of the property is a comma-delimited string. If the multi-valued property is added to the **Table** using a reference by namespace, the format of the values of the property is a variant array.

The following code sample adds the multi-valued  **Categories** property to a **Table** using a name that references its namespace, **urn:schemas-microsoft-com:office:office#Keywords**. To obtain the values for the  **Categories** column for each row in the **Table**, assign



```
oRow("urn:schemas-microsoft-com:office:office#Keywords")
```

to a variant, and enumerate the elements of the variant array. Note that for an item that has not been assigned any categories, to a variant, and enumerate the elements of the variant array. Note that for an item that has not been assigned any categories, 



```
oRow("urn:schemas-microsoft-com:office:office#Keywords")
```

returns an Empty value.



```vb
Sub TableCategories() 
    Dim oT As Outlook.Table 
    Dim oRow As Outlook.Row 
    Dim varCat 
    Dim j As Integer 
    Dim strCategories As String 
 
    Set oT = Application.ActiveExplorer.CurrentFolder.GetTable() 
    oT.Columns.Add ("urn:schemas-microsoft-com:office:office#Keywords") 
    oT.Sort "LastModificationTime", True 
    Do Until oT.EndOfTable 
        Set oRow = oT.GetNextRow 
        'Obtain any values of the Categories property 
        varCat = oRow("urn:schemas-microsoft-com:office:office#Keywords") 
        If Not (IsEmpty(varCat)) Then 
            'Form a string out of the item's categories 
            For j = 0 To UBound(varCat) 
                strCategories = strCategories &; (varCat(j)) &; ", " 
            Next 
            'Remove last trailing ", " 
            strCategories = Left(strCategories, Len(strCategories) - 2) 
        Else 
            'The item does not have any categories 
            strCategories = "" 
        End If 
        Debug.Print ("Subject: " _ 
           &; oRow("Subject") &; vbCrLf &; "Categories: ") &; strCategories &; vbCrLf 
    Loop 
End Sub
```


