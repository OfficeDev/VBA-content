
# Access the Values of a Multi-valued Property in a Table

 **Last modified:** July 28, 2015

 _**Applies to:** Outlook 2013_

Generally, if a multi-valued property is added to a  ** [Table](0affaafd-93fe-227a-acee-e09a86cadc20.md)** using its explicit built-in name, the format of the values of the property is a comma-delimited string. If the multi-valued property is added to the **Table** using a reference by namespace, the format of the values of the property is a variant array.

The following code sample adds the multi-valued  **Categories** property to a **Table** using a name that references its namespace, **urn:schemas-microsoft-com:office:office#Keywords**. To obtain the values for the  **Categories** column for each row in the **Table**, assign



```
oRow("urn:schemas-microsoft-com:office:office#Keywords")
```

to a variant, and enumerate the elements of the variant array. Note that for an item that has not been assigned any categories, to a variant, and enumerate the elements of the variant array. Note that for an item that has not been assigned any categories, 



```
oRow("urn:schemas-microsoft-com:office:office#Keywords")
```

returns an Empty value.



```
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
                strCategories = strCategories &amp; (varCat(j)) &amp; ", " 
            Next 
            'Remove last trailing ", " 
            strCategories = Left(strCategories, Len(strCategories) - 2) 
        Else 
            'The item does not have any categories 
            strCategories = "" 
        End If 
        Debug.Print ("Subject: " _ 
           &amp; oRow("Subject") &amp; vbCrLf &amp; "Categories: ") &amp; strCategories &amp; vbCrLf 
    Loop 
End Sub
```

