
# Table.ConvertToText Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Converts a table to text and returns a  **Range** object that represents the delimited text.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ConvertToText**( **_Separator_**,  **_NestedTables_**)

 _expression_Required. A variable that represents a  ** [Table](996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Separator|Optional| **Variant**|The character that delimits the converted columns (paragraph marks delimit the converted rows). Can be any  **WdTableFieldSeparator** constants.|
|NestedTables|Optional| **Variant**| **True** if nested tables are converted to text. This argument is ignored if Separator is not **wdSeparateByParagraphs**. The default value is  **True**.|

## Remarks
<a name="sectionSection1"> </a>

When you apply the  **ConvertToText** method to a **Table** object, the object is deleted. To maintain a reference to the converted contents of the table, you must assign the **Range** object returned by the **ConvertToText** method to a new object variable. In the following example, the first table in the active document is converted to text and then formatted as a bulleted list.


```
Dim tableTemp As Table 
Dim rngTemp As Range 
 
Set tableTemp = ActiveDocument.Tables(1) 
Set rngTemp = _ 
 tableTemp.ConvertToText(Separator:=wdSeparateByParagraphs) 
 
rngTemp.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdBulletGallery).ListTemplates(1)
```


## Example
<a name="sectionSection2"> </a>

This example creates a table and then converts it to text by using tabs as separator characters.


```
Dim docNew As Document 
Dim tableNew As Table 
Dim intTemp As Integer 
Dim cellLoop As Cell 
Dim rngTemp As Range 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=3) 
 
intTemp = 1 
 
For Each cellLoop In tableNew.Range.Cells 
 cellLoop.Range.InsertAfter "Cell " &amp; intTemp 
 intTemp = intTemp + 1 
Next cellLoop 
 
MsgBox "Click OK to convert table to text." 
Set rngTemp = _ 
 tableNew.ConvertToText(Separator:=wdSeparateByTabs)
```

This example converts the table that contains the selection to text, with spaces between the columns.




```
If Selection.Information(wdWithInTable) = True Then 
 Selection.Tables(1).ConvertToText Separator:=" " 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Table Object](996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6.md)
#### Other resources


 [Table Object Members](5367ee92-b5a3-92c7-787b-46a302586a0d.md)
