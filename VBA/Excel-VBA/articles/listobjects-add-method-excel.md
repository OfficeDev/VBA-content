---
title: ListObjects.Add Method (Excel)
keywords: vbaxl10.chm732078
f1_keywords:
- vbaxl10.chm732078
ms.prod: EXCEL
api_name:
- Excel.ListObjects.Add
ms.assetid: 764dafed-d4e3-82b9-df8c-68a358319491
---


# ListObjects.Add Method (Excel)

Creates a new list object.


## Syntax

 _expression_ . **Add**( **_SourceType_** , **_Source_** , **_LinkSource_** , **_HasHeaders_** , **_Destination_** )

 _expression_ A variable that represents a **ListObjects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceType_|Optional| **[XlListObjectSourceType](xllistobjectsourcetype-enumeration-excel.md)**|Indicates the kind of source for the query. |
| _Source_|Optional| **Variant**|when SourceType =  **xlSrcRange** . A **[Range](range-object-excel.md)** object representing the data source. If omitted, the Source will default to the range returned by list range detection code. when SourceType = **xlSrcExternal** . An array of **String** values specifying a connection to the source, containing the following elements:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>0 - URL to SharePoint site</p></li><li><p>1 - ListName</p></li><li><p>2 - ViewGUID</p></li></ul>|
| _LinkSource_|Optional| **Variant**| **Boolean** . Indicates whether an external data source is to be linked to the **ListObject** object. If SourceType is **xlSrcExternal** , default is **True** . Invalid if SourceType is **xlSrcRange** , and will return an error if not omitted.|
| _TableStyleName_|Optional| **Variant**|An  **[XlYesNoGuess](xlyesnoguess-enumeration-excel.md)** constant that indicates whether the data being imported has column labels. If the Source does not contain headers, Excel will automatically generate headers.|
| _Destination_|Optional| **Variant**|A  **[Range](range-object-excel.md)** object specifying a single-cell reference as the destination for the top-left corner of the new list object. If the **Range** object refers to more than one cell, an error is generated. The Destination argument must be specified when SourceType is set to **xlSrcExternal** . The Destination argument is ignored if SourceType is set to **xlSrcRange** . The destination range must be on the worksheet that contains the **[ListObjects](listobjects-object-excel.md)** collection specified by expression. New columns will be inserted at the Destination to fit the new list. Therefore, existing data will not be overwritten.|

### Return Value

A  **[ListObject](listobject-object-excel.md)** object that represents the new list object.


## Remarks

When the list has headers, the first row of cells will be converted to  **Text** , if not already set to text. The conversion will be based on the visible text for the cell. This means that if there is a date value with a **Date** format that changes with locale, the conversion to a list might produce different results depending on the current system locale. Moreover, if there are two cells in the header row that have the same visible text, an incremental **Integer** will be appended to make each column header unique.






## Example

The following example adds a new  **ListObject** object based on data from a Microsoft SharePoint Foundation site to the default **ListObjects** collection and places the list in cell A1 in the first worksheet of the workbook.


 **Note**  The following code example assumes that you will substitute a valid server name and the list guid in the variables  `strServerName` and `strListGUID`. Additionally, the server name must be followed by "/_vti_bin" or the sample will not work.


```vb
Set objListObject = ActiveWorkbook.Worksheets(1).ListObjects.Add(SourceType:= xlSrcExternal, _ 
Source:= Array(strServerName, strListName, strListGUID), LinkSource:=True, _ 
TableStyleName:=xlGuess, Destination:=Range("A10")) 

```


## See also


#### Concepts


[ListObjects Object](listobjects-object-excel.md)

