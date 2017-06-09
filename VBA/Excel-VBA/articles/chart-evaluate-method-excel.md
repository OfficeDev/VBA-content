---
title: Chart.Evaluate Method (Excel)
keywords: vbaxl10.chm149107
f1_keywords:
- vbaxl10.chm149107
ms.prod: excel
api_name:
- Excel.Chart.Evaluate
ms.assetid: 7a171fd5-e084-7172-f429-5425e0d342d4
ms.date: 06/08/2017
---


# Chart.Evaluate Method (Excel)

Converts a Microsoft Excel name to an object or a value.


## Syntax

 _expression_ . **Evaluate**( **_Name_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **Variant**|The name of the object, using the naming convention of Microsoft Excel.|

### Return Value

Variant


## Remarks

The following types of names in Microsoft Excel can be used with this method:


- A1-style references. You can use any reference to a single cell in A1-style notation. All references are considered to be absolute references.
    
- Ranges. You can use the range, intersect, and union operators (colon, space, and comma, respectively) with references.
    
- Defined names. You can specify any name in the language of the macro.
    
- External references. You can use the ! operator to refer to a cell or to a name defined in another workbook ? for example,  `Evaluate("[BOOK1.XLS]Sheet1!A1")`.
    
- Chart Objects. You can specify any chart object name, such as "Legend", "Plot Area", or "Series 1", to access the properties and methods of that object. For example,  `Charts("Chart1").Evaluate("Legend").Font.Name` returns the name of the font used in the legend.
    

 **Note**  Using square brackets (for example, "[A1:C5]") is identical to calling the  **Evaluate** method with a string argument. For example, the following expression pairs are equivalent.


```vb
[a1].Value = 25 
Evaluate("A1").Value = 25 
 
trigVariable = [SIN(45)] 
trigVariable = Evaluate("SIN(45)") 
 
Set firstCellInSheet = Workbooks("BOOK1.XLS").Sheets(4).[A1] 
Set firstCellInSheet = _ 
 Workbooks("BOOK1.XLS").Sheets(4).Evaluate("A1")
```

The advantage of using square brackets is that the code is shorter. The advantage of using  **Evaluate** is that the argument is a string, so you can either construct the string in your code or use a Visual Basic variable.


## Example

This example turns on bold formatting in cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Activate 
boldCell = "A1" 
Application.Evaluate(boldCell).Font.Bold = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

