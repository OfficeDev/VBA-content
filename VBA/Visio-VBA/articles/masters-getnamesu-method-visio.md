---
title: Masters.GetNamesU Method (Visio)
keywords: vis_sdr.chm10851940
f1_keywords:
- vis_sdr.chm10851940
ms.prod: visio
api_name:
- Visio.Masters.GetNamesU
ms.assetid: bf797a6a-1018-eda6-41e8-c8533638a034
ms.date: 06/08/2017
---


# Masters.GetNamesU Method (Visio)

Returns the universal names of all items in a collection.


## Syntax

 _expression_ . **GetNamesU**( **_localeIndependentNameArray()_** )

 _expression_ A variable that represents a **Masters** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _localeIndependentNameArray()_|Required| **String**|Out parameter. An array that receives the names of members of the indicated object.|

### Return Value

Nothing


## Remarks

If the  **GetNamesU** method succeeds, _localeIndependentNameArray()_ returns a one-dimensional array of _n_ strings indexed from 0 to _n_ - 1, where _n_ equals the **Count** property of the object. The _localeIndependentNameArray()_ parameter is an out parameter that is allocated by the **GetNamesU** method, which passes ownership back to the caller. The caller should eventually perform the **SafeArrayDestroy** procedure on the returned array. Note that the **SafeArrayDestroy** procedure has the side effect of freeing the strings referenced by the array's entries. (Microsoft Visual Basic and Microsoft Visual Basic for Applications (VBA) take care of this for you.)


 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **GetNames** method to get the local name of more than one object. Use the **GetNamesU** method to get the universal name of more than one object.


## Example

The following VBA macro shows how to use the  **GetNamesU** method to get the names of all the **Master** objects in the **Masters** collection of the active document and print them in the Immediate window.


```vb
 
Public Sub GetNamesU_Example() 
 
 Dim strMasterNames() As String 
 Dim intLowerBound As Integer 
 Dim intUpperBound As Integer 
 
 ActiveDocument.Masters.GetNamesU strMasterNames 
 intLowerBound = LBound(strMasterNames) 
 intUpperBound = UBound(strMasterNames) 
 Debug.Print ActiveDocument; " Lower bound:"; intLowerBound; "Upper bound:"; intUpperBound 
 
 While intLowerBound <= intUpperBound 
 
 Debug.Print strMasterNames (intLowerBound) 
 intLowerBound = intLowerBound + 1 
 
 Wend 
 
End Sub
```


