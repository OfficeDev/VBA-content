---
title: ListObject.SharePointURL Property (Excel)
keywords: vbaxl10.chm734095
f1_keywords:
- vbaxl10.chm734095
ms.prod: excel
api_name:
- Excel.ListObject.SharePointURL
ms.assetid: a5b19612-c8e8-4952-e15c-a60da10f65d1
ms.date: 06/08/2017
---


# ListObject.SharePointURL Property (Excel)

 Returns a **String** representing the URL of the SharePoint list for a given **[ListObject](listobject-object-excel.md)** object. Read-only **String** .


## Syntax

 _expression_ . **SharePointURL**

 _expression_ A variable that represents a **ListObject** object.


## Remarks

Accessing this property generates a run-time error if the list is not linked to a SharePoint site.


## Example

The following example sets elements of the  _Target_ parameter of the **[Publish](listobject-publish-method-excel.md)** method to push the **ListObject** object to a SharePoint site. The code sample uses the **SharePointURL** property to assign the URL to the array and the **Name** property to assign the name of the list. The information in the array is then passed to the SharePoint site using the **Publish** method.


```vb
Sub PublishList() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim arTarget(4) As String 
 Dim strSTSConnection As String 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 
 arTarget(0) = "0" 
 arTarget(1) = objListObj.SharePointURL 
 arTarget(2) = "1" 
 arTarget(3) = objListObj.Name 
 
 strSTSConnection = objListObj.Publish(arTarget, True) 
End Sub
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

