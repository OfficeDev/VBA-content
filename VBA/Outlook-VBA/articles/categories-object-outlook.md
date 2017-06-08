---
title: Categories Object (Outlook)
keywords: vbaol11.chm3178
f1_keywords:
- vbaol11.chm3178
ms.prod: outlook
api_name:
- Outlook.Categories
ms.assetid: 319efa26-269d-9f2f-c8ec-33082e80a9e2
ms.date: 06/08/2017
---


# Categories Object (Outlook)

Represents the collection of  **[Category](http://msdn.microsoft.com/library/143ef095-54b0-cbe2-e356-632029061ac2%28Office.15%29.aspx)** objects that define the Master Category List for a namespace.


## Remarks

Microsoft Outlook provides a categorization system by which Outlook items can be easily identified and grouped into user-defined categories. The  **Categories** object represents the set of user-defined categories available to the user of a given mailbox.

Use the  **[Categories](http://msdn.microsoft.com/library/3963afca-3a7e-38d7-1347-7e1467be3a10%28Office.15%29.aspx)** property of the **[NameSpace](namespace-object-outlook.md)** object to obtain a **Categories** object reference, representing the Master Category List for that namespace.

Use the  **[Add](http://msdn.microsoft.com/library/f776c2a2-1b32-f4eb-de5e-6e245a60cac2%28Office.15%29.aspx)** method to create a new **Category** object and append it to the collection. Use the **[Item](http://msdn.microsoft.com/library/7bdf22ec-7c77-1f1f-e4fd-77bdcc0ea288%28Office.15%29.aspx)** method to obtain a **Category** object reference for an existing category, and the **[Remove](http://msdn.microsoft.com/library/8c16b02e-0297-9f36-7cb7-20e6ab0c286b%28Office.15%29.aspx)** method to remove a **Category** object from the collection. Use the **[Count](http://msdn.microsoft.com/library/b78ff508-c5c2-515c-d5f4-f4ab959f207a%28Office.15%29.aspx)** property to return the number of categories contained in the collection.


## Example

The following Visual Basic for Applications (VBA) example displays a dialog box containing the names and identifiers for each  **Category** object contained in the **Categories** collection associated with the default **[NameSpace](namespace-object-outlook.md)** object.


```
Private Sub ListCategoryIDs() 
 Dim objNameSpace As NameSpace 
 Dim objCategory As Category 
 Dim strOutput As String 
 
 ' Obtain a NameSpace object reference. 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 ' Check if the Categories collection for the Namespace 
 ' contains one or more Category objects. 
 If objNameSpace.Categories.Count > 0 Then 
 
 ' Enumerate the Categories collection. 
 For Each objCategory In objNameSpace.Categories 
 
 ' Add the name and ID of the Category object to 
 ' the output string. 
 strOutput = strOutput &amp; objCategory.Name &amp; _ 
 ": " &amp; objCategory.CategoryID &amp; vbCrLf 
 Next 
 End If 
 
 ' Display the output string. 
 MsgBox strOutput 
 
 ' Clean up. 
 Set objCategory = Nothing 
 Set objNameSpace = Nothing 
 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/f776c2a2-1b32-f4eb-de5e-6e245a60cac2%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/7bdf22ec-7c77-1f1f-e4fd-77bdcc0ea288%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/8c16b02e-0297-9f36-7cb7-20e6ab0c286b%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/7488c3e5-4163-9192-0e1d-8aa50f000978%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/3face5dd-a211-0684-eee4-e1316d4eef0c%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/b78ff508-c5c2-515c-d5f4-f4ab959f207a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/97b8f118-3846-72db-c130-4078f445d872%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/f810b08c-bf94-d4f6-563f-b0329af37f74%28Office.15%29.aspx)|

## See also


#### Other resources


[Categories Object Members](http://msdn.microsoft.com/library/36fd8906-69fa-5aa8-b026-a2de208ccd56%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
