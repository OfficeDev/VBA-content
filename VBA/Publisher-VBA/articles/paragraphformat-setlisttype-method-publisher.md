---
title: ParagraphFormat.SetListType Method (Publisher)
keywords: vbapb10.chm5439520
f1_keywords:
- vbapb10.chm5439520
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.SetListType
ms.assetid: 6900aac5-fb3f-5813-309c-1422d38c8301
ms.date: 06/08/2017
---


# ParagraphFormat.SetListType Method (Publisher)

Sets the list type of the specified  **ParagraphFormat** object. .


## Syntax

 _expression_. **SetListType**( **_Value_**,  **_BulletText_**)

 _expression_A variable that represents a  **ParagraphFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **PbListType**|Represents the list type of the specified  **ParagraphFormat** object.|
|BulletText|Optional| **String**| **String** that represents the text of the list bullet.|

## Remarks

If Value is a bulleted list and the BulletText parameter is missing, the first bullet from the  **Bullets and Numbering** dialog box is used.

BulletText is limited to one character. 

A run-time error occurs if the BulletText parameter is provided and the Value parameter is not set to  **pbListTypeBullet**.

The Value parameter can be one of the  **PbListType** constants declared in the Microsoft Publisher type library and shown in the following table.



| **pbListTypeAiueo**|
| **pbListTypeArabic**|
| **pbListTypeArabic1**|
| **pbListTypeArabic2**|
| **pbListTypeArabicLeadingZero**|
| **pbListTypeBullet**|
| **pbListTypeCardinalText**|
| **pbListTypeChiManSty**|
| **pbListTypeChinaDbNum1**|
| **pbListTypeChinaDbNum2**|
| **pbListTypeChinaDbNum3**|
| **pbListTypeChinaDbNum4**|
| **pbListTypeChosung**|
| **pbListTypeCirclenum**|
| **pbListTypeDAiueo**|
| **pbListTypeDArabic**|
| **pbListTypeDbChar**|
| **pbListTypeDbNum1**|
| **pbListTypeDbNum2**|
| **pbListTypeDbNum3**|
| **pbListTypeDbNum4**|
| **pbListTypeDIroha**|
| **pbListTypeGanada**|
| **pbListTypeGB1**|
| **pbListTypeGB2**|
| **pbListTypeGB3**|
| **pbListTypeGB4**|
| **pbListTypeHebrew1**|
| **pbListTypeHebrew2**|
| **pbListTypeHex**|
| **pbListTypeHindi1**|
| **pbListTypeHindi2**|
| **pbListTypeHindi3**|
| **pbListTypeHindi4**|
| **pbListTypeIroha**|
| **pbListTypeKoreaDbNum1**|
| **pbListTypeKoreaDbNum2**|
| **pbListTypeKoreaDbNum3**|
| **pbListTypeKoreaDbNum4**|
| **pbListTypeLowerCaseLetter**|
| **pbListTypeLowerCaseRoman**|
| **pbListTypeLowerCaseRussian**|
| **pbListTypeNone**|
| **pbListTypeOrdinal**|
| **pbListTypeOrdinalText**|
| **pbListTypeSbChar**|
| **pbListTypeTaiwanDbNum1**|
| **pbListTypeTaiwanDbNum2**|
| **pbListTypeTaiwanDbNum3**|
| **pbListTypeTaiwanDbNum4**|
| **pbListTypeThai1**|
| **pbListTypeThai2**|
| **pbListTypeThai3**|
| **pbListTypeUpperCaseLetter**|
| **pbListTypeUpperCaseRoman**|
| **pbListTypeUpperCaseRussian**|
| **pbListTypeVietnamese1**|
| **pbListTypeZodiac1**|
| **pbListTypeZodiac2**|
| **pbListTypeZodiac3**|

## Example

This example tests to see if the list type is a numbered list, specifically  **pbListTypeArabic**. If the  **ListType** property is set to **pbListTypeArabic**, the  **ListSeparator** is set to **pbListSeparatorParenthesis**. Otherwise the  **SetListType** method is called and passed **pbListTypeArabic** as the Value parameter and then the **ListNumberSeparator** property can be set.


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeArabic Then 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 Else 
 .SetListType pbListTypeArabic 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 End If 
End With 
 

```

This example demonstrates how an organized document structure containing named text frames with lists can be configured. This example assumes that the publication has a naming convention for  **TextFrame** objects containing lists that use the word "list" as a prefix. This example uses nested collection iterations to access each of the **TextFrame** objects in each **Shapes** collection of each **Page**. The  **ParagraphFormat** object of each **TextFrame** name with the prefix "list" has the **ListType** and **ListBulletFontSize** set.




```vb
Dim objPage As page 
Dim objShp As Shape 
Dim objTxtFrm As TextFrame 
 
'Iterate through all pages of th ePublication 
For Each objPage In ActiveDocument.Pages 
 'Iterate through the Shapes collection of objPage 
 For Each objShp In objPage.Shapes 
 'Find each TextFrame object 
 If objShp.Type = pbTextFrame Then 
 'If the name of the TextFrame begins with "list" 
 If InStr(1, objShp.Name, "list") <> 0 Then 
 Set objTxtFrm = objShp.TextFrame 
 With objTxtFrm 
 With .TextRange 
 With .ParagraphFormat 
 .SetListType pbListTypeBullet, "*" 
 .ListBulletFontSize = 24 
 End With 
 End With 
 End With 
 End If 
 End If 
 Next 
Next 
 

```


