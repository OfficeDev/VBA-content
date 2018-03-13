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



| <strong>pbListTypeAiueo</strong>|
| 
<strong>pbListTypeArabic</strong>|
| 
<strong>pbListTypeArabic1</strong>|
| 
<strong>pbListTypeArabic2</strong>|
| 
<strong>pbListTypeArabicLeadingZero</strong>|
| 
<strong>pbListTypeBullet</strong>|
| 
<strong>pbListTypeCardinalText</strong>|
| 
<strong>pbListTypeChiManSty</strong>|
| 
<strong>pbListTypeChinaDbNum1</strong>|
| 
<strong>pbListTypeChinaDbNum2</strong>|
| 
<strong>pbListTypeChinaDbNum3</strong>|
| 
<strong>pbListTypeChinaDbNum4</strong>|
| 
<strong>pbListTypeChosung</strong>|
| 
<strong>pbListTypeCirclenum</strong>|
| 
<strong>pbListTypeDAiueo</strong>|
| 
<strong>pbListTypeDArabic</strong>|
| 
<strong>pbListTypeDbChar</strong>|
| 
<strong>pbListTypeDbNum1</strong>|
| 
<strong>pbListTypeDbNum2</strong>|
| 
<strong>pbListTypeDbNum3</strong>|
| 
<strong>pbListTypeDbNum4</strong>|
| 
<strong>pbListTypeDIroha</strong>|
| 
<strong>pbListTypeGanada</strong>|
| 
<strong>pbListTypeGB1</strong>|
| 
<strong>pbListTypeGB2</strong>|
| 
<strong>pbListTypeGB3</strong>|
| 
<strong>pbListTypeGB4</strong>|
| 
<strong>pbListTypeHebrew1</strong>|
| 
<strong>pbListTypeHebrew2</strong>|
| 
<strong>pbListTypeHex</strong>|
| 
<strong>pbListTypeHindi1</strong>|
| 
<strong>pbListTypeHindi2</strong>|
| 
<strong>pbListTypeHindi3</strong>|
| 
<strong>pbListTypeHindi4</strong>|
| 
<strong>pbListTypeIroha</strong>|
| 
<strong>pbListTypeKoreaDbNum1</strong>|
| 
<strong>pbListTypeKoreaDbNum2</strong>|
| 
<strong>pbListTypeKoreaDbNum3</strong>|
| 
<strong>pbListTypeKoreaDbNum4</strong>|
| 
<strong>pbListTypeLowerCaseLetter</strong>|
| 
<strong>pbListTypeLowerCaseRoman</strong>|
| 
<strong>pbListTypeLowerCaseRussian</strong>|
| 
<strong>pbListTypeNone</strong>|
| 
<strong>pbListTypeOrdinal</strong>|
| 
<strong>pbListTypeOrdinalText</strong>|
| 
<strong>pbListTypeSbChar</strong>|
| 
<strong>pbListTypeTaiwanDbNum1</strong>|
| 
<strong>pbListTypeTaiwanDbNum2</strong>|
| 
<strong>pbListTypeTaiwanDbNum3</strong>|
| 
<strong>pbListTypeTaiwanDbNum4</strong>|
| 
<strong>pbListTypeThai1</strong>|
| 
<strong>pbListTypeThai2</strong>|
| 
<strong>pbListTypeThai3</strong>|
| 
<strong>pbListTypeUpperCaseLetter</strong>|
| 
<strong>pbListTypeUpperCaseRoman</strong>|
| 
<strong>pbListTypeUpperCaseRussian</strong>|
| 
<strong>pbListTypeVietnamese1</strong>|
| 
<strong>pbListTypeZodiac1</strong>|
| 
<strong>pbListTypeZodiac2</strong>|
| 
<strong>pbListTypeZodiac3</strong>|

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


