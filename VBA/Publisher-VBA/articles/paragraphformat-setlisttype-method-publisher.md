---
title: "Метод ParagraphFormat.SetListType (издатель)"
keywords: vbapb10.chm5439520
f1_keywords: vbapb10.chm5439520
ms.prod: publisher
api_name: Publisher.ParagraphFormat.SetListType
ms.assetid: 6900aac5-fb3f-5813-309c-1422d38c8301
ms.date: 06/08/2017
ms.openlocfilehash: 34c502c6aea5d24f6ee352993d6fc4fe4e427b98
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatsetlisttype-method-publisher"></a>Метод ParagraphFormat.SetListType (издатель)

Задает тип списка на указанный объект **ParagraphFormat** . .


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetListType** ( **_Значение_**, **_BulletText_**)

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **PbListType**|Представляет тип списка на указанный объект **ParagraphFormat** .|
|BulletText|Необязательный| **String**| **Строка** , представляющая текст маркированный список.|

## <a name="remarks"></a>Заметки

Если значение равно маркированный список, а параметр BulletText отсутствует, используется первый маркер из диалогового окна **список** .

BulletText применяется только для одного символа. 

Если параметр BulletText указан, значение параметра не задано значение **pbListTypeBullet**, возникает ошибка времени выполнения.

Значение параметра может иметь одно из **PbListType** константы объявляются в библиотеке типов Microsoft Publisher и показаны в следующей таблице.



| **pbListTypeAiueo**|| **pbListTypeArabic**|| **pbListTypeArabic1**|| **pbListTypeArabic2**|| **pbListTypeArabicLeadingZero**|| **pbListTypeBullet**|| **pbListTypeCardinalText**|| **pbListTypeChiManSty**|| **pbListTypeChinaDbNum1**|| **pbListTypeChinaDbNum2**|| **pbListTypeChinaDbNum3**|| **pbListTypeChinaDbNum4**|| **pbListTypeChosung**|| **pbListTypeCirclenum**|| **pbListTypeDAiueo**|| **pbListTypeDArabic**|| **pbListTypeDbChar**|| **pbListTypeDbNum1**|| **pbListTypeDbNum2**|| **pbListTypeDbNum3**|| **pbListTypeDbNum4**|| **pbListTypeDIroha**|| **pbListTypeGanada**|| **pbListTypeGB1**|| **pbListTypeGB2**|| **pbListTypeGB3**|| **pbListTypeGB4**|| **pbListTypeHebrew1**|| **pbListTypeHebrew2**|| **pbListTypeHex**|| **pbListTypeHindi1**|| **pbListTypeHindi2**|| **pbListTypeHindi3**|| **pbListTypeHindi4**|| **pbListTypeIroha**|| **pbListTypeKoreaDbNum1**|| **pbListTypeKoreaDbNum2**|| **pbListTypeKoreaDbNum3**|| **pbListTypeKoreaDbNum4**|| **pbListTypeLowerCaseLetter**|| **pbListTypeLowerCaseRoman**|| **pbListTypeLowerCaseRussian**|| **pbListTypeNone**|| **pbListTypeOrdinal**|| **pbListTypeOrdinalText**|| **pbListTypeSbChar**|| **pbListTypeTaiwanDbNum1**|| **pbListTypeTaiwanDbNum2**|| **pbListTypeTaiwanDbNum3**|| **pbListTypeTaiwanDbNum4**|| **pbListTypeThai1**|| **pbListTypeThai2**|| **pbListTypeThai3**|| **pbListTypeUpperCaseLetter**|| **pbListTypeUpperCaseRoman**|| **pbListTypeUpperCaseRussian**|| **pbListTypeVietnamese1**|| **pbListTypeZodiac1**|| **pbListTypeZodiac2**|| **pbListTypeZodiac3**|

## <a name="example"></a>Пример

В этом примере проверяется, если тип списка — нумерованный список, а именно **pbListTypeArabic**. Если свойство **ListType** **pbListTypeArabic**, **ListSeparator** задано значение **pbListSeparatorParenthesis**. В противном случае метод **SetListType** вызван и передается в качестве значения параметра **pbListTypeArabic** и задайте свойство **ListNumberSeparator** .


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

В этом примере показано, как можно настроить структуру упорядоченный документа, содержащую именованные текстовые рамки со списками. В этом примере предполагается, что публикация содержит схему именования **TextFrame** объекты, содержащие списки, использующих word «список» с префиксом. В этом примере использует итерации вложенного набора для доступа к объектам **TextFrame** в каждой коллекции **фигур** каждой **страницы**. Объект **ParagraphFormat** имени каждого **TextFrame** с префикса «список» имеет **ListType** и **ListBulletFontSize** задания.




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


