---
title: "Свойство ParagraphFormat.ListNumberSeparator (издатель)"
keywords: vbapb10.chm5439526
f1_keywords: vbapb10.chm5439526
ms.prod: publisher
api_name: Publisher.ParagraphFormat.ListNumberSeparator
ms.assetid: 63189011-12a0-c7bc-f6c6-7b17b0dcedf2
ms.date: 06/08/2017
ms.openlocfilehash: 5fdd2b99e03c7f7f2e3568a0897f026eed960d09
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlistnumberseparator-property-publisher"></a>Свойство ParagraphFormat.ListNumberSeparator (издатель)

Задает или получает **PbListSeparator** константа, представляющий список разделитель указанного абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ListNumberSeparator**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbListNumberSeparator


## <a name="remarks"></a>Заметки

Прежде чем задать свойство **ListNumberSeparator** , необходимо установить свойство **ListType** в тип нумерованного списка. Возвращает сообщение «Доступ запрещен», если список не нумерованного списка.

Значение свойства **ListNumberSeparator** может иметь одно из следующих констант **PbListSeparator** .



| **pbListSeparatorColon**|| **pbListSeparatorDoubleHyphen**|| **pbListSeparatorDoubleParen**|| **pbListSeparatorDoubleSquare**|| **pbListSeparatorParenthesis**|| **pbListSeparatorPeriod**|| **pbListSeparatorPlain**|| **pbListSeparatorSquare**|| **pbListSeparatorWideComma**|

## <a name="example"></a>Пример

В этом примере проверяется, если тип списка — нумерованный список, а именно **pbListTypeArabic**. Если свойство **ListType** имеет значение **pbListTypeArabic** **ListNumberSeparator** задано значение **pbListSeparatorParenthesis**. В противном случае вызывается метод **SetListType** и **pbListTypeArabic** передается как параметр pbListType и задайте свойство **ListNumberSeparator** .


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeArabic Then 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 Else 
 .SetListType pbListTypeArabic 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 End If 
End With 

```


