---
title: "Свойство ParagraphFormat.ListType (издатель)"
keywords: vbapb10.chm5439521
f1_keywords: vbapb10.chm5439521
ms.prod: publisher
api_name: Publisher.ParagraphFormat.ListType
ms.assetid: 04ae7157-e864-4e95-74ff-59821eceb286
ms.date: 06/08/2017
ms.openlocfilehash: ad9667186a1fe8e273e56ca623cc3ea3bda05382
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlisttype-property-publisher"></a>Свойство ParagraphFormat.ListType (издатель)

Возвращает константу **PbListType** из указанного объекта **ParagraphFormat** . Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ListType**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

pbListType


## <a name="remarks"></a>Заметки

Это свойство доступно только для чтения. Чтобы задать свойство **ListType** объекта **ParagraphFormat** , используйте метод **SetListType** .

Значение свойства **ListType** может иметь одно из **[PbListType](pblisttype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере проверяется, если тип списка — нумерованный список, а именно **pbListTypeArabic**. Если свойство **ListType** **pbListTypeArabic**, **pbListSeparatorParenthesis**присваивается значение свойства **ListNumberSeparator** .


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeArabic Then 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 End If 
End With 
 

```


