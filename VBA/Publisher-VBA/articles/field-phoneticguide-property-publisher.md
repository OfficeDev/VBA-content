---
title: "Свойство Field.PhoneticGuide (издатель)"
keywords: vbapb10.chm6094856
f1_keywords: vbapb10.chm6094856
ms.prod: publisher
api_name: Publisher.Field.PhoneticGuide
ms.assetid: c68dba00-56c0-abba-0be8-5bd1a725f678
ms.date: 06/08/2017
ms.openlocfilehash: 281cc51c8f16b4fbffef32f8a39e113f793979ce
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldphoneticguide-property-publisher"></a>Свойство Field.PhoneticGuide (издатель)

Возвращает объект **PhoneticGuide** , который представляет свойства фонетическое текста, примененные к диапазону текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PhoneticGuide**

 переменная _expression_A, представляющий объект **поля** .


### <a name="return-value"></a>Возвращаемое значение

PhoneticGuide


## <a name="example"></a>Пример

В этом примере добавляется фонетическое текст для выделения и отображает текст, к которому применяется фонетическое текста, которая является изначально выбранный текст. В этом примере предполагается, что выделенный текст. Если текст не выделен, в окне сообщения будет пустым.


```vb
Sub AddPhoneticText() 
 With Selection.TextRange.Fields.AddPhoneticGuide _ 
 (Range:=Selection.TextRange, Text:="ver-E nIs") 
 MsgBox "The base text is " &; .PhoneticGuide.BaseText 
 End With 
End Sub
```


