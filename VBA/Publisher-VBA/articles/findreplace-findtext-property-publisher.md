---
title: "Свойство FindReplace.FindText (издатель)"
keywords: vbapb10.chm8323076
f1_keywords: vbapb10.chm8323076
ms.prod: publisher
api_name: Publisher.FindReplace.FindText
ms.assetid: 5c8d2803-174e-a82f-d94c-3d96c4b4a2eb
ms.date: 06/08/2017
ms.openlocfilehash: 85234d4390051ac5cefc5066c429edd85798778f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacefindtext-property-publisher"></a>Свойство FindReplace.FindText (издатель)

Задает или получает **строку** представляющий текст для поиска в указанном диапазоне или выбора. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FindText**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Свойство **FindText** возвращает обычная, неформатированный текст текущего выбора. Если этому свойству присвоено, текст поиска указан. Можно выполнить поиск специальных символов, указав коды знаков. Например «^ p» соответствует знак абзаца и «^ t» соответствует символ табуляции.

Значение по умолчанию для свойства **FindText** представляет собой пустую строку. Поддерживается поиск только текст, **FindText** необходимо явно задать во избежание ошибок времени выполнения.


## <a name="example"></a>Пример

В этом примере заменяет все вхождения слова «» в выделении «,» в каждой открытой публикации.


```vb
Dim objDocument As Document 
 
For Each objDocument In Documents 
 With objDocument.Find 
 .Clear 
 .MatchCase = True 
 .FindText = "This" 
 .ReplaceWithText = "That" 
 .ReplaceScope = pbReplaceScopeAll 
 .Forward = True 
 .Execute 
 End With 
Next objDocument 

```


