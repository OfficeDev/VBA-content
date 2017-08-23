---
title: "Свойство FindReplace.ReplaceWithText (издатель)"
keywords: vbapb10.chm8323077
f1_keywords: vbapb10.chm8323077
ms.prod: publisher
api_name: Publisher.FindReplace.ReplaceWithText
ms.assetid: 7bd0457f-c55e-3350-fe16-b9eac7d7d4fa
ms.date: 06/08/2017
ms.openlocfilehash: 33143cac5e1fb56196c4fc5aaa229cbef257b646
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacereplacewithtext-property-publisher"></a>Свойство FindReplace.ReplaceWithText (издатель)

Задает или получает **строку** , представляющую текст замены в указанный диапазон или выделить фрагмент. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ReplaceWithText**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

По умолчанию свойства **ReplaceWithText** — это пустая **строка**.

Если свойство **ReplaceScope** имеет значение **pbReplaceScopeOne** или **pbReplaceScopeAll** и **ReplaceWithText** свойство не задано, это текст будет заменен пустая строка по умолчанию, что приведет к удалению текст.


## <a name="example"></a>Пример

Следующий пример заменяет все вхождения слова «hello» с «goodbye» в активном документе.


```vb
With ActiveDocument.Find 
 .Clear 
 .FindText = "hello" 
 .ReplaceWithText = "goodbye" 
 .MatchWholeWord = True 
 .ReplaceScope = pbReplaceScopeAll 
 .Execute 
End With
```


