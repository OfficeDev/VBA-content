---
title: "Свойство FindReplace.ReplaceScope (издатель)"
keywords: vbapb10.chm8323085
f1_keywords: vbapb10.chm8323085
ms.prod: publisher
api_name: Publisher.FindReplace.ReplaceScope
ms.assetid: 555fe65b-9edb-8888-03f0-15ce34813d5f
ms.date: 06/08/2017
ms.openlocfilehash: 02cef638162c92d276af738cd01e63685fbabe23
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacereplacescope-property-publisher"></a>Свойство FindReplace.ReplaceScope (издатель)

Указывает, сколько замены область следует обратить внимание: один, все или нет. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ReplaceScope**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

PbReplaceScope


## <a name="remarks"></a>Заметки

Значение свойства **ReplaceScope** может иметь одно из **[PbReplaceScope](pbreplacescope-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Значение по умолчанию свойства **ReplaceScope** — **pbReplaceScopeNone**.


## <a name="example"></a>Пример

Следующий пример заменяет все вхождения слова «Привет» «hello» в активный документ.


```vb
With ActiveDocument.Find 
 .Clear 
 .FindText = "hi" 
 .ReplaceWithText = "hello" 
 .MatchWholeWord = True 
 .ReplaceScope = pbReplaceScopeAll 
 .Execute 
End With
```


