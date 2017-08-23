---
title: "Свойство FindReplace.Forward (издатель)"
keywords: vbapb10.chm8323078
f1_keywords: vbapb10.chm8323078
ms.prod: publisher
api_name: Publisher.FindReplace.Forward
ms.assetid: a1a0046c-81be-62d6-8739-5dc843d249bc
ms.date: 06/08/2017
ms.openlocfilehash: 544b76f291f8b02e5933b32a1f4ce04fce3e3816
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplaceforward-property-publisher"></a>Свойство FindReplace.Forward (издатель)

Задает или получает **логическое** представляющее направление поиска текста. **Значение true,** Если операция поиска осуществляет поиск вперед в документе. **Значение false,** если оно выполняет поиск в документе. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вперед**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Прямого должно быть присвоено **значение True,** Если замена текста.


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


