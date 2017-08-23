---
title: "Свойство FindReplace.MatchDiacritics (издатель)"
keywords: vbapb10.chm8323081
f1_keywords: vbapb10.chm8323081
ms.prod: publisher
api_name: Publisher.FindReplace.MatchDiacritics
ms.assetid: e23d01a1-9252-4077-c52f-87c53b5c0589
ms.date: 06/08/2017
ms.openlocfilehash: e8f1c9c47ba874a0f3ebe96d692b6227e58994cc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacematchdiacritics-property-publisher"></a>Свойство FindReplace.MatchDiacritics (издатель)

Задает или возвращает значение **типа Boolean** представляющее ли операция поиска будет соответствовать диакритические знаки. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MatchDiacritics**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство не может быть доступно в зависимости от языка в операционной системе. Значение по умолчанию — **False**.

Возвращает **доступ запрещен** при правильного языка, например арабский, не включено.


## <a name="example"></a>Пример

В этом примере выполняется поиск первого появления слово «gegenüber» в документе Германии. 


```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "gegenüber" 
 .MatchDiacritics = True 
 .Execute 
End With 

```


