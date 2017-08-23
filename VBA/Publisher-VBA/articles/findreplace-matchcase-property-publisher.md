---
title: "Свойство FindReplace.MatchCase (издатель)"
keywords: vbapb10.chm8323080
f1_keywords: vbapb10.chm8323080
ms.prod: publisher
api_name: Publisher.FindReplace.MatchCase
ms.assetid: 4fabf2f8-f1e4-bc70-e8e6-96dd09cd23d8
ms.date: 06/08/2017
ms.openlocfilehash: bb7f3459e281ce67468eddd03713cbebc2ab3d3e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="findreplacematchcase-property-publisher"></a>Свойство FindReplace.MatchCase (издатель)

Задает или возвращает значение **типа Boolean** , представляющий регистра операции поиска. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MatchCase**

 переменная _expression_A, представляет собой объект- **FindReplace** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Значение по умолчанию для **MatchCase** присвоено **значение False**.


## <a name="example"></a>Пример

В этом примере будет выбран первое слово «фабрики» вне зависимости от случая.


```vb
With ActiveDocument.Find 
 .Clear 
 .MatchCase = False 
 .FindText = "factory" 
 .Execute 
End With 

```


