---
title: "Свойство Options.SequenceCheck (издатель)"
keywords: vbapb10.chm1048625
f1_keywords: vbapb10.chm1048625
ms.prod: publisher
api_name: Publisher.Options.SequenceCheck
ms.assetid: a2801af8-5c89-9256-80a6-d9dac17b6066
ms.date: 06/08/2017
ms.openlocfilehash: 376f3b3bb18827f1b2a18da35cbac2beb6d51cd0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionssequencecheck-property-publisher"></a>Свойство Options.SequenceCheck (издатель)

 **Значение true,** для проверки последовательности independent символов на восточноазиатских языках. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SequenceCheck**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере включается проверка последовательности, позволяя у пользователя ввод допустимую последовательность независимой знаков недействительный символ ячеек в южно-азиатских текста.


```vb
Sub CheckSequence() 
 Options.SequenceCheck = True 
End Sub
```


