---
title: "Свойство Options.AutoKeyboardSwitching (издатель)"
keywords: vbapb10.chm1048627
f1_keywords: vbapb10.chm1048627
ms.prod: publisher
api_name: Publisher.Options.AutoKeyboardSwitching
ms.assetid: 05f22aa6-332d-e033-ab9d-550eb08f1018
ms.date: 06/08/2017
ms.openlocfilehash: 04d4bcf51ed4799b257a166945719d48b6c28dc5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsautokeyboardswitching-property-publisher"></a>Свойство Options.AutoKeyboardSwitching (издатель)

 **Значение true** для Microsoft Publisher для автоматическое переключение языка клавиатуры для языка, используемого для текста в позиции курсора. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutoKeyboardSwitching**

 переменная _expression_A, представляющий объект **параметров** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере включается автоматическое переключение языка клавиатуры в необходимых языков.


```vb
Sub SetGlobalOptions() 
 Options.AutoKeyboardSwitching = True 
End Sub
```


