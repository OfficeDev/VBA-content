---
title: "Свойство Options.AutoFormatWord (издатель)"
keywords: vbapb10.chm1048579
f1_keywords: vbapb10.chm1048579
ms.prod: publisher
api_name: Publisher.Options.AutoFormatWord
ms.assetid: b0466bd7-f0a1-44a8-480f-5d046e24e759
ms.date: 06/08/2017
ms.openlocfilehash: f716ee1656616f1c31d3a39a302d2b8a413d39de
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsautoformatword-property-publisher"></a>Свойство Options.AutoFormatWord (издатель)

 **Значение true** для Microsoft Publisher автоматически форматирование целого слова в позиции курсора даже в том случае, если текст не выделен. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutoFormatWord**

 переменная _expression_A, представляющий объект **параметров** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Если выбраны один или два символов в word, только только выбранные символы, влияют необходимые изменения, слово целиком.


## <a name="example"></a>Пример

В этом примере задается глобальных параметров для Microsoft Publisher, в том числе включения автоматического форматирования слово целиком.


```vb
Sub SetGlobalOptions() 
 With Options 
 .AutoFormatWord = True 
 .AutoKeyboardSwitching = True 
 .AutoSelectWord = True 
 .DragAndDropText = True 
 .UseCatalogAtStartup = False 
 .UseHelpfulMousePointers = False 
 End With 
End Sub
```


