---
title: "Свойство AdvancedPrintOptions.AllowBleeds (издатель)"
keywords: vbapb10.chm7077906
f1_keywords: vbapb10.chm7077906
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.AllowBleeds
ms.assetid: 0c12a611-4e1e-468b-ada2-f07d01fd4445
ms.date: 06/08/2017
ms.openlocfilehash: 878b5dc33fdda1137b537c953e526d8903b141f7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsallowbleeds-property-publisher"></a>Свойство AdvancedPrintOptions.AllowBleeds (издатель)

 **Значение true,** чтобы разрешить край печати для указанной публикации. По умолчанию используется **значение True**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AllowBleeds**

 переменная _expression_A, представляющий объект **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Если предоставлено край эти объекты частично странице печать в одной восьмой дюйма за пределами определенного размера страницы.

Если вы включили край в документе, можно указать, ли метки край печатаются, используя свойство **[PrintBleedMarks](advancedprintoptions-printbleedmarks-property-publisher.md)** объекта **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** .

Это свойство соответствует **Разрешить выход за край** элемента управления на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .


## <a name="example"></a>Пример

В следующем примере задается публикации на разрешение край и печати.


```vb
Sub AllowBleedsAndPrintMarks() 
 With ActiveDocument.AdvancedPrintOptions 
 .AllowBleeds = True 
 .PrintBleedMarks = True 
 End With 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект AdvancedPrintOptions](advancedprintoptions-object-publisher.md)

