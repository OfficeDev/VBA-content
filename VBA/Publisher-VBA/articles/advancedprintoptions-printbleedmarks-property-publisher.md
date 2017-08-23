---
title: "Свойство AdvancedPrintOptions.PrintBleedMarks (издатель)"
keywords: vbapb10.chm7077907
f1_keywords: vbapb10.chm7077907
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.PrintBleedMarks
ms.assetid: f0c69d5f-4bfd-7a4c-3607-714859bcc86c
ms.date: 06/08/2017
ms.openlocfilehash: 15a219e2bca977b63e6e0309a7df5f892b538b9f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsprintbleedmarks-property-publisher"></a>Свойство AdvancedPrintOptions.PrintBleedMarks (издатель)

 **Значение true** для печати в указанной публикации. Значение по умолчанию — **False**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PrintBleedMarks**

 переменная _expression_A, представляет собой объект- **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Метки край степени край и печать восьмому дюйма за пределами метки обрезки.

Это свойство доступно только в том случае, если край разрешены в указанной публикации. Используйте свойство **[AllowBleeds](advancedprintoptions-allowbleeds-property-publisher.md)** объекта **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** для указания край разрешены. Возвращает «Отказано в разрешении», если край не допускается в публикации.

Это свойство соответствует управления **метки выхода за край** на вкладке **Параметры страницы** диалоговое окно **Дополнительные параметры печати** .


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

