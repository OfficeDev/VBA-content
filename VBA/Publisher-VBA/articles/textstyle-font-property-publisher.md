---
title: "Свойство TextStyle.Font (издатель)"
keywords: vbapb10.chm5963780
f1_keywords: vbapb10.chm5963780
ms.prod: publisher
api_name: Publisher.TextStyle.Font
ms.assetid: 80d7177a-fef9-c3fd-f559-94644a2ba0f7
ms.date: 06/08/2017
ms.openlocfilehash: 223b572b42edbfd5acd4ce940fc92187b398bbf7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstylefont-property-publisher"></a>Свойство TextStyle.Font (издатель)

Задает или возвращает объект **[шрифта](font-object-publisher.md)** , представляющий атрибуты форматирования символ применяются на указанный объект. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Шрифт**

 переменная _expression_A, представляющий объект **стиля текста** .


## <a name="example"></a>Пример

В этом примере выбирает текст и форматирование шрифта как полужирным шрифтом.


```vb
Sub test2() 
 With Selection.TextRange 
 .Start = 50 
 .End = 150 
 .Font.Bold = msoTrue 
 End With 
End Sub
```


