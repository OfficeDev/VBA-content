---
title: "Свойство TextRange.Font (издатель)"
keywords: vbapb10.chm5308419
f1_keywords: vbapb10.chm5308419
ms.prod: publisher
api_name: Publisher.TextRange.Font
ms.assetid: c5795f33-4e7b-f765-9ba8-f5b6706561d6
ms.date: 06/08/2017
ms.openlocfilehash: 2ed98ee671f8a726122c3943f711a24a40bdecc0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangefont-property-publisher"></a>Свойство TextRange.Font (издатель)

Задает или возвращает объект **[шрифта](font-object-publisher.md)** , представляющий атрибуты форматирования символ применяются на указанный объект. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Шрифт**

 переменная _expression_A, представляющий объект **TextRange** .


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


