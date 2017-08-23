---
title: "Свойство TextRange.Start (издатель)"
keywords: vbapb10.chm5308433
f1_keywords: vbapb10.chm5308433
ms.prod: publisher
api_name: Publisher.TextRange.Start
ms.assetid: 40604058-7c3e-b4c7-c793-bbf09091b4c1
ms.date: 06/08/2017
ms.openlocfilehash: 68578c4c32fc27aec52191ab2f126eca689b0d7b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangestart-property-publisher"></a>Свойство TextRange.Start (издатель)

Возвращает или задает **времени** , представляющий позиция первого знака диапазон текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Запуск**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Если это свойство задано значение больше **End** property, то же значение, что и свойство **при запуске** задано **End** property.


## <a name="example"></a>Пример

В этом примере выполняются первых 15 символов диапазона выделенный текст полужирным шрифтом. В этом примере предполагается, что в активной публикации выбранного текста.


```vb
Sub SetSelectionRange() 
 With Selection 
 With .TextRange 
 .Start = 0 
 .End = 15 
 .Font.Bold = msoTrue 
 End With 
 End With 
End Sub
```


