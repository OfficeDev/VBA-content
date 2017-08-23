---
title: "Свойство Font.Tracking (издатель)"
keywords: vbapb10.chm5373984
f1_keywords: vbapb10.chm5373984
ms.prod: publisher
api_name: Publisher.Font.Tracking
ms.assetid: c703a5ec-e8d7-36ce-ac50-d41265ce92db
ms.date: 06/08/2017
ms.openlocfilehash: cf03d6f48bba4a9f2ff1c052ceae40dbec3b39f7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fonttracking-property-publisher"></a>Свойство Font.Tracking (издатель)

Возвращает или задает **Variant** , указывающее, отслеживания значение, используемое для отображения пространство между символами в диапазоне указанный текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Отслеживание**

 переменная _expression_A, представляющий объект **Font** .


## <a name="remarks"></a>Заметки

Допустимые значения — от 0.0 для 600.0 точек. Для свойства значение 0.0 отключает отслеживание. Неопределенное значения возвращаются в виде -2.


## <a name="example"></a>Пример

В этом примере отключается отслеживание во второй материал, задав свойство **отслеживания** нулевое значение.


```vb
Sub DisableTracking() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Tracking = 0.0 
 
End Sub
```


