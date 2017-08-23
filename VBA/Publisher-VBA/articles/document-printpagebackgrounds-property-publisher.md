---
title: "Свойство Document.PrintPageBackgrounds (издатель)"
keywords: vbapb10.chm196743
f1_keywords: vbapb10.chm196743
ms.prod: publisher
api_name: Publisher.Document.PrintPageBackgrounds
ms.assetid: 6d1d6e6a-fd66-2afa-2172-4a6552d5cce4
ms.date: 06/08/2017
ms.openlocfilehash: 13abd401a423e24e7c2f15bf7a99d0d0d07f1405
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentprintpagebackgrounds-property-publisher"></a>Свойство Document.PrintPageBackgrounds (издатель)

Возвращает или задает **значение True** для включения фона страницы при печати страниц с указанной публикации. Значение по умолчанию — **True**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PrintPageBackgrounds**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Использование объекта **[PageBackground](pagebackground-object-publisher.md)** для создания, изменения или удаления фона указанной странице.


## <a name="example"></a>Пример

В следующем примере задается фона страницы для печати для активной публикации.


```vb
ActiveDocument.PrintPageBackgrounds = True
```


