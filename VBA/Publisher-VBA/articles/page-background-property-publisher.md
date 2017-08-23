---
title: "Свойство Page.Background (издатель)"
keywords: vbapb10.chm393249
f1_keywords: vbapb10.chm393249
ms.prod: publisher
api_name: Publisher.Page.Background
ms.assetid: 1bba32dc-0e7e-40ca-0f29-b67be6be518d
ms.date: 06/08/2017
ms.openlocfilehash: d7d4f96732705795dd2f1452b6d7f85154f0642e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagebackground-property-publisher"></a>Свойство Page.Background (издатель)

Задает или возвращает объект **PageBackground** , представляющий фон для указанной страницы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Фон**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

PageBackground


## <a name="remarks"></a>Заметки

Это свойство соответствует публикации только для страниц. Любая попытка создания фон для главной страницы возвратит ошибку «Отказано в разрешении».


## <a name="example"></a>Пример

В следующем примере создается объект **PageBackground** и задает фон первой страницы активных документов.


```vb
Dim objPageBackground As PageBackground 
Set objPageBackground = ActiveDocument.Pages(1).Background 
 
```


