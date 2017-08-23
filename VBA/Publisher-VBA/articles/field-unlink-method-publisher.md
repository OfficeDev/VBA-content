---
title: "Метод Field.Unlink (издатель)"
keywords: vbapb10.chm6094857
f1_keywords: vbapb10.chm6094857
ms.prod: publisher
api_name: Publisher.Field.Unlink
ms.assetid: 4dfe5c29-eb1e-b071-fd86-6ee222455c4e
ms.date: 06/08/2017
ms.openlocfilehash: bf611346d682ce9ffe6f1df93b5b54e130ba62c4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldunlink-method-publisher"></a>Метод Field.Unlink (издатель)

Заменяет указанное поле или коллекции **[полей](fields-object-publisher.md)** с их последние результаты.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Разорвать связь**

 переменная _expression_A, представляющий объект **поля** .


## <a name="remarks"></a>Заметки

Когда вы разорвать связь с полем, его текущего результат преобразуется в текст или графику и уже не могут быть обновлены автоматически.


## <a name="example"></a>Пример

В этом примере удаляется связь первого поля в форме одну на первой странице active публикации.


```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Fields(1).Unlink
```


