---
title: "Метод MailMergeDataSource.Close (издатель)"
keywords: vbapb10.chm6291493
f1_keywords: vbapb10.chm6291493
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.Close
ms.assetid: c215743b-590a-6db9-e902-b9179b67bb8e
ms.date: 06/08/2017
ms.openlocfilehash: d7aeee15595d833dbe0381590ca662d64fa56298
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceclose-method-publisher"></a>Метод MailMergeDataSource.Close (издатель)

Закрывает указанного источника данных, отменяет слияния почты и преобразует все поля данных слияния почты в обычный текст.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Закрытие**

 переменная _expression_A, представляющий объект **вывода** .


## <a name="remarks"></a>Заметки

Закрытие источника данных для слияния удаляет фигуры, представляющий область данных страницы публикации, связанный с источником данных.


## <a name="example"></a>Пример

В следующем примере закрывается источник данных для публикации active слияния.


```vb
ActiveDocument.MailMerge.DataSource.Close
```


