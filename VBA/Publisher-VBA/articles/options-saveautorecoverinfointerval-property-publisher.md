---
title: "Свойство Options.SaveAutoRecoverInfoInterval (издатель)"
keywords: vbapb10.chm1048600
f1_keywords: vbapb10.chm1048600
ms.prod: publisher
api_name: Publisher.Options.SaveAutoRecoverInfoInterval
ms.assetid: 3d6a6c4f-7e2b-18ff-67a4-20dee4fbcf5b
ms.date: 06/08/2017
ms.openlocfilehash: be90e967d06e935dc82c955b4ab44797717ddd4f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionssaveautorecoverinfointerval-property-publisher"></a>Свойство Options.SaveAutoRecoverInfoInterval (издатель)

Возвращает или задает **времени** , представляющий временной интервал в минутах для автоматического сохранения публикации для восстановления, если приложение неожиданно завершить работу. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SaveAutoRecoverInfoInterval**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

Этот пример включает параметр глобального автоматическое восстановление и задает сохранения интервал для каждые пять минут.


```vb
Sub SetAutoRecoverInfo() 
 With Options 
 .SaveAutoRecoverInfo = True 
 .SaveAutoRecoverInfoInterval = 5 
 End With 
End Sub
```


