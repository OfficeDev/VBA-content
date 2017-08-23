---
title: "Свойство Options.SaveAutoRecoverInfo (издатель)"
keywords: vbapb10.chm1048599
f1_keywords: vbapb10.chm1048599
ms.prod: publisher
api_name: Publisher.Options.SaveAutoRecoverInfo
ms.assetid: 1cbb7960-8995-37f4-5989-01b97152269f
ms.date: 06/08/2017
ms.openlocfilehash: c1ea2b01d23da1d906fbb58cffb02f1a0230c453
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionssaveautorecoverinfo-property-publisher"></a>Свойство Options.SaveAutoRecoverInfo (издатель)

 **Значение true,** Если Microsoft Publisher автоматически сохраняет публикации для восстановления, если приложение неожиданно завершить работу. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SaveAutoRecoverInfo**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Свойство **[SaveAutoRecoverInfoInterval](options-saveautorecoverinfointerval-property-publisher.md)** определяет, как часто возникают сохраняет автоматическое восстановление.


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


