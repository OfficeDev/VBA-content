---
title: "Свойство Options.ShowScreenTipsOnObjects (издатель)"
keywords: vbapb10.chm1048608
f1_keywords: vbapb10.chm1048608
ms.prod: publisher
api_name: Publisher.Options.ShowScreenTipsOnObjects
ms.assetid: b5503200-31fd-72ac-de28-ace55a7123b3
ms.date: 06/08/2017
ms.openlocfilehash: c0512115bca23bcf5bc5aa2e30683853e91930eb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsshowscreentipsonobjects-property-publisher"></a>Свойство Options.ShowScreenTipsOnObjects (издатель)

 **Значение true** для Microsoft Publisher отобразить всплывающие подсказки при наведении указателя мыши на текстовое поле, фигуру или другой объект. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShowScreenTipsOnObjects**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере отключается отображение всплывающих подсказок на объекты.


```vb
Sub DisableScreenTips() 
 Options.ShowScreenTipsOnObjects = False 
End Sub
```


