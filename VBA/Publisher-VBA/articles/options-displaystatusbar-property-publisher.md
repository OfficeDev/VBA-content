---
title: "Свойство Options.DisplayStatusBar (издатель)"
keywords: vbapb10.chm1048583
f1_keywords: vbapb10.chm1048583
ms.prod: publisher
api_name: Publisher.Options.DisplayStatusBar
ms.assetid: 335b2f1e-03ff-fd90-5ec2-27d5219b27e7
ms.date: 06/08/2017
ms.openlocfilehash: def1d2401a0845aebb103d4280995eb30dc9e155
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsdisplaystatusbar-property-publisher"></a>Свойство Options.DisplayStatusBar (издатель)

 **Значение true** для Microsoft Publisher для отображения в строке состояния в нижней части окна Publisher. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DisplayStatusBar**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере скрывается в строке состояния из представления.


```vb
Sub HideStatusBar() 
 Options.DisplayStatusBar = False 
End Sub
```


