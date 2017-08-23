---
title: "Свойство Application.Build (издатель)"
keywords: vbapb10.chm131078
f1_keywords: vbapb10.chm131078
ms.prod: publisher
api_name: Publisher.Application.Build
ms.assetid: e0d4bb8e-5185-3d3c-fd80-c1e3c3902b2c
ms.date: 06/08/2017
ms.openlocfilehash: b66b7b9a32949c6e9380f68c5986864d2c6fea90
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationbuild-property-publisher"></a>Свойство Application.Build (издатель)

Возвращает **Срок** , представляющий Microsoft Publisher номер сборки. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Построение**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере отображается номер сборки Publisher.


```vb
Sub BuildNumber() 
 MsgBox Prompt:="The current Microsoft Publisher build number is " &; _ 
 Application.Build, Title:="Microsoft Publisher Build" 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

