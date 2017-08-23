---
title: "Свойство Application.ProductCode (издатель)"
keywords: vbapb10.chm131105
f1_keywords: vbapb10.chm131105
ms.prod: publisher
api_name: Publisher.Application.ProductCode
ms.assetid: aacd5ff6-dad1-af86-f4e0-af9012ae93f8
ms.date: 06/08/2017
ms.openlocfilehash: b0e141f6131427b5569ba7c8d000ba76daa51b01
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationproductcode-property-publisher"></a>Свойство Application.ProductCode (издатель)

Возвращает **строку** , указывающую Microsoft Publisher глобальный уникальный идентификатор (GUID). Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Код продукта**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В следующем примере отображается код продукта для Publisher.


```vb
MsgBox "The product code for Microsoft Publisher is " _ 
 &; ProductCode
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

