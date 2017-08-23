---
title: "Метод Application.ShowWizardCatalog (издатель)"
keywords: vbapb10.chm131189
f1_keywords: vbapb10.chm131189
ms.prod: publisher
api_name: Publisher.Application.ShowWizardCatalog
ms.assetid: a8307ff9-a6c1-7655-8127-284f3781dae9
ms.date: 06/08/2017
ms.openlocfilehash: 083d13ade3ea72eb91032e80bcced64e58c4858a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationshowwizardcatalog-method-publisher"></a>Метод Application.ShowWizardCatalog (издатель)

Отображает каталога **Публикаций** для мастера для указанного типа.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShowWizardCatalog** ( **_Мастер_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Мастер|Необязательный| **PbWizard**|Тип каталога мастера для отображения.|

## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **ShowWizardCatalog** для отображения каталога мастера для брошюры.


```vb
Public Sub ShowWizardCatalog_Example() 
 Application.ShowWizardCatalog (pbWizardBrochures) 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

