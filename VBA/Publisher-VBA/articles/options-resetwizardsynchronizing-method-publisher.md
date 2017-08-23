---
title: "Метод Options.ResetWizardSynchronizing (издатель)"
keywords: vbapb10.chm1048617
f1_keywords: vbapb10.chm1048617
ms.prod: publisher
api_name: Publisher.Options.ResetWizardSynchronizing
ms.assetid: 1027a113-45aa-b722-b625-a6bb7bbcc3e6
ms.date: 06/08/2017
ms.openlocfilehash: e12b2feafbc59a2703128dbda5cca26a0648e7d9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsresetwizardsynchronizing-method-publisher"></a>Метод Options.ResetWizardSynchronizing (издатель)

Восстанавливает данные, которые использует Microsoft Publisher для автоматического изменения подобных объектов же форматирование или содержимого.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ResetWizardSynchronizing**

 переменная _expression_A, представляющий объект **параметров** .


## <a name="remarks"></a>Заметки

Непредвиденные изменения форматирования может быть в результате синхронизации объектов издателя. Сброс синхронизации данных остановит эти изменения.


## <a name="example"></a>Пример

В следующем примере сбрасывается данных синхронизации, Publisher используется для предоставления подобных объектов и то же форматирование.


```
Options.ResetWizardSynchronizing
```


