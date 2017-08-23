---
title: "Свойство Application.WizardCatalogVisible (издатель)"
keywords: vbapb10.chm131173
f1_keywords: vbapb10.chm131173
ms.prod: publisher
api_name: Publisher.Application.WizardCatalogVisible
ms.assetid: 99323335-aabd-6799-b6aa-c5d95b88064f
ms.date: 06/08/2017
ms.openlocfilehash: df1b9984b7ecd86ed1774e97ab3d0416b7099644
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationwizardcatalogvisible-property-publisher"></a>Свойство Application.WizardCatalogVisible (издатель)

Возвращает или задает значение **Boolean** , указывающее, отображается ли мастер каталога. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WizardCatalogVisible**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В следующем примере сохраняется текущее состояние каталога мастера, чтобы его можно восстановить более поздней версии.


```vb
Sub WizardCatalogExample() 
 Dim blnWizardCatalog As Boolean 
 
 ' Store current state of Wizard Catalog. 
 blnWizardCatalog = Application.WizardCatalogVisible 
 
 ' Code can run here that shows or hides the Wizard 
 ' Catalog as necessary; the original setting 
 ' will be restored at the end of the procedure. 
 
 ' Restore original state of Wizard Catalog. 
 Application.WizardCatalogVisible = blnWizardCatalog 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

