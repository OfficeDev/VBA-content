---
title: "Метод WebNavigationBarSet.DeleteSetAndInstances (издатель)"
keywords: vbapb10.chm8519683
f1_keywords: vbapb10.chm8519683
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.DeleteSetAndInstances
ms.assetid: 89bbd9b9-d0c9-ecac-eb3e-7425bd177aec
ms.date: 06/08/2017
ms.openlocfilehash: c76f5e315056fadaa76632388eb8a6517a078880
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetdeletesetandinstances-method-publisher"></a>Метод WebNavigationBarSet.DeleteSetAndInstances (издатель)

Удаляет панель набора и все экземпляры в текущем документе навигации Web.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DeleteSetAndInstances**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


## <a name="example"></a>Пример

В следующем примере итерацию по коллекции **WebNavigationBarSets** и удаляет каждый набор из активных документов.


```vb
Dim objWebNavBarSet As WebNavigationBarSet 
For Each objWebNavBarSet In ActiveDocument.WebNavigationBarSets 
 objWebNavBarSet.DeleteSetAndInstances 
Next objWebNavBarSet
```


