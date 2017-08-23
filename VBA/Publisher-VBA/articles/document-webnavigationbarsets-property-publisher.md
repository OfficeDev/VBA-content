---
title: "Свойство Document.WebNavigationBarSets (издатель)"
keywords: vbapb10.chm196741
f1_keywords: vbapb10.chm196741
ms.prod: publisher
api_name: Publisher.Document.WebNavigationBarSets
ms.assetid: 4193dbce-a2e3-2587-5282-43b4c3cec921
ms.date: 06/08/2017
ms.openlocfilehash: 623a8a29b9f33692c2bf8da8b3b30d20eea97250
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentwebnavigationbarsets-property-publisher"></a>Свойство Document.WebNavigationBarSets (издатель)

Возвращает объект **WebNavigationBarSets** , представляющий коллекцию объектов все **WebNavigationBarSet** в указанный документ. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebNavigationBarSets**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

WebNavigationBarSets


## <a name="example"></a>Пример

В следующем примере задается объектную переменную в коллекцию панель навигации задается в активном документе и добавляет новую панель навигации значение его.


```vb
Dim objWebNavBarSets As WebNavigationBarSets 
 
Set objWebNavBarSets = ActiveDocument.WebNavigationBarSets 
objWebNavBarSets.AddSet _ 
 Name:="WebNavBarSet1", _ 
 Design:=pbnbDesignBracket, _ 
 AutoUpdate:=True 

```


