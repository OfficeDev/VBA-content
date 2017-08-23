---
title: "Метод WebNavigationBarSets.AddSet (издатель)"
keywords: vbapb10.chm8454148
f1_keywords: vbapb10.chm8454148
ms.prod: publisher
api_name: Publisher.WebNavigationBarSets.AddSet
ms.assetid: 5b998e14-b1eb-2a4a-2ed5-9a1ef16d69c1
ms.date: 06/08/2017
ms.openlocfilehash: 4be8f00d2030206821c241dc6a2dd0c3e9d27f45
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetsaddset-method-publisher"></a>Метод WebNavigationBarSets.AddSet (издатель)

Добавляет новый объект **WebNavigationBarSet** , представляющий панель навигации, задайте значение указанной коллекции **WebNavigationBarSets** . .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddSet** ( **_Имя_**, **_разработки_**, **_Автоматическое обновление_**)

 переменная _expression_A, представляет собой объект- **WebNavigationBarSets** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|Обязательное свойство.| **String**|Имя панель навигации для добавления. Этот параметр должен быть уникальным.|
|Разработка|Необязательный| **PbWizardNavBarDesign**|Задает схему конструктора панели навигации.|
|Автоматическое обновление|Необязательный| **Boolean**| **Значение true,** Если все страницы с помощью свойства **AddHyperlinkToWebNavBar** задано значение **True,**добавляются в качестве ссылок на панель навигации и хранения обновленные на панели навигации.|

### <a name="return-value"></a>Возвращаемое значение

WebNavigationBarSet


## <a name="remarks"></a>Заметки

Параметр **Name** должно быть уникальным во избежание ошибок времени выполнения.


## <a name="example"></a>Пример

В следующем примере добавляет **WebNavigationBarSet** объект в коллекцию **WebNavigationBarSets** активного документа, затем задает некоторые свойства.


```vb
Dim objWebNavBarSet As WebNavigationBarSet 
 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.AddSet( _ 
 Name:="WebNavBarSet1", _ 
 Design:=pbnbDesignAmbient, _ 
 AutoUpdate:=True) 
 
With objWebNavBarSet 
 .AddToEveryPage Left:=50, Top:=10 
 .ButtonStyle = pbnbDesignTopLine 
 .ChangeOrientation pbNavBarOrientHorizontal 
End With
```


