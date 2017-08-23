---
title: "Объект WebNavigationBarHyperlinks (издатель)"
keywords: vbapb10.chm540671
f1_keywords: vbapb10.chm540671
ms.prod: publisher
api_name: Publisher.WebNavigationBarHyperlinks
ms.assetid: 4dfa7273-4770-d77c-275c-6b7eeae04aa5
ms.date: 06/08/2017
ms.openlocfilehash: bbe89c72d9a6b7f7ab3fffc9ef5776a0944555ab
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarhyperlinks-object-publisher"></a>Объект WebNavigationBarHyperlinks (издатель)

**WebNavigationBarHyperlinks** представляет коллекцию всех объектов **гиперссылки** на указанный объект **WebNavigationBarSet** .
 


## <a name="example"></a>Пример

Свойство **ссылки** коллекции **WebNavigationBarSets** возвращает объект **WebNavigationBarHyperlinks** . В следующем примере добавляется гиперссылки на первый **WebNavigationBarSet** активных документов.
 

 

```
Dim objWebNavLinks As WebNavigationBarHyperlinks 
Set objWebNavLinks = ActiveDocument.WebNavigationBarSets(1).Links 
objWebNavLinks.Add Address:="www.microsoft.com", _ 
 TextToDisplay:="Microsoft"
```

Используйте **WebNavigationBarHyperlinks** **. Число** для возврата Long, представляющее количество гиперссылок в коллекцию **WebNavigationBarHyperlinks** на указанный объект **WebNavigationBarSet** . Следующий пример показывает число гиперссылки в первом **WebNavigationBarSet** активных документов.
 

 



```
MsgBox ActiveDocument.WebNavigationBarSets(1).Links.Count
```

С помощью **WebNavigationBarHyperlinks**. Item(Index), где индекс — номер индекса для получения определенного объекта **гиперссылки** из коллекции. В этом примере отображается отображаемого текста первый элемент в коллекции **WebNavigationBarHyperlinks** первого **WebNavigationBarSet** активных документов.
 

 



```
MsgBox ActiveDocument.WebNavigationBarSets(1).Links.Item(1).TextToDisplay
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](webnavigationbarhyperlinks-add-method-publisher.md)|
|[Элемент](webnavigationbarhyperlinks-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](webnavigationbarhyperlinks-application-property-publisher.md)|
|[Count](webnavigationbarhyperlinks-count-property-publisher.md)|
|[Родительский раздел](webnavigationbarhyperlinks-parent-property-publisher.md)|

