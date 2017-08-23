---
title: "Объект HeaderFooter (издатель)"
keywords: vbapb10.chm7536639
f1_keywords: vbapb10.chm7536639
ms.prod: publisher
api_name: Publisher.HeaderFooter
ms.assetid: d38e5e7e-45d7-667b-b6f2-9ad8e764af79
ms.date: 06/08/2017
ms.openlocfilehash: 704376967ca0cf9e1ecacf9cfe6dbda99801cce5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="headerfooter-object-publisher"></a>Объект HeaderFooter (издатель)

Представляет верхних и нижних колонтитулов главной страницы.
 


## <a name="example"></a>Пример

Используйте **MasterPages.Header** или **MasterPages.Footer** для возврата объекта **HeaderFooter** . Следующий пример добавляет текст заголовка первого главной страницы активных документов.
 

 

```
Dim objHeader As HeaderFooter 
Set objHeader = ActiveDocument.MasterPages(1).Header 
objHeader.TextRange.Text = "Master Page 1 Header" 

```

Используйте **HeaderFooter.Delete** для удаления какой-либо контент из верхнего или нижнего колонтитула. Вызывающий этот метод не удаляет рамки, только содержимое его. В следующем примере удаляется весь контент верхний и нижний колонтитулы главных страниц в публикации.
 

 



```
Dim objMasterPage As page 
For Each objMasterPage In ActiveDocument.masterPages 
 objMasterPage.Header.Delete 
 objMasterPage.Footer.Delete 
Next
```

Используйте **HeaderFooter.TextRange** , чтобы получить объект **TextRange** , представляющий верхних и нижних колонтитулов главной страницы. Обработка содержимого любого верхних и нижних колонтитулов выполняется с помощью на это свойство объекта **HeaderFooter** . В следующем примере сначала удаляет какой-либо контент и затем добавляет некоторые стандартный текст заголовка главной страницы.
 

 



```
Dim objHeader As HeaderFooter 
Set objHeader = ActiveDocument.MasterPages(1).Header 
With objHeader 
 .Delete 
 .TextRange.Text = "<Insert Address Here>" 
End With
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](headerfooter-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](headerfooter-application-property-publisher.md)|
|[IsHeader](headerfooter-isheader-property-publisher.md)|
|[Родительский раздел](headerfooter-parent-property-publisher.md)|
|[TextRange](headerfooter-textrange-property-publisher.md)|

