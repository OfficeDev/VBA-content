---
title: "Возвращение объекта из коллекции (издатель)"
ms.prod: publisher
ms.assetid: 08b8c469-f4f1-8717-a767-ab57c792606b
ms.date: 06/08/2017
ms.openlocfilehash: 34b8ea5957dad93bcc7296e3777133a4fc0a152d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="returning-an-object-from-a-collection-publisher"></a>Возвращение объекта из коллекции (издатель)

Метод **Item** возвращает объект из коллекции. В следующем примере задается переменная объекту **[страницы](page-object-publisher.md)** , который представляет первой страницы в коллекции **[страниц](pages-object-publisher.md)** .


```vb
Sub SetFirstPage() 
 Dim pgFirst As Page 
 Set pgFirst = ActiveDocument.Pages.Item(1) 
End Sub
```


Метод **Item** — это метод по умолчанию для большинства семейств сайтов, так же инструкции можно написать более кратко, не указывайте ключевое слово **элемента** .




```vb
Sub SetFirstPage() 
 Dim pgFirst As Page 
 Set pgFirst = ActiveDocument.Pages(1) 
End Sub
```


