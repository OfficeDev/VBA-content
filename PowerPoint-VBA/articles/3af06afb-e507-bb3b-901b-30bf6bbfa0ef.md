
# Comment.Replies Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Property value](#sectionSection2)


Returns a  [Comments](1f29db7c-90fa-db9f-5229-136534ce803d.md) collection of **Comment** objects that are children of the specified comment. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Replies**

 _expression_A variable that represents a  **Comment** object.


## Remarks
<a name="sectionSection1"> </a>

Calling the  [Add](ab520c51-2a8b-2e37-2e4c-8fce7a70a5ab.md) method on the returned collection of replies adds a new reply, unless the collection was accessed from a reply to a reply.


## Property value
<a name="sectionSection2"> </a>

 **COMMENTS**

