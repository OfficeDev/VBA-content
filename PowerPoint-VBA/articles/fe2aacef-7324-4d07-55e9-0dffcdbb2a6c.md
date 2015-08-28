
# SlideShowView Members (PowerPoint)
Represents the view in a slide show window.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [DrawLine](d4c3c1c9-cd12-67ba-b1b9-4d7e924bd084.md)|Draws a line in the specified slide show view.|
| [EndNamedShow](1b829558-a729-8aa1-c260-8b7410501153.md)|Switches from running a custom, or named, slide show to running the entire presentation of which the custom show is a subset. When the slide show advances from the current slide, the next slide displayed will be the next one in the entire presentation, not the next one in the custom slide show.|
| [EraseDrawing](d1ccb77b-c591-f3ec-bb88-1f317f057103.md)|Removes lines drawn during a slide show by using either the  ** [DrawLine](d4c3c1c9-cd12-67ba-b1b9-4d7e924bd084.md)**method or the pen tool.|
| [Exit](9abcb628-395b-02bf-3a61-d0c7b8429741.md)|Ends the specified slide show.|
| [First](5f360832-2deb-b3df-7b55-5a3c964d0057.md)|Sets the specified slide show view to display the first slide in the presentation.|
| [FirstAnimationIsAutomatic](689b2dfc-a441-51c6-9eea-de99194ba203.md)|Returns  **True** if the current slide has an initial animation that runs automically.|
| [GetClickCount](3df28d31-4da1-1ea3-e1d6-5ff334018ebc.md)|Returns the number of mouse clicks that are defined for a slide.|
| [GetClickIndex](678feca3-79d4-e4e8-83aa-3484f5c099e9.md)|Returns the index number of the current mouse click for an animation that is actively playing on a slide or has just finished.|
| [GotoClick](b41dec86-96a9-447a-5895-0b28fc4bd6b2.md)|Plays an animation associated with a specified mouse click and any animations that follow on the slide.|
| [GotoNamedShow](7e26b77f-bb7b-fd32-eabf-bc8f568e5c62.md)|Switches to the specified custom, or named, slide show during another slide show. When the slide show advances from the current slide, the next slide displayed will be the next one in the specified custom slide show, not the next one in current slide show.|
| [GotoSlide](f733f46d-a632-02cb-3dbf-f29122fe347a.md)|Switches to the specified slide during a slide show. You can specify whether you want the animation effects to be rerun.|
| [Last](1188d75f-9561-b92c-e2d1-9ceb03eae904.md)|Sets the specified slide show view to display the last slide in the presentation.|
| [Next](cf95eef7-4fd7-4c47-4436-037ec1882d4c.md)|Displays the slide immediately following the slide that's currently displayed. |
| [Player](d7bb6b02-516b-07bb-42b4-ae245ce20262.md)|Allows access to playback controls for the associated view in the current window.|
| [Previous](a53741b0-8325-696c-51e5-ffd3f9358ca8.md)|Shows the slide immediately preceding the slide that's currently displayed. |
| [ResetSlideTime](aa00c585-d3c3-9cdc-860d-8c1f2f0a6ef3.md)|Resets the elapsed time (represented by the  ** [SlideElapsedTime](e9250ea3-c37e-ebed-c8a8-9774dab77f37.md)** property) for the slide that's currently displayed to 0 (zero).|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [AcceleratorsEnabled](04db702f-af30-1868-0cab-17e692892e82.md)|Determines whether shortcut keys are enabled during a slide show. Read/write.|
| [AdvanceMode](cdc2a780-c591-b96d-cc2e-7b0571056491.md)|Returns a value that indicates how the slide show in the specified view advances. Read-only.|
| [Application](bdfbaf89-cd91-2a3a-481c-346c11b889e7.md)|Returns an  ** [Application](978c2b99-4271-b953-4283-73b5f3d96f41.md)**object that represents the creator of the specified object.|
| [CurrentShowPosition](390eb2c3-059f-f7e9-e91a-0e8cf9a0ddff.md)|Returns the position of the current slide within the slide show that is showing in the specified view. Read-only.|
| [IsNamedShow](a68632b2-bff4-9047-f0b8-6acb22a29071.md)|Determines whether a custom (named) slide show is displayed in the specified slide show view. Read-only.|
| [LastSlideViewed](47647e03-d898-47b5-cb50-79f3e368b56f.md)|Returns a  ** [Slide](afe42344-6898-00d2-ecc1-b0ed23a71fe8.md)** object that represents the slide viewed immediately before the current slide in the specified slide show view.|
| [MediaControlsHeight](523732d6-6b6a-7658-a8f0-dbdeb9e3e68e.md)|Returns the height of the media control bounding box. Read-only.|
| [MediaControlsLeft](1cc3c3a2-63d8-e43b-2056-3638caa039fe.md)|Returns the distance, in points, from the left edge of the media control bounding box to the left edge of the  **Slide**. Read-only.|
| [MediaControlsTop](e530dad8-ab23-e37d-fde3-5edb79c51365.md)|Returns the distance, in points, from the top edge of the media control bounding box to the top edge of the  **Slide** object. Read-only.|
| [MediaControlsVisible](0d9d9807-bd5f-4633-001f-9aa4f63c5c28.md)|Indicates whether the media controls are visible. Read-only.|
| [MediaControlsWidth](02a81c3e-c19d-183a-c9e4-08decf01d30f.md)|Returns the width, in points, of the media control bounding box. Read-only.|
| [Parent](0e21d9e5-48d3-2a4c-fe64-8a33e4341417.md)|Returns the parent object for the specified object.|
| [PointerColor](29f4c5e0-0927-1dbb-7bc9-b147ae38ff88.md)|Returns a  **ColorFormat** object that represents the pointer color for the specified presentation during one slide show. Read-only.|
| [PointerType](58f40da1-ae25-4604-86bc-6fb884b8fd16.md)|Returns or sets the type of pointer used in the slide show. Read/write.|
| [PresentationElapsedTime](6f710354-1691-4673-f83f-395d510d6999.md)|Returns the number of seconds that have elapsed since the beginning of the specified slide show. Read-only.|
| [Slide](4fdee96b-9b0d-64ba-19de-b810bf07987b.md)|Returns a  ** [Slide](afe42344-6898-00d2-ecc1-b0ed23a71fe8.md)** object that represents the slide that's currently displayed in the specified slide show window view. Read-only.|
| [SlideElapsedTime](e9250ea3-c37e-ebed-c8a8-9774dab77f37.md)|Returns the number of seconds that the current slide has been displayed. Read/write.|
| [SlideShowName](63efa2d8-7321-dc72-3c25-ab5ab4ba5c0a.md)|Returns the name of the custom slide show that's currently running in the specified slide show view. Read-only.|
| [State](749fe106-fed4-6ccc-f127-2e8a80196309.md)|Returns or sets the state of the slide show. Read/write.|
| [Zoom](92a303f0-b37f-a017-bedb-6537e235f753.md)|Returns the zoom setting of the specified slide show window view as a percentage of normal size. Read-only.|
