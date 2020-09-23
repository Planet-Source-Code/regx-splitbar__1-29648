
This control makes it very easy to add a horizontal split bar to your application.
Eventually this control will include vertical splitbars as well.

To use this control:

1. Place the control on a form.
2. Place controls above and below the splitbar control.
   This control needs at least one control above it and one control
   below it to function. The idea is that instead of using code to
   set the maximum and minimum scrolling positions controls are used instead.
   This makes it easy to visualize the range of motion for the control.
   Also, you can reposition the controls in response to form resize events, and
   the splitbar will use the new positions. Controls that are used only for
   setting splitbar ranges can be hidden at run time.

In the future I will probaly change this so that the splitbar uses its current position
for the maximum top or maximum bottom position if no controls are added to the top or bottom.
3. In the form load event add references to the controls to the splitbar using the
   AddControlTop and AddControlBottom events.
i.e. 
    splitbar1.AddControlTop Text1
    splitbar1.AddControlBottom Text2
    splitbar1.Update
 
Calling  splitbar.Update will align the bottom of all TopControls controls with the top of the splitbar
         and the tops of the BottomControls to the bottom of the splitbar.


Known bugs.

1. Currently the splitbar control brings itself to the top when a user moves it.
   This was intended to make using the control easier, but it will hide any labels contained in
   the control.

I think thats it, if you find more please let me know.
Also, please send me any ideas you may have to make this control better.