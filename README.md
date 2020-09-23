<div align="center">

## Classes and Active X User Controls and Property Pages\.\.\.\.\.oh my\.


</div>

### Description

Demonstrates how to use classes, Active X user controls and property pages among other features ot Visual Basic. Also includes a masking control as a demonstration application.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-09-20 13:24:38
**By**             |[Sean Street](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sean-street.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD100219202000\.zip](https://github.com/Planet-Source-Code/sean-street-classes-and-active-x-user-controls-and-property-pages-oh-my__1-11573/archive/master.zip)





### Source Code

<FONT SIZE=3>
<P>Included with this tutorial is an excellent project of a control that handles masking for several different masking data. These masks
include:</P>
<P>Date masking with long, medium and short date types. This is probably the best feature of the control. The control actually attempts to
predict the month and day that is being submitted. If the developer has selected long date as the mask type, for example, and the user enters the
letter "F", then the control will automatically return "February ". The same is true if the user enters a number "2", since "February is
understandably the second month of the year.</P>
<P>Phone masking. Allows the developer to define whether parenthesis, dashes and/or spaces are allowed.</P>
<P>Social Security Number masking. Allows the developer to define whether dashes are allowed or not.</P>
<P>Zip code masking. Allows the developer to define either 5 or 9 numbered zip codes.</P>
<P>Email masking. Only accepts well-formed email address.</P>
<P>Custom masking. Allows the developer to decide if aplha characters are allowed, numeric characters, and user-defined characters. Also, allows
a maximum length of the control to be defined.</P>
<P>The source code for this control is provided as well (although it was written in VB6 with SP4), as well as a sample application that uses each
type of masking format. Please feel free to alter andor distribute the code as desired.</P>
<P>Please direct any questions, comments, suggestions, and/or bugs to <a href="mailto:sean28681@yahoo.com">Sean L. Street</a></P>
<BR>
<B><P>Classes</P></B>
<P>A class object can be thought of as a template of sorts. The way I&#8217;ve adapted to teaching my students is as follows. Imagine that you are
standing in front of a vending machine that accepts only quarters, dimes and nickels; a change machine that accepts only one dollar, five dollar,
and ten-dollar bills; and a bubble gum machine that accepts only pennies. First, you must decide what you desire, then you determine what type of
currency you have in your pocket (the pocket class). Lets assume that you have a five-dollar bill, and four pennies. You&#8217;ve determined that you
want a candy bar from the vending machine that costs 50 cents and a piece of bubble gum that costs a penny. You inset the penny into the
bubblegum "class" and low and behold! out comes a piece of bubble gum. Next, you&#8217;re stuck in a dilemma. You only have a five and the vending
machine accepts only silver change. Being the genius that you are, you realize that you need to first inset your five into the change machine and
then take the result of that process and insert a portion of it into the vending machine. I use this scenario to also describe the purpose of
child (also called sub) classes I do not use any child classes here, so I&#8217;m not going to go into detail about them here. The relationship between
classes and our example is this:</P>
<P>Classes are like templates that only accept certain types of data. They can return results determined by the inputted data, or they can just
be storage of data in either case, they are not used until they are needed. In some cases, when compiled in a DLL for example, classes can be
used by other people. This is a great way to reduce in code writing. Lets look at our scenario again.</P>
<P>In this example we basically have four classes:</P>
<P>Pocket Class (this class stores your currency of any type)</P>
<P>Change Class (this class converts dollar bills into silver change)</P>
<P>Vending Class (this class converts silver change into food)</P>
<P>BubbleGum Class (this class converts pennies into gum)</P>
<P>Let&#8217;s say that you are happily eating your candy bar, when your spouse witnesses your delights. Your spouse demands the contents of your
Pocket Class so that they may indulge in the pleasures of the Vending Class as well. In this case, you have just shared the Pocket Class with
another &quot;application.&quot; </P>
<P>Our masking control uses classes in somewhat the same method. First, we are passed values from the interface. Next, we determine which class
we need to use. Then, we filter that data accordingly. Finally, we return the results of our processes back to the interface.</P>
<B><P>User Controls</P>
</B><P>An Active-X User Control is very similar to a Visual Basic form. In our case, we have a textbox on our user control. We then handle all
events from that textbox within the user control itself. The only thing the user sees is the result of our filtering and manipulation of the
passed data. We allow the user to set properties to allow some flexibility of the outcome of the data, but we ultimately control the processing
of data within our control. This allows our users to simply place the control on their forms and demand the respective output and not have to
negotiate the inputted data.</P>
<B><P>Property Pages</P></B>
<P>The property page is the interface that allows our users to define the type of masking that is to take place. When a user &#8216;right clicks&#8217; our
control at design time, the control will display our property pages. To have a property page appear when a user &quot;right clicks&quot; our
control, we have to include it in the PropertyPages property of the control itself. When we click the ellipse of the PropertyPages property of
the control, we get a list of our user-defined properties as well as a few predefined ones. If we wanted to include the predefined
&quot;Font&quot; property page, we would simply place a check next to it on the Connect Property Pages screen. Once we have defined all of the
property pages we want to display, they will appear when the user right clicks the control at design time. Inside of our property pages, we have
functions to manage when properties are changed. If a property is changed, then the changed flag is raised and in turn, the Apply button of the
property page is enabled. When the Apply button is clicked or the property page looses focus, then the PropertyPage_ApplyChanges event of the
property page is fired. When this event is fired, we save the changes to our instantiated class object (See Classes above). Code within the
class object will then save the changes out to an INI file. In our property pages we handle the functions of loading the data for the property
page. This is accomplished in the PropertyPage_Paint event of the page. Here, we determine if the selected page is the type of mask selected.
If it is, we allow all the controls on the page to be visible. Otherwise we make the controls invisible. (View the comments within the pagDate
property page of the project).</P></FONT>

