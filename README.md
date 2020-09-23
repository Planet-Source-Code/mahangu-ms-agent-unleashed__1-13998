<div align="center">

## Ms Agent Unleashed


</div>

### Description

This tutorial covers everything - from building your first Ms Agent app, to Using Ms Agent on your website, to using the Office Character files in your apps. Because of popular demand, Speech Recognition section added. Also features a section describing making your own character files. This tutorial shows you how to do nearly everything the Agent Control can do. 
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mahangu](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mahangu.md)
**Level**          |Beginner
**User Rating**    |4.8 (121 globes from 25 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mahangu-ms-agent-unleashed__1-13998/archive/master.zip)





### Source Code

<p align="center">&nbsp;</p>
<p align="center"><b><font face="Arial" color="#000080" size="5">Ms Agent
Unleashed</font></b></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080">Introduction</font></p>
<p align="left"><font face="Arial">This tutorial covers everything - from
building your first Ms Agent app, to Using Ms Agent on your website, to using
the Office Character files in your apps. Because of popular demand, Speech
Recognition section added. Also features a section describing making your own
character files. This tutorial shows you how to do nearly everything
the Agent Control can do.&nbsp;</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080">Understanding this tutorial</font></p>
<p align="left"><font color="#000000" face="Arial">Through out this tutorial you
will see text like this - <i>italic text and </i></font><font face="Arial" color="#008000"><i>green
italic text</i> . </font><font face="Arial" color="#000000">The normal <i>italic
text</i> means that the text is code and can be copied and pasted straight into
your application. The </font><i><font face="Arial" color="#008000">green italic
text</font></i><font face="Arial" color="#000000"> means that the text is a
comment (you will often see this type of text beside code) that was place to
show you how to do something or to give you an example.</font></p>
<p align="center">&nbsp;</p>
<p align="center"><b><font face="Arial" color="#000080" size="4">Index</font></b></p>
<p align="left"><font face="Arial" color="#000080"><b>Getting Started</b></font><font face="Arial" color="#000000">
- <i>Provides all the data you need to jump start your Agent application</i></font></p>
<p align="left"><b><font face="Arial" color="#000080">Declaring the Character
File</font></b><font face="Arial" color="#000000"> - <i>Shows how to declare the
Character file for use in VB</i></font></p>
<p align="left"><font face="Arial"><font color="#000080"><b>Initializing the
Character</b></font> - <i>Shows how to initialize the Character file</i></font></p>
<p align="left"><font face="Arial" color="#000080"><b>Getting to Know
The Different Characters </b></font><font face="Arial"><i>- Familiarize yourself
with the different characters</i></font></p>
<p align="left"><font face="Arial"><font color="#000080"><b>Displaying Various
Animations</b></font> -&nbsp;<i> Shows how to get the Character to display
various animations</i></font></p>
<p align="left"><font face="arial "><font color="#000080"><b>Using Ms Agent With
VB Script</b></font> - <i>Shows you how to use Ms Agent with VB Script</i></font></p>
<p align="left"><font face="Arial" color="#000080"><b>Using the Office Character
Files in Your Ms Agent Apps</b></font><i><font face="arial "> - Shows how to include
office character files in your applications</font></i></p>
<p align="left"><font face="Arial" color="#000080"><b>Speech Recognition</b></font><i><font face="arial ">
- Shows how to initialize speech recognition&nbsp;</font></i></p>
<p align="left"><font face="Arial" color="#000080" size="3"><b>Making your Own
Character Files</b></font><i><font face="arial "> - Describes how to create your
own character files for use with Ms Agent</font></i></p>
<p align="left"><font face="Arial"><font color="#000080"><b>Events and
Properties of the Agent Control</b></font> - <i>Describes the Events and
Properties of the Agent Control</i></font></p>
<p align="left"><font face="Arial"><font color="#000080"><b>Fun Agent Code to Add to
your Applications</b></font> - <i>Gives some cool code which makes the Character
do some fun things</i></font></p>
<p align="left"><font face="Arial"><font color="#000080"><b>Examples of
How&nbsp; you can use the Agent Control</b></font> - <i>Gives some ideas as to
how you can use the Agent Control</i></font></p>
<p align="left"><font face="Arial" color="#000080"><b>Cool Web Links</b></font><font face="Arial"><i>
- Links to the best Ms Agent resource sites on the web</i></font></p>
<p align="left"><font face="Arial"><font color="#000080"><b>Frequently Asked
Questions</b></font> - <i>Various related questions and answers.</i></font></p>
<p align="center">&nbsp;</p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Getting Started</font></p>
<p align="left"><font face="arial ">In a nutshell, Ms Agent is an ActiveX
control, created by Microsoft that lets you add a user friendly touch to your
apps via the use of animated characters.</font></p>
<p align="left"><font face="arial ">In order to use this tutorial you will need
Microsoft Visual Basic 5 or 6 (parts of this tutorial may work in VB 4 if you
have Agent 1.5 installed). I am not sure about VB 7 (VB.NET). You will also need the Speech Synthesis libraries
from MSDN along with a Microsoft Agent Character File (*.acs file). An open mind and good cup of coffee (or any other
preferred beverage :)
will be helpful.</font></p>
<p align="left"><font face="Arial" color="#000000">MS Agent is an ActiveX
control supplied with Microsoft Visual Basic 5 and 6. It can be used in many
other ways but the most popular use is for creating 'Desktop Pets'. At the
moment there are 4 different characters to chose from - Peedy the Parrot, The
Genie, Merlin the Wizard and Robby the Robot. In this tutorial I have used
Peedy the Parrot as an example.</font></p>
<p align="left"><font face="Arial" color="#000000">To start making your first
Microsoft Agent application, open Visual Basic and chose standard exe. Then
right click the toolbar and add the the Microsoft Agent Control. You will see a
new Icon (it looks like a secret agent with sunglasses). Then
double click on the icon on the toolbar to place the control on the form. You
can rename this control&nbsp; to whatever you want but in the code I'm going to
call it Agent1.</font></p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Declaring the Character
file</font></p>
<p align="left"><font face="Arial" color="#000000">We need to to tell VB that we
are using the character file so we need add the following code to the general
declarations.</font></p>
<p align="left"><font face="Arial"><i>Dim char As IAgentCtlCharacterEx '<font color="#008000">Declare
the String char as the Character file</font></i></font></p>
<p align="left"><font face="Arial"><i>Dim Anim as String <font color="#008000">'Dim
the Anim string which we will use later on (declaring this will make it easy for
us to change the character with ease, later on)</font>
</i></font></p>
<p align="left"><i><font face="Arial" color="#000000">Char.LanguageID = &amp;H409
</font><font face="Arial" color="#008000">'This code is optional. The code
worked fine without it but we will add it for usability purposes (it sets the
language ID to English)</font></i></p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Initializing the
Character</font></p>
<p align="left"><font face="Arial">We need to tell VB, who the character is and
where his *.acs file is. So we'll use the following code.</font></p>
<p align="left"><font face="Arial"><i>Anim = "Peedy"&nbsp;&nbsp;&nbsp; <font color="#008000">'We
set the Anim String to &quot;Peedy&quot; . You can set this to Genie, or Merlin,
or Robby too.</font><br>
</i></font></p>
<p align="left"><font face="Arial"><i>Agent1.Characters.Load Anim, Anim &amp; ".acs"&nbsp;&nbsp;&nbsp;
<font color="#008000">'This is how we tell VB where to find the character's acs
file. VB by default looks in the <a href="file:///C:/Windows/MsAgent/Chars/">C:\Windows\MsAgent\Chars\</a>
folder for the character file</font><br>
</i></font></p>
<p align="left"><font face="Arial"><i>Set char = Agent1.Characters(Anim)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font color="#008000">'Remember we declared the char string earlier? Now we set
char to equal Agent1.Charachters property. Note that the because we used the
Anim string we can now change the character by changing only one line of code.</font><br>
</i></font></p>
<p align="left"><font face="Arial"><i>char.AutoPopupMenu = False <font color="#008000">'So
the Character wont keep displaying it's annoying popup menu every time you right
click him. You can now add your own popup menu (see examples).</font></i></font></p>
<p align="left"><font face="Arial"><i>Char.Show <font color="#008000">'Shows the
Character File (If set to &quot;Peedy&quot; he comes flying out of the
background)</font></i></font></p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Getting to Know
The Different Characters</font></p>
<p align="center"><font face="Arial">As far as I know, there are 4 default
characters you can use with Ms Agent. You can download them all from the Ms
Agent Developers Website ( <a href="http://msdn.microsoft.com/msagent">http://msdn.microsoft.com/msagent</a>
). Although you can configure each character to your own liking, they tend to
convey different types of impressions.&nbsp;</font></p>
<p align="left"><font face="Arial" color="#000080"><b>Peedy</b> </font><font face="Arial" color="#000000">-
The first agent character (I think). He is a temperamental parrot (that's the
way I see him). I use him mostly to add sarcasm to my apps. Has an (sort of)
annoying voice - squeaky in parroty sort of way. You use him to some cool stuff
though.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Genie</b> </font><font face="Arial" color="#000000">-
Cool little guy to add to your apps. Can do some neat stuff too! Use him to add
a touch of class and mystery to your apps. Has an OK voice and has a cool way of
moving around.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Merlin</b> </font><font face="Arial" color="#000000">-
Your friendly neighborhood Wizard! Always has the look that he is total control. Also has
a vague look of incomprehension (that's the way I see it!). Useful little dude
but I don't like the way he moves around (wears beanie and flies).</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Robby</b> </font><font face="Arial" color="#000000">-
Probably the newest addition to the series. Looks like an Robot from some space
movie. Has a very metallic, robotic voice. Moves around using jetpacks.</font></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" color="#000000">What? You don't like any of
these characters? Wanna create you're own? It's not easy.. but you can give it a
shot... Just visit the MSDN page for Ms Agent (check FAQs for web
address).&nbsp;</font></p>
<p align="center"><font face="Arial" color="#000000">You can also download some
customs files. The Agentry, a cool site that has lots of sample applications,
also has over 300 character files and some of them are free. Look for the URL in
the 'Cool Web Links' section.</font></p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Displaying Various
Animations</font></p>
<p align="left"><font face="Arial">Through code, we can make the character do
some cool stuff. Apart from talking he can do <font color="#000000">various
interesting things. The following code may be pasted into any event in VB (Form_Load,
Command1_Click).&nbsp;</font></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Showing the Character</b></font></p>
<p align="left"><font face="Arial" color="#000000">This code is used to bring
the character on to the screen.</font></p>
<p align="left"><font face="Arial"><i>char.show</i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Hiding the Character</b></font></p>
<p align="left"><font face="Arial" color="#000000">This code is used to hide the
character (take him off the screen).</font></p>
<p align="left"><font face="Arial"><i>char.hide</i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Talk</b></font></p>
<p align="left"><font face="Arial" color="#000000">The code for this is
relatively simple and this works with every character. </font><font face="Arial"><font color="#000000"></font></font><font color="#000000"><font face="Arial">You
can customize this code for him to say anything. The text appears in a speech
bubble but can also be heard.</font></font></p>
<p align="left"><i><font face="Arial" color="#000000">Char.Speak &quot;Your
Message Here&quot; </font><font face="Arial" color="#008000">'Says &quot;Your
Message Here&quot;</font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Think</b></font></p>
<p align="left"><font face="Arial" color="#000000">The code for this is
relatively simple and this works with every character. You
can customize this code and make him think of anything. The text appears in a
thought bubble and cannot be heard.</font></p>
<p align="left"><i><font face="Arial" color="#000000">Char.Think &quot;Your
Message Here&quot; </font><font face="Arial" color="#008000">' &quot;Your
message here&quot; appears in a though bubble</font></i></p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Move To
Somewhere Else On The Screen</b></font></p>
<p align="left"><font face="Arial" color="#000000">This code too is pretty
simple and works on every character. You can move him anywhere on the screen be
changing the co ordinates. Please note that screen co ordinates vary from
resolution to resolution. For example on a 640 x 480 resolution monitor 300,500
is off the screen wile on a 800 x 600 monitor the co ordinates are on the
screen.</font></p>
<p align="left"><i><font face="Arial">char.MoveTo 300, 300</font></i><i><font face="Arial">
<font color="#008000">'This code will move him to the screen co ordinates
300,300</font></font></i></p>
<p align="left"><font face="arial ">Also note that in the code <i>300,300</i> we
are referring to the screen as x , y (horizontal , vertical).</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Stay In His
Rest Pose</b></font></p>
<p align="left"><font face="Arial" color="#000000">This code brings him back to
the way he was started</font></p>
<p align="left"><i><font face="Arial" color="#000000">char.play &quot;Restpose&quot;
</font><font face="Arial" color="#008000">'Note - To get out of the rest pose
you will have to use the char.stop function (see below)</font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Stop Whatever
He Is Doing</b></font></p>
<p align="left"><font face="Arial">Sometimes you may need to stop the Character
from doing something. This code makes him stop everything and wait.</font></p>
<p align="left"><i><font face="Arial">char.stop <font color="#008000">'Character
stops whatever he is doing</font></font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Read, Write,
Process and Search</b></font></p>
<p align="left"><font face="Arial">The character can various animations that may
prove useful in your applications.&nbsp;</font></p>
<p align="left"><font face="Arial"><i>char.Play &quot;Write&quot; <font color="#008000">'The
character writes for a while and then stops</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Play &quot;Writing&quot; <font color="#008000">'The
character writes until the char.stop function is executed</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Play &quot;Read&quot; <font color="#008000">'The
character reads for a while and then stops</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Play &quot;Reading&quot; <font color="#008000">'The
character reads until the char.stop function is executed</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Play &quot;Process&quot; <font color="#008000">'The
character processes for a while and then stops</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Play &quot;Processing&quot; <font color="#008000">'The
character processes until the char.stop function is executed</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Play &quot;Search&quot; <font color="#008000">'The
character searches for a while and then stops</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Play &quot;Searching&quot; <font color="#008000">'The
character searches until the char.stop function is executed</font></i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Show Facial
Expressions</b></font></p>
<p align="left"><font face="Arial">The character can show various facial
expressions that may be useful in your application.</font></p>
<p align="left"><i><font face="Arial">char.play &quot;Acknowledge&quot; <font color="#008000">'This
code makes the character acknowledge something</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Alert&quot; <font color="#008000">'This
code makes the character look alert&nbsp;</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Blink&quot; <font color="#008000">'This
code makes the character blink</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Confused&quot; <font color="#008000">'This
code makes the character look confused</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Decline&quot; <font color="#008000">'This
code makes the character decline something</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;DontRecognize&quot; <font color="#008000">'This
code makes the character look like he doesn't recognize something</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Hearing_1&quot; <font color="#008000">'This
code makes the character look like he is listening (left)</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Hearing_2&quot; <font color="#008000">'This
code makes the character look like he is listening (right)</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Hearing_3&quot; <font color="#008000">'This
code makes the character look like he is listening (both sides)</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Hearing_4&quot; <font color="#008000">'This
code makes the character look like he is listening (does not work on peedy)</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Pleased&quot; <font color="#008000">'This
code makes the character look pleased</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Sad&quot; <font color="#008000">'This
code makes the character look sad</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Surprised&quot; <font color="#008000">'This
code makes the character look surprised</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Uncertain&quot; <font color="#008000">'This
code makes the character look uncertain</font></font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Look Somewhere</b></font></p>
<p align="left"><font face="Arial">The character can look at different angles.</font></p>
<p align="left"><i><font face="Arial">char.play &quot;LookDown&quot; <font color="#008000">'Looks
Down</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;LookDownBlink&quot;&nbsp; <font color="#008000">'Looks
and Blinks</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;LookDownReturn&quot; <font color="#008000">'Stops
looking and returns to restpose</font></font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><i><font face="Arial">char.play &quot;LookUp&quot; <font color="#008000">'Looks
Up</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;LookUpBlink&quot; '<font color="#008000">Looks
and Blinks</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;LookUpReturn&quot; <font color="#008000">'Stops
looking and returns to restpose</font></font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><i><font face="Arial">char.play &quot;LookRight&quot; <font color="#008000">'Looks
to the Right</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;LookRighBlink&quot; <font color="#008000">'Looks
and Blinks</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;LookRightReturn&quot; <font color="#008000">Stops
looking and returns to restpose</font></font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><i><font face="Arial">char.play &quot;LookLeft&quot; <font color="#008000">'Looks
to the Left</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;LookLeftBlink&quot; <font color="#008000">'Looks
and Blinks</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;LookLeftReturn&quot; <font color="#008000">'Stops
looking and returns to restpose</font></font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making Him Do Various
Gestures</b></font></p>
<p align="left"><font face="Arial">The character can do various gestures that
can be quite useful.</font></p>
<p align="left"><font face="Arial"><i>char.play &quot;GestureUp&quot; <font color="#008000">'Gestures
Up</font></i></font></p>
<p align="left"><font face="Arial"><i>char.play &quot;GestureRight&quot; <font color="#008000">'Gestures
Right</font></i></font></p>
<p align="left"><font face="Arial"><i>char.play &quot;GestureLeft&quot; <font color="#008000">'Gestures
Left</font></i></font></p>
<p align="left"><font face="Arial"><i>char.play &quot;GestureDown&quot; <font color="#008000">'Gestures
Down</font></i></font></p>
<p align="left"><i><font face="Arial" color="#000000">char.play
&quot;Explain&quot; </font><font face="Arial" color="#008000">&quot;Explains
Something</font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;GetAttention&quot; <font color="#008000">'Gets
the users attention</font></font></i></p>
<p align="left"><i><font face="Arial">char.play &quot;Greet&quot; <font color="#008000">'Greets
the User (by action)</font></font></i></p>
<p align="left"><font face="Arial" color="#000000"><i>char.play
&quot;Announce&quot;&nbsp;</i></font></p>
<p align="left"><i><font face="Arial" color="#000000">char.play &quot;Congratulate_1&quot;
</font><font color="#008000"><font face="Arial">'</font><font face="Arial">Congratulates</font><font face="Arial">
user&nbsp;</font></font></i></p>
<p align="left"><font face="Arial" color="#000000"><i>char.play &quot;Congratulate_2&quot;
</i></font><i><font face="Arial" color="#008000">'</font><font color="#008000"><font face="Arial">Congratulates</font><font face="Arial">
user </font></font></i></p>
<p align="left"><font face="Arial"><i>char.play &quot;DoMagic1&quot; <font color="#008000">'Does
Magic 1 - Can be used with DoMagic2</font></i></font></p>
<p align="left"><font face="Arial"><i>char.play &quot;DoMagic2&quot;</i></font></p>
<p align="left"><font face="Arial"><i>char.play &quot;StartListening&quot; <font color="#008000">'Starts
Listening</font></i></font></p>
<p align="left"><font face="Arial"><i>char.play &quot;StoptListening&quot; <font color="#008000">'Stops
Listening</font></i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Making him Gesture at a
specific location on Screen</b></font></p>
<p align="left"><font face="Arial" color="#000000">Using the GestureAt property
you can get the Character to point at a specific screen co ordinate. More useful
than GestureRight and GestureLeft because using this you can point diagonally
too.</font></p>
<p align="left"><font face="Arial"><i>char.GestureAt 300,300 <font color="#008000">'Character
points at screen co ordinate 300,300</font></i></font></p>
<hr>
<p align="left">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Events and
Properties of the Agent Control</font></p>
<p align="center">&nbsp;</p>
<p align="left"><font face="Arial" size="4" color="#000080">Events</font></p>
<p align="left"><font face="Arial" color="#000080"><b>Using the Agent1_IdleStart
event to set what the Agent does when He is Idle</b></font></p>
<p align="left"><font face="Arial">You can place code in the Agent1_IdleStart
event to tell VB what the agent does when he is idle.</font> <font face="Arial">The
Agent can do the following idle stuff. Please note that some functions may not
work for some characters. You can put the following functions in a loop or just
let them run. Also note that some functions cannot be stopped unless the <i>char.stop</i>
command is used. You may also include any other functions in the
Agent1_IdleStart event.</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle1_1&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle1_2&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle1_3&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle1_4&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle1_5&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle1_6&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle2_1&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle2_2&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle2_3&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle3_1&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle3_2&quot;</font></p>
<p align="left"><font face="Arial">char.play &quot;Idle3_3&quot; <i><font color="#008000">'This
one works only for Peedy I think! - He listens to music!</font></i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Using the Agent1_Complete
event to set what the Agent does when He is finished idling</b></font></p>
<p align="left"><font face="Arial">This tells VB what to with the agent once he
is finished idling. Example -</font></p>
<p align="left"><i><font face="Arial">char.play &quot;Restpose&quot;<font color="#008000">
'This will put the character in his default rest pose</font></font></i></p>
<p align="left"><font face="Arial">&nbsp;</font></p>
<p align="left"><font face="Arial" color="#000080"><b>Using the Agent1_Click
event to Set what happens when the Character is clicked</b></font></p>
<p align="left"><font face="Arial">You can place some code in the Agent1_Click
event to tell VB what to do when the user clicks on the character.&nbsp; You can
place almost any command here. Example -</font></p>
<p align="left"><i><font face="Arial">char.play &quot;Alert&quot;</font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Using the Agent1_Move
event to Set what happens when the Character is moved</b></font></p>
<p align="left"><font face="Arial">You can place some code in the Agent1_Move
event to tell VB what to do when the user moves the character.&nbsp; You can
place almost any command here. Example -</font></p>
<p align="left"><i><font face="Arial">char.play &quot;Surprised&quot;</font></i></p>
<p align="left"><font face="Arial" color="#000080"><b>Using the Agent1_DragStart
event to Set what happens when the user starts to drag the Character</b></font></p>
<p align="left"><font face="Arial">You can place some code in the
Agent1_DragStart event to tell VB what to do when the user starts to drag the
character.&nbsp; You can place almost any command here. Example -</font></p>
<p align="left"><i><font face="Arial">char.play &quot;Think&quot;</font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Using the Agent1_DragStop
event to Set what happens when the user stops dragging the Character</b></font></p>
<p align="left"><font face="Arial">You can place some code in the
Agent1_DragStop event to tell VB what to do when the user stops dragging the
character.&nbsp; You can place almost any command here. Example -&nbsp;</font></p>
<p align="left"><i><font face="Arial">char.play &quot;Blink&quot;</font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial"><b>Using the Agent1_BalloonHide
event to Set what happens when the Character's speech balloon is shown</b></font></p>
<p align="left"><font face="Arial">Using this event you can set what happens
every time the speech balloon is shown (basically every time the character
starts speaking).</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial"><b>Using the Agent1_BalloonShow
event to Set what happens when the Character's speech balloon is hidden</b></font></p>
<p align="left"><font face="Arial">Using this event you can set what happens
every time the speech balloon is hidden (basically every time the character
stops speaking).</font></p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" size="4" color="#000080">Properties</font></p>
<p align="left"><font face="Arial" color="#000080"><b>Using the SoundEffectsOn
property to switch the Characters sound effects on / off</b></font></p>
<p align="left"><font face="Arial">Using this property you can toggle the
characters sound effects on an off. Useful if you want the character to stay
silent for a while</font></p>
<p align="left"><font face="Arial"><i>char.SoundEffectsOn = True <font color="#008000">Turns
sound effects on</font></i></font></p>
<p align="left"><font face="Arial"><i>char.SoundEffectsOn = False <font color="#008000">'Turns
sound effects off</font></i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Using the IdleOn
property to toggle the Character's idle mode on / off</b></font></p>
<p align="left"><font face="Arial">Using this property you can toggle the
character's idle mode on an off.&nbsp;</font></p>
<p align="left"><font face="Arial"><i>char.IdleOn = True <font color="#008000">'Sets
Idle Mode On</font></i></font></p>
<p align="left"><font face="Arial"><i>char.IdleOn = False <font color="#008000">'Sets
Idle Mode Off</font></i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Using the AutoPopupMenu
property to toggle the default (Agent's) popup menu on and off</b></font></p>
<p align="left"><font face="Arial">Using this propert you can set the agent's
popup menu on or off. This menu has only one option (hide) ,so by it is not
really useful. If you want a popup menu for your character see the Agent Right
Click Popup Menu Example (below) on how to create custom popup menus. As you may
have noticed, in the 'Initializing the Character' section I have turned off the
auto popupmenu. Never the less you can use the following code to toggle it on or
off.</font></p>
<p align="left"><font face="arial "><i>char.AutoPopupMenu = True <font color="#008000">'Turns
Auto PopMenu On</font></i></font></p>
<p align="left"><font face="Arial"><i>char.AutoPopupMenu = False </i></font><font face="arial "><i><font color="#008000">Turns
Auto PopMenu Off</font></i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial"><b>Using the Connected
property to set whether the Agent is connected to the Microsoft Agent Server</b></font></p>
<p align="left"><font face="Arial">Using this you can set whether the control is
connected to the Microsoft Agent Server (useful for creating client / server
applications).</font></p>
<p align="left"><i><font face="Arial">char.Connected = True <font color="#008000">'Not
Connected</font></font></i></p>
<p align="left"><i><font face="Arial">char.Connected = False <font color="#008000">'Connected</font></font></i></p>
<hr>
<p align="left">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Using Ms Agent
with VB Script</font></p>
<p align="center"><font face="Arial">Ms Agent can be used in VB script too. VB
script 2.0 is needed to do so. Here is an example. Using VB script is very
useful if you want to include MS Agent on your web page. Please note - I am not
too familiar with VB script so If there are any syntax errors please let me
know.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial"><b>Using the Connected
property to set whether the Agent is connected to the Microsoft Agent Server</b></font></p>
<p align="left"><font face="Arial">Using this you can set whether the control is
connected to the Microsoft Agent Server (useful for creating client / server
applications).</font></p>
<p align="left"><i><font face="Arial">char.Connected = True <font color="#008000">'Not
Connected</font></font></i></p>
<p align="left"><i><font face="Arial">char.Connected = False <font color="#008000">'Connected</font></font></i></p>
<p align="left">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Initializing The Character</b></font></p>
<p align="left"><font face="Arial">To initialize the character you will need to
contact the Agent Server.</font></p>
<p class="code"><font face="Arial"><i>&lt;SCRIPT LANGUAGE = &#8220;VBSCRIPT&#8221;&gt;</i></font></p>
<p class="code"><font face="Arial"><i>&lt;!&#8212;-</i></font></p>
<p class="code"><font face="Arial"><i>&nbsp;<span style="mso-spacerun: yes">&nbsp;&nbsp;
</span>Dim Char<font color="#008000"> 'Declare the String Char</font></i></font></p>
<p class="code"><span style="mso-spacerun: yes"><i><font face="Arial">&nbsp;&nbsp;&nbsp;
</font></i></span><i><font face="Arial">Sub window_OnLoad <font color="#008000">'Window_Onload
Event</font></font></i></p>
<p class="code"><span style="mso-spacerun: yes"><i><font face="Arial">&nbsp;&nbsp;
</font></i></span><i><font face="Arial">AgentCtl.Characters.Load
&quot;Genie&quot;, &quot;http://agent.microsoft.com/characters/v2/genie/genie.acf&quot;</font></i></p>
<p class="code"><font face="Arial" color="#008000"><i>&nbsp;<span style="mso-spacerun: yes">&nbsp;&nbsp;
</span>&#8216;Create an object with reference to the character on the Microsoft
server&nbsp;</i></font></p>
<p class="code"><span style="mso-spacerun: yes"><i><font face="Arial">&nbsp;&nbsp;
</font></i></span><i><font face="Arial">set Char= AgentCtl.Characters
(&quot;Genie&quot;) <font color="#008000">'Set the the Char string to = The
Agent Cotnrol</font></font></i></p>
<p class="code"><i><font face="Arial">Char.Get &quot;state&quot;,
&quot;Showing&quot;&nbsp;</font></i><font face="Arial"><i><span style="mso-spacerun: yes">
</span><font color="#008000">&#8216;Get the Showing state animation</font></i></font></p>
<p class="code"><i><font face="Arial">Char.Show <font color="#008000">'Show the
Character</font></font></i></p>
<p class="code"><font face="Arial"><i>&nbsp;<span style="mso-spacerun: yes">&nbsp;&nbsp;
</span>End Sub</i></font></p>
<p class="code"><font face="Arial"><i>&nbsp;--&gt;</i></font></p>
<p class="code"><span style="mso-spacerun: yes"><i><font face="Arial">&nbsp;&nbsp;
</font></i></span><i><font face="Arial">&lt;/SCRIPT&gt;</font></i></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Sending Requests to the
Server</b></font></p>
<p class="code"><font face="Arial">You will need to send requests to the agent
server in order to do certain commands.</font></p>
<p class="code"><font face="Arial"><i><span style="mso-spacerun: yes">&nbsp;&nbsp;
</span>Dim Request</i></font></p>
<p class="code"><font face="Arial"><i><span style="mso-spacerun: yes">&nbsp;&nbsp;
</span>Set Request = Agent1.Characters.Load (&quot;Genie&quot;, &quot;<span style="text-decoration:none;text-underline:none" class="MsoHyperlink">http://agent.microsoft.com/characters<a name="_Hlt390052700">/v2/genie/</a>genie.acf</span>&quot;)
<font color="#008000">'Sets the request</font><o:p>
</o:p>
</i></font></p>
<p class="code"><font face="Arial"><i><span style="mso-spacerun: yes">&nbsp;&nbsp;
</span>If (Request.Status = 2) then <font color="#008000">'Request is in
Queue&nbsp;</font></i></font></p>
<p class="code"><font face="Arial" color="#008000"><i>'Add your code here (you
can send text to status bar or something)</i></font><i><font face="Arial"><o:p>
</o:p>
</font></i></p>
<p class="code"><font face="Arial"><i><span style="mso-spacerun: yes">&nbsp;&nbsp;
</span>Else If (Request.Status = 0) then <font color="#008000">'Request
successfully completed</font></i></font></p>
<p class="code"><font face="Arial" color="#008000"><i>'Add your code here (you
can do something like display the annimation)</i></font><i><font face="Arial"><o:p>
</o:p>
</font></i></p>
<p class="code"><font face="Arial"><i><span style="mso-spacerun: yes">&nbsp;&nbsp;
</span>End If</i></font></p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Showing Animations</b></font></p>
<p align="left"><font face="Arial">If you are using VB script you will need to
get the animations from a server using the <i>Get</i> method. For example the
following code will get all the 'Moving' animations which the character needs.</font></p>
<p align="left"><font face="Arial"><i><span style="mso-fareast-font-family: Times New Roman; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA">AgentCtl.Characters
(&quot;Peedy&quot;).Get &quot;Animation&quot;, &quot;Moving&quot;, True&nbsp;</span></i></font></p>
<p align="left"><font face="Arial">After an animation is loaded you should be
able to play it in the usual way.</font></p>
<p align="left">&nbsp;</p>
<hr>
<p align="left">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Using the Office
Character Files in Your Ms Agent Apps</font></p>
<p align="center"><font face="Arial" color="#000000">As far as I know, those
character files are not freeware and cannot be distributed except with office,
so please don't distribute them with your apps. Use this section for educational
purposes only.</font></p>
<p align="center"><font face="Arial" color="#000000">The office character files
can do very little (very few animations) and have no speech support, so you'd be
better off using the Ms Agent character files anyway. But hey, I was doing some
research and I found this out so I thought I would add this section. So here we
go...</font></p>
<p align="center"><font face="Arial" color="#000000">First find all the files on
your hard disk with the extension *.acs . You will see some familiar office
names too (e.g - Clippit, maybe Rocky). Just copy these files to the Ms Agent \
Chars folder. Then&nbsp; change the <i>Anim </i>property to equal the character
name. Example for Clippit -</font></p>
<p align="center"><i><font face="Arial" color="#000000">Anim = &quot;Clippit&quot;
</font><font face="Arial" color="#008000">Changes the Anim property to Clippit</font></i></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" color="#000000">You can't really do much
with these acs files, but I just thought I'd include this section.&nbsp;</font></p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Speech </font><font face="Arial" size="4" color="#000080">Recognition</font></p>
<p align="center"><font face="Arial" color="#000000">Another, feature of Ms
Agent is it's ability to recognize speech. You will need a microphone or a
similar gadget that lets you input speech into your PC. The following speech
engine can be used with Ms Agent. Check out the MSDN homepage for Ms Agent for
the latest speech engine updates. I've never tried to use the speech recognition
facility, so if you find any trouble please email me. If you want to find more
about voice recognition I recommend that you visit the MSDN site (URL in the FAQ
section).</font></p>
<p align="center">&nbsp;</p>
<p><font face="Arial" color="#000080"><b>L&amp;H TruVoice Text-To-Speech
-American English</b></font></p>
<p class="tabletext"><font face="Arial">This will recognize the usual American
Voice I think.</font></p>
<p class="tabletext"><font face="Arial">CLS ID =
B8F2846E-CE36-11D0-AC83-00C04FD97575</font><o:p>
</o:p>
</p>
<p class="tabletext"><font face="Arial">Version = 6,0,0,0</font></p>
<p>&nbsp;</p>
<p><font face="Arial">Here is some example code of how to create an object of
the speech engine (VB Script).</font></p>
<p>&nbsp;</p>
<p><font face="Arial"><i>&lt;OBJECT width=0 height=0<font color="#008000">
'Opens the Object Tag</font><br>
CLASSID=&quot;</i></font><i><font face="Arial">B8F2846E-CE36-11D0-AC83-00C04FD97575</font><o:p>
</o:p>
</i><font face="Arial"><i>&quot; <font color="#008000">'Tells the Class ID</font><br>
CODEBASE=&quot;#VERSION=</i></font><i><font face="Arial">6,0,0,0</font></i><font face="Arial"><i>&quot;&gt;<font color="#008000">
'Tells the version number</font><br>
&lt;/OBJECT&gt; <font color="#008000">'Closes the Object Tag</font></i></font></p>
<p align="center">&nbsp;</p>
<hr>
<p align="center"><font face="Arial" size="4" color="#000080">Making your Own
Character Files</font></p>
<p align="center"><font face="Arial">Sometime or the other you may need to
create a character that is unique to your application. This section describes
briefly how to do this.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Using the Microsoft Agent
Character Editor</b></font></p>
<p align="left"><font face="Arial">This tool is used to assemble, sequence and
time the frames. Also this is what is used to input other character details
(name, description) and to finally compile it to a acs file. You can download it
from the following URL -</font></p>
<p align="left"><a href="http://msdn.microsoft.com/msagent/charactereditor.asp"><font face="Arial">http://msdn.microsoft.com/msagent/charactereditor.asp</font></a></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Frames</b></font></p>
<p align="left"><font face="Arial">Every animation a character does is a timed
sequence of frames. It is like a cartoon movie or the little 'flip and look'
cartoons we used to make (remember those?!). Ok so we want to make the character
wave - we need to draw different shots of his hand at different stages of the
wave but we can still keep his body the same. This is called overlaying. You
just change the part of the image you want and let the rest be. The number of
frames in your animation can be any amount you chose but the usual is around 14
frames (takes around 6 seconds to process). This also helps to keep the size of
the animation small enough for transfer via the web. Frame size should be 128 x
128 (pixels). Using the Microsoft Agent Character Editor, you have the ability
to set how long a frame is displayed before the next one is shown. The typical
duration would be 10 hundredths of a second (about 10 frames a second).</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Creating Images</b></font></p>
<p align="left"><font face="Arial">Animations need Bitmaps (*.bmp files). The
images must be designed on a 256 colour pallete, preserving the standard windows
colours in their usual positions (first ten and last ten colours). That means
that your palette can use up to 236 other colours. Also if you use many other
colours, they may be remapped when your character is displayed on systems that
have a 8 bit colour setting. Using lots of different colours also may increase
the overall size of your character file. The 11th image in your palette is the
'alpha colour'. Agent will use this colour to render transparent pixels in your
application. This can also be changed using the Microsoft Agent Character
Editor.</font></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial">Author's Note - I have never really tried
doing this. For more information visit the MSDN Ms Agent page (see FAQ for URL).
If you attempt this and succeed (or don't succeed) please tell me.</font></p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Examples of
How&nbsp; you can use the Agent Control</font></p>
<p align="center">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Agent Right Click Popup
Menu Example</b></font></p>
<p align="left"><font face="Arial" color="#000000">This code is very useful if
you only want to have the agent visible on the screen and not the form. Now you
can set the agent to display a popup menu so that you wont have to display the
form. To use this you will need a Form called frmMain and in that form a Menu
Item called mnuMain. mnuMain must have submenus. You can type the following code
into the Agent1_Click Event</font></p>
<p align="left"><i><font face="Arial"><font color="#000000">if Button =
vbRightButton then frmMain.popupmenu mnuMain </font><font color="#008000">'This
code will display the popup menu only if the user right click son the age</font></font></i></p>
<p align="left"><font face="Arial">Now all you have to do is to add submenus and
functions to the mnuMain menu item!</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Agent</b></font><font face="Arial" color="#000080"><b>1_IdleStart
Event Example</b></font></p>
<p align="left"><font face="Arial" color="#000000">When the user does not click
on or interact with the Agent for a long time it automatically sets itself to
idle. So you may want to add some functions to make the agent do stuff while the
user is not working with him. You may add the following code to the
Agent1_IdleStart Event -</font></p>
<p align="left"><font face="Arial"><i>10<font color="#008000"> 'Specify line
number so that we can loop back later</font></i></font></p>
<p align="left"><i><font face="Arial" color="#000000">char.play
&quot;think&quot;&nbsp;</font></i></p>
<p align="left"><font face="Arial" color="#000000"><i>char.play &quot;read&quot;</i></font></p>
<p align="left"><font face="Arial" color="#000000"><i>char.play
&quot;write&quot;</i></font></p>
<p align="left"><font face="Arial"><i>Goto 10 <font color="#008000">'Tells VB to
go to the line number which was specified earlier</font></i></font></p>
<p align="left"><font face="Arial">You may also want to add the following code
to the Agent1_Click Event so that the character will stop doing hid idle part
when the user clicks on&nbsp; him - <i>char.stop</i></font></p>
<hr>
<p align="left">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Fun Agent Code to Add to
your Applications</font></p>
<p align="left"><font face="Arial" color="#000080"><b>Character 'Dive' Code
Example</b></font></p>
<p align="left"><font face="Arial" color="#000000">This is some fun code I
sometimes use in applications. It creates a cool effect.&nbsp;</font></p>
<p align="left"><i><font face="Arial">char.Play "LookDownBlink" '<font color="#008000">Looks
down and blinks</font><br>
char.Play "LookDownBlink" '<font color="#008000">Looks down and blinks</font><br>
char.Play "LookDownBlink" <font color="#008000">'Looks down and blinks</font><br>
char.Play "LookDownReturn" <font color="#008000">'Stops looking down</font><br>
char.Stop <font color="#008000"> 'Stops what he is doing</font><br>
char.MoveTo 300, 700 <font color="#008000"> 'Moves him to co ordinates 300,700
(off the screen!)</font><br>
char.Speak "Man It's really dark ..inside your monitor!" <font color="#008000">'Speaks</font>&nbsp;</font></i>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<i><font face="Arial">char.MoveTo 300, 50 <font color="#008000">'Move him to co
ordinates 300,50</font><br>
char.Speak "Nice to be back!"&nbsp; <font color="#008000">'Speaks</font><br>
</font></i></p>
<p align="left"><font face="Arial" color="#000080"><b>Character 'Move Around'
Code Example</b></font></p>
<p align="left"><font face="Arial" color="#000000">This is some fun code I
sometimes use in applications. It looks really funny on Peedy! Note - you may
have to change the screen co ordinates to suite your resolution.</font></p>
<p align="left"><i><font face="Arial">char.MoveTo 2000, 300 <font color="#008000"> 'Moves
him to co ordinates 2000,300 (off the screen!)</font><br>
char.MoveTo 300, 300 '<font color="#008000">Moves to co ordinates 300,300 (lower
middle of screen)</font><br>
char.Play "confused" '<font color="#008000">Looks Confused</font><br>
char.Speak "Nothing like a little flying to clear the head!" '<font color="#008000">Speaks</font><br>
char.Play "pleased" '<font color="#008000">Looks pleased</font><br>
</font></i></p>
<p align="left"><font face="Arial" color="#000080"><b>Character 'Open Notepad'
Code Example</b></font></p>
<p align="left"><font face="arial ">This code makes the character look like he
is writing in his notepad while you use your notepad.</font></p>
<p align="left"><font face="Arial"><i>char.MoveTo 50, 1 '<font color="#008000">Moves
character to upper left hand corner of the screen</font><br>
char.Speak "Let's use notepad!" '<font color="#008000">Speaks</font><br>
char.Play "Writing" <font color="#008000">'Character starts writing</font><br>
Shell &quot;Notepad.exe&quot;, vbNormalFocus <font color="#008000"> 'Opens Notepad
with Normal Focus<br>
</font></i></font></p>
<p align="left"><font face="Arial" color="#000080"><b>Character 'Grow' Code
Example</b></font></p>
<p align="left"><font face="Arial">This code makes the Character grow big! Looks
really cool (you tend to see the pixels though). You can customize the code to
make the character any size you want.</font></p>
<p align="left"><font face="Arial"><i>char.Height = &quot;750&quot; <font color="#008000">'Sets
the Characters Height</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Width = &quot;450&quot; <font color="#008000">'Sets
the Characters Width</font></i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Character 'Shrink' Code
Example</b></font></p>
<p align="left"><font face="Arial">This code makes the Character shrink! Looks
really cool (the animations don't look as good though). You can customize the
code to make the character any size you want.</font></p>
<p align="left"><font face="Arial"><i>char.Height = &quot;75&quot; <font color="#008000">'Sets
the Characters Height</font></i></font></p>
<p align="left"><font face="Arial"><i>char.Width = &quot;25&quot; <font color="#008000">'Sets
the Characters Width</font></i></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080"><b>Using an Input Box to let
the User specify what the Character Says</b></font></p>
<p align="left"><font face="Arial">This code is very useful because it lets the
user decide what the the character says.&nbsp;</font></p>
<p align="left"><font face="Arial"><i>Message = InputBox("What do you want Peedy to say?")
<font color="#008000">'Sets the Message String to equal the input box. Also sets
the input box's heading</font><br>
char.Speak Message <font color="#008000">'Speaks out the text in the Message
String</font><br>
</i></font></p>
<p align="left"><font face="Arial" color="#000080"><b>Using a Text Box to let
the User specify what the Character Says</b></font></p>
<p align="left"><font face="Arial">This code is useful to make the character
read a whole document. You can load text in to a text box and then tell the
character to read it. The following example requires a text box called Text1.</font></p>
<p align="left"><i><font face="Arial">if Text1.text &lt;&gt; &quot; &quot; then
char.speak text1.text <font color="#008000">'Checks to see if the text box is
empty. If it is not empty then it tells the character to speak the text.</font></font></i></p>
<p align="left"><i><font face="Arial">End if</font></i></p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Cool Web Links</font></p>
<p align="center"><font face="Arial">Here are a few URLs where you will find
information on Ms Agent related programs.</font></p>
<p align="center"><font face="Arial"><a href="http://msdn.microsoft.com/msagent">http://msdn.microsoft.com/msagent</a>
- The official Ms Agent site. Has developer downloads and the official developer
documents.</font></p>
<p align="center"><font face="Arial"><a href="http://agentry.net">http://agentry.net</a>
- Probably the biggest site on Ms Agent (apart from MSDN). Has over 300
characters, and a few are even free for download. A must see site!</font></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial"><a href="http://www.msagentring.org/">http://www.msagentring.org/</a>
- A collection of the best Ms Agent sites on the web. You can practically find almost
anything on Ms Agent here.</font></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial"><a href="http://members.theglobe.com/costas5">http://members.theglobe.com/costas5</a>
- Has some cool stuff including how to use Ms Agent in Word 97.</font></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial">Author's Note - I am not responsible for
content you find on these sites. Also if there are any cool resource sites (that
have source code or other stuff for developers), just email me and I'll add them
here in the next update.</font></p>
<hr>
<p align="center">&nbsp;</p>
<p align="center"><font face="Arial" size="4" color="#000080">Frequently Asked
Questions</font></p>
<p align="left"><font face="Arial" color="#000080">How do I know if I have a
Microsoft Agent Character file(s) on my computer?</font></p>
<p align="left"><font face="Arial" color="#000000">Just goto Start &gt; Find
&gt; Files or Folders and search for the extension *.acs . If you find any
such&nbsp; files in your <a href="file:///C:/Windows/MsAgent/Chars/">C:\Windows\MsAgent\Chars\</a>
folder then you are luck. If you have a file called Peedy.acs then this tutorial
will work. Otherwise just specify Anim = &quot;Your Character's Name).</font></p>
<p align="center">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080">Hey I'm too lazy to go
sifting through all that... is there some way I can do it through code?</font></p>
<p align="left"><font face="Arial" color="#000000"><i>Yes there is a way.. just
add this code to a form that has a agent control on it called Agent 1. </i> This code
will show a box which has all the character files installed on your computer.
Look through that and you will know if you have character files or not. Here is
the code&nbsp;</font></p>
<p align="left"><font face="Arial"><font color="#000000">Agent1.</font>ShowDefaultCharacterProperties</font></p>
<p align="center">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080">I don't have the file(s).
Where can I download them from? Are they freeware?</font></p>
<p align="left"><font face="Arial">The agent files can be freely downloaded, but
you are never the less bound by the Microsoft EULA (End User License Agreement).
For more information go to the URL specified below. The agent files (inlcuding the character
files) are available for download on <a href="http://msdn.microsoft.com/msagent">http://msdn.microsoft.com/msagent</a>
. You can also find custom animations created by various people at <a href="http://agentry.net">http://agentry.net</a></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080">How big are the character
files?</font></p>
<p align="left"><font face="Arial">The character files at MSDN range from 1.6 MB
to around 2 MB so they will take some time to download (depending on your
connection speed).</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080">Why don't some functions
(commands) work on some character files?</font></p>
<p align="left"><font face="Arial">Some versions of character files will
have more functions, so in order use
all the functions you may need to get a new character file. For example the char.play
&quot;Idle3_3&quot; function does not work on Robby.</font></p>
<p align="center">&nbsp;</p>
<p align="left"><font face="Arial" color="#000080">Sometimes the character
doesn't stop what he is doing for a long time... how can I force him to stop?</font></p>
<p align="left"><font face="Arial">Some functions take a long time to finish or
may even loop for ever so
you may have to force a stop. Just add the char.Stop or the char.StopAll
function to an event to stop the character. When this function is called the
character will automatically stop doing what he was doing and go to his rest
pose.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">Can I use the Ms Agent freely
in my
applications?</font></p>
<p align="left"><font face="Arial">Yes! as far as I know Microsoft is
distributing this across the internet. You can use the control in your apps but
please check out the licensing information first <span style="font-size: 12.0pt; mso-fareast-font-family: Times New Roman; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA"><a href="http://www.microsoft.com/workshop/imedia/agent/licensing.asp">http://www.microsoft.com/workshop/imedia/agent/licensing.asp</a></span></font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">How do I distribute Ms Agent
with my apps?</font></p>
<p align="left"><font face="Arial">You need to get the Cabinet (*.cab) files
from the MSDN site. Then you can include a reference to it in your installation
program. In order to do this too you need to agree with Microsoft's licensing
information (see above).</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">How can I change the
character file?</font></p>
<p align="left"><font face="Arial">In lots of examples I have seen, in order to
change the character file you need to change a lot of code. But if you used my
code you only have to change one line of code. All you have to do is to set the
Anim String to equal the character you want. For example to choose Peedy just
type the following code <i>Anim = &quot;Peedy&quot;</i>. Note that you can only
change the character if you have the character installed on your machine.</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">Can I use Ms Agent in VB 4.0?</font></p>
<p align="left"><font face="Arial">I have got reports that you can use Ms Agent
1.5 in Visual Basic 4. I am not sure if it will work in VB 4.0 (16 Bit), but it
should work in VB 4.0 (32 Bit).&nbsp;</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">Can I use Ms Agent in Java?</font></p>
<p align="left"><font face="Arial">As far as I know you can. I saw some Java
code on the MSDN site. You may want to check out the site (see below for URL).</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">Can I use Ms Agent in C and
C++?</font></p>
<p align="left"><font face="Arial">Yes, I think you can. There were some C++
examples on the MSDN site (I think). Check out the site - you may find some
sample code (URL below).</font></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">Where can I get more info on
Ms Agent?</font></p>
<p align="left"><span class="MsoHyperlink"><font face="Arial"><span style="font-size: 12.0pt; mso-fareast-font-family: Times New Roman; color: black; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA">Microsoft's
official Ms Agent developer page is at - <a href="http://msdn.microsoft.com/msagent">http://msdn.microsoft.com/msagent</a></span></font></span></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">What are some popular
commercial / shareware applications made with Ms Agent?</font></p>
<p align="left"><span class="MsoHyperlink"><font face="Arial"><span style="font-size: 12.0pt; mso-fareast-font-family: Times New Roman; color: black; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA">Well
the most famous app is probably Bonzi Buddy (<a href="http://www.bonzibuddy.com">www.bonzibuddy.com</a>).
Although this app initially used Peedy, I think they have now developed their
own character(s).</span></font></span></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">I can't understand a part (or
part's) of this tutorial. Can you help?</font></p>
<p align="left"><span class="MsoHyperlink"><font face="Arial"><span style="font-size: 12.0pt; mso-fareast-font-family: Times New Roman; color: black; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA">Of
course! Just email me (address below)! I will be happy to help in anyway I can.</span></font></span></p>
<p align="left">&nbsp;</p>
<p align="left"><font color="#000080" face="Arial">How can I make sure that I
will get to see more tutorials like this?&nbsp;</font></p>
<p align="left"><span class="MsoHyperlink"><font face="Arial"><span style="font-size: 12.0pt; mso-fareast-font-family: Times New Roman; color: black; mso-ansi-language: EN-US; mso-fareast-language: EN-US; mso-bidi-language: AR-SA">I
am greatly encouraged by your comments, suggestions and especially your votes.
Your support will help me to write more tutorials like this one.</span></font></span></p>
<p align="left">&nbsp;</p>
<p align="center"><b><font face="Arial" color="#000080">THE END</font></b></p>
<p align="center"><font face="Arial" color="#000000">A <b>lot</b> of hard work
has gone into this tutorial. I have spent <b>many</b> hours writing this article
in an easy to understand manner. If you like this please <b>vote</b> for me.
Also feel free to post any <b>comments</b> or <b>suggestions</b> as to what I
can include in the next version. Feel free to mail me at <a href="mailto:vbdude777@email.com">vbdude777@email.com</a>
and also check out my website at <a href="http://mahangu.tripod.com">http://mahangu.tripod.com</a></font></p>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>

