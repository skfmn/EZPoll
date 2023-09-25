**************************************
* EZPoll v4 
* copyright 2007 - 2023 by Steve Frazier                     
*                            
* www.aspjunction.com                               
*                                                   
* You may not sell this script                      
* You may modify it and distribute it free as       
* long as this readme.txt with this copyright       
* header remains with it.                           
**************************************


*******************Installation Instructions*******************************

Create an MSSQL Database for your Poll, if your not sure how contact your hosting provider

Upload the poll folder to your server and navigate to:

/poll/install/install.asp

Then follow the instructions

You will be able to login at:

/poll/admin/admin_login.asp

Login using "admin" as your login name AND password. 
Once you have logged in you can and should change your login info.
	
<!-- #include virtual="/poll/poll.asp"-->
		 
************Customization Hint***************************************

You can wrap the code above in a span or div tag and use CSS to define the objects in the code.

  EXAMPLE:
	
<div id="poll_wrapper">
  <!-- #include virtual="/poll/poll.asp"-->
</div>

<style type="text/css">
  @media screen and (min-width: 980px) {
    div#main {margin-left: -310px;}
    #poll_wrapper {font-size:12px;color:#BBBBBB;}
  }
  @media screen and (max-width: 980px) {
    div#main {margin-left: -100px;}
    #poll_wrapper {width:750px;font-size:12px;color:#BBBBBB;}
  }
  #poll_wrapper a {color:#AAAAAA;}
	#poll_wrapper a.button {color:#FFFFFF}
</style>

You can define most attributes like fonts, alignments and such!

You can also display it in an iframe:

<iframe src="/poll/demo.asp" height="600" width="900"></iframe>

I can install any of the EZCodes for a fee.
If you would like some custom ASP coding done I am available and I charge reasonable rates

Please contact admin@aspjunction.com


Change Log:

V 1.0 - Basic version with only one poll and little else.
        .1 - .5 - Buch of stuff was changed including a new look and multiple Polls.
	
V 2.0 - New look
        .1 Implemented jQuery tabs and Fancybox plugin.
				
V 3.0 - Completely re-written code
        .1 New look.
				.2 Modified to allow voters to change vote.
				.3 Added ability to add more options on the edit page.

V 4.0 - Completely re-written code...again
        New look.
		Added Admins page.
		Added settings page. 
