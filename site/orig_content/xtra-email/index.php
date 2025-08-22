<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Excel VBA Primer</title>
<link href="../../nljstyles.css" rel="stylesheet" type="text/css" />
<link href="../primer.css" rel="stylesheet" />
<link href="../../prism/prism.css" rel="stylesheet" />

</head>

<body>
<script src="../../prism/prism.js"></script>

<?php 
require "../header.php";
?>

<h1>Sending E-Mail Using VBA in Excel</h1>
<p>It is often useful to send e-mail using VB code. You can also send text messages using this technique. This page includes sample code and instructions.</p>
<h2>Adding the Email Utilities Module</h2>
<p>First of all, you need to add some code to your project. Do the following:</p>
<p>Unzip the file. Some e-mail clients reject  attachments containing code, so I put it in a zip archive. In the archive you  will see a file called <strong>email_utilities.bas</strong>.</p>
<ol>
  <li>Right-click here (<a href="email_utilities.bas">email_utilities.bas</a>) to download a file containing VB code for sending e-mail.</li>
  <li>Go to your VB Editor window and right-click in the  Project Explorer window on the left and select <strong>Import File</strong>.</li>
  <li>Path to the <strong>email_utilities.bas</strong> file and open it. It will then be added as a new module. </li>
  <li>Open the module and look at the functions and  subs. You can now call these from anywhere in your code.</li>
</ol>
<p>So to use the  <strong>send_mail</strong> sub, you just need to call it and pass the arguments. Like this:</p>

<pre><code class="language-vb">send_mail "myaddress@gmail.com", "TEST", "Hello world."
</code></pre>


<p>This will create a new message and launch your default e-mail client (Outlook, Thunderbird, etc). If you are working in the CAEDM network, there may not be an e-mail client set up on your account so the first time you run this you will probably need to go through a setup process. If you want to be able to send mail without having to hit the Send button on each message, you will need to use the Outlook version of the send_mail function.</p>
<h2>Sending Text Messages</h2>
<p>Next you need to decide if you want to send e-mail or text  messages. To send a text message you use the same code, but you use the phone  number and carrier to formulate an e-mail address for the text message. First  you have to determine the correct e-mail suffix from the following list:</p>
<blockquote>
<table border="0" cellspacing="0" cellpadding="0" width="300">
  <tr>
    <td width="87" nowrap="nowrap" valign="bottom"><strong>Company</strong></td>
    <td width="213" nowrap="nowrap" valign="bottom"><strong>E-mail tag</strong></td>
  </tr>
  <tr>
    <td width="87" nowrap="nowrap" valign="bottom">Sprint</td>
    <td width="213" nowrap="nowrap" valign="bottom">messaging.sprintpcs.com</td>
  </tr>
  <tr>
    <td width="87" nowrap="nowrap" valign="bottom">AT&amp;T</td>
    <td width="213" nowrap="nowrap" valign="bottom">mmode.com</td>
  </tr>
  <tr>
    <td width="87" nowrap="nowrap" valign="bottom">Cingular</td>
    <td width="213" nowrap="nowrap" valign="bottom">mobile.mycingular.com</td>
  </tr>
  <tr>
    <td width="87" nowrap="nowrap" valign="bottom">Nextel</td>
    <td width="213" nowrap="nowrap" valign="bottom">messaging.nextel.com</td>
  </tr>
  <tr>
    <td width="87" nowrap="nowrap" valign="bottom">T-Mobile</td>
    <td width="213" nowrap="nowrap" valign="bottom">tmomail.net</td>
  </tr>
  <tr>
    <td width="87" nowrap="nowrap" valign="bottom">Verizon</td>
    <td width="213" nowrap="nowrap" valign="bottom">vtext.com</td>
  </tr>
</table>
</blockquote>
<p>You can look up more complete lists on the internet. Then you take the phone number (&quot;111-222-3333&quot;) and remove the dashes so that it is nothing but numbers (&quot;1112223333&quot;) using the <strong>Replace</strong> string function. This creates the prefix. Then you combine the prefix and the suffix to generate the e-mail addresses as follows:</p>
<blockquote><table cellspacing="0" cellpadding="0">
  <col width="101" />
  <col width="63" />
  <col width="226" />
  <tr>
    <td width="101"><strong>Mobile #</strong></td>
    <td width="63"><strong>Carrier</strong></td>
    <td width="226"><strong>E-mail</strong></td>
  </tr>
  <tr>
    <td>801-111-1111</td>
    <td>Sprint</td>
    <td>8011111111@messaging.nextel.com</td>
  </tr>
  <tr>
    <td>801-222-2222</td>
    <td>Nextel</td>
    <td>8012222222@messaging.nextel.com</td>
  </tr>
  <tr>
    <td>801-333-3333</td>
    <td>Cingular</td>
    <td>8013333333@mobile.mycingular.com</td>
  </tr>
  <tr>
    <td>801-444-4444</td>
    <td>Nextel</td>
    <td>8014444444@messaging.nextel.com</td>
  </tr>
  <tr>
    <td>801-555-5555</td>
    <td>Nextel</td>
    <td>8015555555@messaging.nextel.com</td>
  </tr>
  <tr>
    <td>801-666-6666</td>
    <td>AT&amp;T</td>
    <td>8016666666@mmode.com</td>
  </tr>
  <tr>
    <td>801-777-7777</td>
    <td>Verizon</td>
    <td>8017777777@vtext.com</td>
  </tr>
  <tr>
    <td>801-888-8888</td>
    <td>AT&amp;T</td>
    <td>8018888888@mmode.com</td>
  </tr>
  <tr>
    <td>801-310-9291</td>
    <td>T-Mobile</td>
    <td>8013109291@tmomail.net</td>
  </tr>
  <tr>
    <td>801-209-5114</td>
    <td>AT&amp;T</td>
    <td>8012095114@mmode.com</td>
  </tr>
</table></blockquote>
<p>Sometimes it is easiest to formulate the addresses just using Excel formulas. You can use a VLOOKUP to get the proper suffix from the carrier. Then you use your VB code to loop through the table and send your e-mail messages.</p>
<h2>Other Resources</h2>
<p>For more information on sending email using VBA, see the following:</p>
<p><a href="http://www.rondebruin.nl/sendmail.htm">http://www.rondebruin.nl/sendmail.htm</a> (great resource here - extensive set of sample code)</p>
<p><a href="http://msdn.microsoft.com/en-us/library/office/ff458119%28v=office.11%29.aspx">http://msdn.microsoft.com/en-us/library/office/ff458119%28v=office.11%29.aspx</a> (Outlook sample code)</p>
<p><a href="http://msdn.microsoft.com/en-us/library/office/ff519602%28v=office.11%29.aspx">http://msdn.microsoft.com/en-us/library/office/ff519602%28v=office.11%29.aspx</a> (Part 2 from previous link)</p>
<p><a href="http://www.makeuseof.com/tag/send-emails-excel-vba/">http://www.makeuseof.com/tag/send-emails-excel-vba/</a> (The CDO method - I haven't tried this, but it looks pretty slick)</p>
<?php 
require "../footer.php";
?>
</body>
</html>
