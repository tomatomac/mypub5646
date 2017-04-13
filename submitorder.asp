<!--#Include file="Include/incFunctions.asp" -->
<!--#Include file="Include/conn.asp" -->
<%
'strLoginName=trim(sqlEncode("txtLogin"))
'strpass=trim(sqlEncode("txtPass"))
strfName=trim(SqlEnCode("strfirstname"))
strlName=trim(SqlEnCode("strLastname"))
strAdd=trim(SqlEnCode("txtAdd"))
strCountry=trim(SqlEnCode("strCountry"))
strCity=trim(SqlEnCode("strCity"))
strZip=trim(SqlEnCode("strZip"))
strState=trim(SqlEnCode("strState"))
strPhone=trim(SqlEnCode("strPhone"))
strEmail=trim(SqlEnCode("txtemail"))
strFax=trim(SqlEnCode("txtFax"))
if request("paymode1")<>"" then
paymode=request("paymode1")
end if
if request("paymode2")<>"" then
paymode=request("paymode2")
end if
if request("paymode3")<>"" then
paymode=request("paymode3")
end if
if request("paymode4")<>"" then
paymode=request("paymode4")
end if

strSfName=trim(SqlEnCode("strSfirstname"))
strSlName=trim(SqlEnCode("strSLastname"))
strSAdd=trim(SqlEnCode("txtSAdd"))
strSCity=trim(SqlEnCode("strSCity"))
strSCountry=trim(SqlEnCode("strSCountry"))
strSstate=trim(SqlEnCode("strSState"))
strSZip=trim(SqlEnCode("strSZip"))
strSPhone=trim(SqlEnCode("strSPhone"))
strSEmail=trim(SqlEnCode("strSEmail"))
strName=strfName&" " &strlName
If request("flag")=1 then
set rst=server.CreateObject("adodb.recordset")
sql="select * from user_details where Id="&request("userid")
rst.open sql,con
If not rst.EOF then
	'if request("txtLoginId")=trim(rst("User_Id")) and request("txtLoginPass")=trim(rst("pass")) then
	UId=rst("Id")
	'con.Execute("Update user_Details set User_Id='" & strLoginName & "',pass='" & strpass & "',fName='" & strfirstName & "',lName='" & strlastName & "',Address='"& strAdd &"',City='"&strCity&"',ZipCode='"&strZip&"',State='"&strState&"',Country='"&strCountry&"',Phone='"&strPhone&"',Fax='"&strFax&"',Email='"&strUserEmail&"',website='"& strUserWebsite &"',currency_pref='"&strCurrency&"'  where ID=" & UId)
	con.Execute("Update user_Details set fName='" & strfName & "',lName='" & strlName & "',Address='"& strAdd &"',City='"&strCity&"',ZipCode='"&strZip&"',State='"&strState&"',Country='"&strCountry&"',Phone='"&strPhone&"',Fax='"&strFax&"',Email='"&strEmail&"',SfName='" & strSfName & "',SlName='" & strSlName & "',SAddress='"& strSAdd &"',SCity='"&strSCity&"',SZipCode='"&strSZip&"',SState='"&strSState&"',SCountry='"&strSCountry&"',SPhone='"&strSPhone&"',SEmail='"&strSEmail&"',GFName='" & strGfName & "',GlName='" & strGlName & "',GEmail='"&strGEmail&"',GMessage='"&strGMsg&"' where ID=" & UId)
		strmsg="Login Information Updated Succesfully"
		set rst=nothing
		rst.close
	'else
	'	strmsg="Login Id is allready exist. Please use anothor Login id."
'		response.Redirect("register.asp?errmsg="&strmsg)
end if
else
		con.Execute("Insert into user_Details(fName,lName,Address,City,ZipCode,State,Country,Phone,Email,Fax,SfName,SlName,SAddress,SCity,SZipCode,SState,SCountry,SPhone,SEmail,GFName,GLName,GEmail,GMessage) values('" & strfName & "','"&strlName&"','" & strAdd & "','" & strCity & "','" & strzip & "','" & strState & "','" & strCountry & "','" & strPhone & "','" & strEmail & "','" & strFax & "','" & strSfName & "','"&strSlName&"','" & strSAdd & "','" & strSCity & "','" & strSzip & "','" & strSState & "','" & strSCountry & "','" & strSPhone & "','" & strSEmail& "','" & strGFName & "','" & strGlName & "','"&strGEmail&"','"&strGMsg&"')")
		strmsg="Inserted Succesfully"
		sql="select id from user_details"
		set rst1=server.CreateObject("adodb.recordset")
		rst1.open sql,con,3,3
		rst1.movelast
		UId=rst1("Id")
		'set rst1=nothing
		'rst1.close
end if

''******************Order Details***********************************
scartitemnums=session("cartitemnumarray")
scartitemname=session("cartitemnamearray")
scartItemPrice=session("cartitempricearray")
ncartitemqtys=session("cartitemqtyarray")
'scartitemimage=session("cartitemimagearray")
'scartitemcurr=session("cartitemcurrarray")
scartitemisbn=session("cartitemisbn")
nnumcartitems=ubound(scartitemnums)

if nnumcartitems > 0 then
'response.Write("enter")
set rsorder1 = server.createobject("adodb.recordset")
rsorder1.Open "select max(order_id) as m from order_details",con,3,3
if not rsorder1.eof then
if rsorder1("m")>0 then
oNo=rsorder1("m")
else
oNo=0
end if
end if
session("ono")=ono+1
rsorder1.close
set rsorderdetail = server.CreateObject("adodb.recordset")
rsorderdetail.open "select * from order_details",con,3,3
for i=1 to nnumcartitems
amt=mid(formatcurrency(ncartitemqtys(i)*scartItemPrice(i),2),2)
if scartitemname(i)<>"" and scartitemnums(i)<>"" then
rsorderdetail.addnew
rsorderdetail("user_id")=strName
rsorderdetail("order_id")=session("ono")
'rsorderdetail("item_no")=scartitemnums(i)
rsorderdetail("Product_code")=scartitemisbn(i)
rsorderdetail("item_name")=scartitemname(i)
rsorderdetail("Qty")=ncartitemqtys(i)
rsorderdetail("ShipCharges")=session("ShipCharge")
'rsorderdetail("Extra_Charges")=session("extra")
rsorderdetail("unit_price")=scartItemPrice(i)
rsorderdetail("amount")=amt
rsorderdetail("Grand_Total")=session("GR")
rsorderdetail("Payment_Mode")=paymode
rsorderdetail.update
end if
next

end if
''*****************Email to Admin***********************
strhtml=strhtml &"<table width='60%' border=0 cellpadding=1 cellspacing=1 align=center>"
strhtml=strhtml & "<th colspan=2 bgcolor='#000099'><font  color='#FFFFFF'>User Billing Details</font></th>"
'strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><strong><font color=19398c> Login Id (Email)</font></strong></td>"
'strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strLoginName&"</td></tr>" 
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><strong><font color=19398c> Order Id</font></strong></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&session("ono")&"</td></tr>" 
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><strong><font color=19398c> First Name</font></strong></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strfName&"</font></td> </tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><font color='19398c'><b>Last Name</b> </font></td>" 
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strlName&"</font></td></tr>" 
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td valign='top' bgcolor='#ECF9F9'><font color='19398c'><b>Address</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strAdd&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' ><font color='19398c'><b>City</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strCity&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' ><font color='19398c'><b>State</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strState&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' ><font color='19398c'><b>Zip</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strZip&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><font color='19398c'><b>Phone</b></font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strPhone&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' ><font color='19398c'><b>Fax</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strFax&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><font color='19398c'> <b>Email</b></font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strEmail&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><font color='19398c'><b>Country</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strCountry&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor=#ECF9F9 ><font color='19398c'><b>Payment Mode</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&paymode&"</font></td></tr>"
strhtml=strhtml & "</table><br><br>"
strhtml=strhtml &"<table width=50% border=0 cellpadding=1 cellspacing=1 align=center>"
strhtml=strhtml & "<th colspan=2 bgcolor='#000099'><font  color='#FFFFFF'>User Shipping Details</font></th>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><strong><font color=19398c> First Name</font></strong></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSfName&"</font></td> </tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><font color='19398c'><b>Last Name</b> </font></td>" 
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSlName&"</font></td></tr>" 
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td valign='top' bgcolor='#ECF9F9'><font color='19398c'><b>Address</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSAdd&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' ><font color='19398c'><b>City</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSCity&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor=#ECF9F9 ><font color='19398c'><b>State</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSState&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor=#ECF9F9 ><font color='19398c'><b>Zip</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSZip&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor=#ECF9F9 ><font color='19398c'><b>Phone</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSPhone&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><font color='19398c'> <b>Email</b></font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSEmail&"</font></td></tr>"
strhtml=strhtml & "<tr bgcolor='#ECF9F9'><td bgcolor='#ECF9F9' width='25%'><font color='19398c'><b>Country</b> </font></td>"
strhtml=strhtml & "<td><font face='Tahoma, Verdana' size=2 color='#000000'>"&strSCountry&"</font></td></tr>"
strhtml=strhtml & "</table><br><br>"


strhtml1=strhtml1 &"<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1'  bordercolor='#996600' bgcolor='#336699'>"
strhtml1=strhtml1 & "<th colspan=5 bgcolor='#336699'><font face='Tahoma, Verdana' size=2 color='#FFFFFF'>Order Details</font></th>"
'strhtml1=strhtml1 & "<tr>" 
strhtml1=strhtml1 & "<tr ><td width='15%' bgcolor='#336699' align='center'><font face='Tahoma, Verdana' size=2 color='#FFFFFF'><b>Order No </b></font></td>"
strhtml1=strhtml1 & "<td width='40%' align='left' bgcolor='#ffffff' colspan=4><font face='Tahoma, Verdana' size=2 color='black'>&nbsp;<b>"&session("ono")&"</b></td></tr>"
strhtml1=strhtml1 & "<tr ><td width='18%' align='center' bgcolor='CDD6E5'><font face='Tahoma, Verdana' size=2 color='#336699'><b>ISBN </b></font></td>"
strhtml1=strhtml1 & "<td width='49%' align='center' bgcolor='CDD6E5'><font face='Tahoma, Verdana' size=2 color='#336699'><b>Description </b></font></td>"
strhtml1=strhtml1 & "<td width='5%' align='center'  bgcolor='CDD6E5'><font face='Tahoma, Verdana' size=2 color='#336699'><b>Qty </b></font></td>"
strhtml1=strhtml1 & "<td width='12%' align='center'  bgcolor='CDD6E5'><font face='Tahoma, Verdana' size=2 color='#336699'><b>Price($) </b></font></td>"
strhtml1=strhtml1 & "<td width='16%' align='center' bgcolor='CDD6E5'><font face='Tahoma, Verdana' size=2 color='#336699'><b>Amount($)</b></font></td>"
strhtml1=strhtml1 & "</tr>"
         
for i=1 to session("count")
Isbn=session("isbn"&i)
ItemName=session("Item"&i)
qty=session("qty"&i)
price=session("price"&i)
amt=session("amt"&i)
strhtml1=strhtml1 & "<tr>" 
strhtml1=strhtml1 & "<td width='18%' align='left' bgcolor='#ffffff'><font face='Tahoma, Verdana' size=2 color='black'>"&Isbn&"</td>"
strhtml1=strhtml1 & "<td width='49%' align='left' bgcolor='#ffffff'><font face='Tahoma, Verdana'  size=2 color='black'>"&ItemName&"</td>"
strhtml1=strhtml1 & "<td width='5%' align='center' bgcolor='#ffffff'><font face='Tahoma, Verdana' size=2 color='black'>"&qty&"</td>"
strhtml1=strhtml1 & "<td width='12%' align='center' bgcolor='#ffffff'><font face='Tahoma, Verdana' size=2 color='black'>"&price&"</td>"
strhtml1=strhtml1 & "<td width='16%' align='center' bgcolor='#ffffff'><font face='Tahoma, Verdana' size=2 color='black'>"&amt&"</td>"
strhtml1=strhtml1 & "</tr>"
next
'strhtml1=strhtml1 & "<tr bgcolor='#ECF9F9'>"
'strhtml1=strhtml1 & "<td colspan='4' align=right ><b>Sub Total</b></td>"
'strhtml1=strhtml1 & "<td width='50%' align='center'><b>£ "&session("ST")&"</b></td>"
'strhtml1=strhtml1 & "</tr>"
strhtml1=strhtml1 & "<tr bgcolor='#ffffff'>"
strhtml1=strhtml1 & "<td colspan='4' align=right ><font face='Tahoma, Verdana' size=2 color='black'><b>Shipping Charges</b></font></td>"
strhtml1=strhtml1 & "<td width='50%' align='center'><font face='Tahoma, Verdana' size=2 color='black'><b>$ "&mid(formatcurrency(session("ShipCharge"),2),2)&"</b></td>"
strhtml1=strhtml1 & "</tr>"
'strhtml1=strhtml1 & "<tr bgcolor='#ffffff'>"
'strhtml1=strhtml1 & "<td colspan='4' align=right ><font face='Tahoma, Verdana' size=2 ><b>Extra Charges</b></td>"
'strhtml1=strhtml1 & "<td width='50%' align='center'><font face='Tahoma, Verdana' size=2><b>$  "&session("extra")&"</b></td>"
'strhtml1=strhtml1 & "</tr>"
strhtml1=strhtml1 & "<tr bgcolor='#ffffff'><td colspan=2>&nbsp;</td>"
strhtml1=strhtml1 & "<td colspan='2' align=right bgcolor='#336699'><font face='Tahoma, Verdana' size=2 color='white'><b>Grand Total</b></td>"
strhtml1=strhtml1 & "<td width='50%' align='center' bgcolor='#336699'><font face='Tahoma, Verdana' size=2 color='white'><b>$  "&session("GR")&"</b></td>"
strhtml1=strhtml1 & "</tr>"

strhtml1=strhtml1 & "</table>"

strhtml=strhtml&strhtml1
'response.Write(strhtml) 

''************Email to Administrator**************




'Set myMail=CreateObject("CDO.Message")
'myMail.Subject="Order details from Fatherson.com website"
'myMail.From=strEmail
'myMail.To="lance@fatherson.com"
'myMail.HTMLBody = strhtml
'myMail.Send
'set myMail=nothing











'Set Mailer= server.CreateObject("CDONTS.NewMail") 
'Dim iMsg, iConf
'Set iMsg = Server.CreateObject("CDO.Message")
'Set iConf = iMsg.Configuration
'With iConf.Fields
 '    .Item(cdoSendUsingMethod) = cdoSendUsingPort    ' http://msdn.microsoft.com/en-us/library/ms527265.aspx
  '   .Item(cdoSMTPServer)      = "smtp.mydomain.com" ' http://msdn.microsoft.com/en-us/library/ms527294.aspx
   '  .Update
'End With

'With iMsg
    ' iMsg.To       = "itbyteservice1@gmail.com"
     'iMsg.From     = """Fatherson Website"" &strEmail"
     'iMsg.Subject  = "Order detail from Fatherson.com website"
     'iMsg.TextBody = strhtml
     'iMsg.Send
'End With
'Set iMsg = Nothing
'Set iConf = Nothing


'set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'Mailer.From   = strEmail
'Mailer.To ="lance@fatherson.com"
'Mailer.BodyFormat = 0
'Mailer.MailFormat = 0
'Mailer.Subject    = "Order detail at Fatherson.com"
'Mailer.Body=strhtml
'Mailer.Send()
'set Mailer=nothing


'****Email to user************'


'NB -- removed'
'Set myMail=CreateObject("CDO.Message")
'myMail.Subject="Order details from Fatherson.com website"
'myMail.From="lance@fatherson.com"
'myMail.To=strEmail
'myMail.HTMLBody = strhtml1
'myMail.Send
'set myMail=nothing







'Set Mailer= server.CreateObject("CDONTS.NewMail") 
'set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'Mailer.From   = "lance@fatherson.com"
'Mailer.To =stremail
'Mailer.BodyFormat = 0
'Mailer.MailFormat = 0
'Mailer.Subject    = "Order detail from Fatherson.com"
'Mailer.Body=strhtml
'Mailer.Send()
'set Mailer=nothing

%>

	<%
if paymode="Credit Card" then

	'*******************************session end*********************]
dim sItemNums(0),nItemQty(0),sItemPrice(0),sitemName(0),sitemimage(0),sitemweight(0),sitemisbn(0)
sItemNums(0)=0
sitemName(0)=""
nitemqty(0)=0
sItemPrice(0)=0
sitemimage(0)=""
sitemweight(0)=""
sitemisbn(0)=""
session("cartItemNumArray")=sItemNums
session("cartitemnamearray")=sitemName
session("cartItemQtyArray")=nitemqty
session("cartItemImageArray")=sitemimage
session("cartItemPriceArray")=sitemprice
session("cartitemweightarray")=sitemweight
session("cartItemisbn")=sitemisbn
session("casevar")=0
session("salecomplete")=false
session("frompage")=""
session("frompage1")=""
session("newsLetterType")=""


end if
%>


<HTML><HEAD><TITLE>Welcome to Father Son Publishing inc.</TITLE>
<META http-equiv=Content-Type content="text/html; charset=windows-1252">
<SCRIPT language=JavaScript>
<!--
function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function fwLoadMenus() {
  if (window.fw_menu_0) return;
  window.fw_menu_0 = new Menu("root",109,18,"Verdana, Arial, Helvetica, sans-serif",11,"#000066","#ffffff","#7b96cd","#000066");
  fw_menu_0.addMenuItem("Laptops");
  fw_menu_0.addMenuItem("Desktops");
  fw_menu_0.addMenuItem("Accessories");
  fw_menu_0.addMenuItem("CD-RW Drives");
  fw_menu_0.addMenuItem("Handhelds");
  fw_menu_0.addMenuItem("Hard Drives");
  fw_menu_0.hideOnMouseOut=true;

  fw_menu_0.writeMenus();
} // fwLoadMenus()

//-->
</SCRIPT>

<SCRIPT language=JavaScript1.2 src="mainimages/fw_menu.js"></SCRIPT>
<LINK href="mainimages/styles.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 
onload="MM_preloadImages('images/home_f2.gif','images/top_sellers_f2.gif','images/products_f2.gif','images/shopping_cart_f2.gif','images/contact_f2.gif','images/laptops_f2.gif','images/desktops_f2.gif','images/accessories_f2.gif','images/handhelds_f2.gif','images/hard_drives_f2.gif','images/keyboards_f2.gif','images/monitors_f2.gif','images/printers_f2.gif','images/scanners_f2.gif','images/sound_cards_f2.gif','images/software_f2.gif')" 
marginwidth="0" marginheight="0">
<SCRIPT language=JavaScript1.2>fwLoadMenus();</SCRIPT>


<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><!--#include file=include/header.asp --></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="59" valign="top"> 
      <!--#include file=include/leftbar.asp-->
      <img height=300 src="mainimages/right_fill.gif" 
      width=176 border=0 name=right_fill> </td>
    <td width="624" valign="top"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="12" align="left" class="title1"><strong><font color="#990000" size="4">Registration 
            Details </font></strong></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><br> <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="redbig">
              <tr> 
                <td colspan=3  class="font_heading"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="18" colspan="3" align=right> &nbsp;<A href=submitorderprint.asp><font size=2></font></a></td>
                    </tr>
                    <tr> 
                      <td width="3%" height="9" bgcolor="#336699">&nbsp;</td>
                      <td width="50%" bgcolor="#336699" class="subtitle"><font color="#FFFFFF"><strong>Payment 
                        Details</strong></font></td>
                      <td width="47%" height="9" bgcolor="#336699" align=right class="font_heading"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td  colspan="2"  class="subtitle"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Payment 
                  Type :</strong></font> <span class="title"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=paymode %></font> 
                  </span> <div align="right"></div></td>
                <td width="46%"  class="font_text">&nbsp; </td>
              </tr>
              <tr> 
                <td colspan=3 > <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr bgcolor="#336699"> 
                      <td width="3%" height="18">&nbsp;</td>
                      <td width="97%" class="subtitle"><font color="#FFFFFF"><strong>Billing 
                        Details</strong></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr > 
                <td width="48%" valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Your 
                    Name :<br>
                    </font></div></td>
                <td width="6%" valign="top" class="text"> <div align="left" class="title"> 
                    <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strfName%>&nbsp;<%=strlName%> 
                      </font></p>
                  </div></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
                    Billing Address :</font></div></td>
                <td valign="top" class="title"> <div align="left" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%=strAdd%> </font></div></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">City,State,Zip 
                    : </font></div></td>
                <td valign="top" class="title"> <div align="left" class="cell"> 
                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strcity%>&nbsp;<%=strState%>&nbsp;<%=strzip%> 
                    </font></div></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Country 
                    :</font></div></td>
                <td valign="top" class="title"> <div align="left"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strCountry%> 
                    </font></div></td>
              </tr>
              <tr class=font_text> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Telephone 
                    Number :<br>
                    </font></div></td>
                <td valign="top" class="title"> <div align="left"  class="cell"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strPhone%> 
                    </font></div></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Email 
                    Address :</font></div></td>
                <td valign="top" class="title"> <div align="left" class="cell"> 
                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=stremail%> 
                    </font></div></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Fax 
                    :</font></div></td>
                <td valign="top" class="title"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strFax%></font></td>
              </tr>
              <tr class="font_text"> 
                <td colspan=3  class="font_heading"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr bgcolor="#336699"> 
                      <td width="3%" height="18">&nbsp;</td>
                      <td width="97%" bgcolor="#336699" class="subtitle"><font color="#FFFFFF"><strong>Shipping 
                        Details</strong></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Your 
                    Name :<br>
                    </font></div></td>
                <td valign="top" class="title"> <div align="left"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strSfName%>&nbsp; 
                    <%=strSLName%> </font></div></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Delivery 
                    Address :</font></div></td>
                <td valign="top" class="title"> <div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%=strSAdd%> </font> </div></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">City,State,Zip 
                    : </font></div></td>
                <td valign="top" class="title"> <div align="left" class="cell"> 
                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strSCity%>&nbsp;&nbsp; 
                    <%=strSState%>&nbsp;&nbsp;<%=strSZip%> </font></div></td>
              </tr>
              <tr class="font_text"> 
                <td height="22" valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Country 
                    :</font></div></td>
                <td valign="top" class="title"> <div align="left"> 
                    <!--<input type="text" name="Delivery_Cust_Coutry" size="20" maxlength=60  value=""class="formborder">-->
                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%=strSCountry%> </font></div></td>
              </tr>
              <tr class=font_text> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Telephone 
                    Number&nbsp;: </font></div></td>
                <td valign="top" class="title"> <div align="left"  class="cell"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strSPhone%></font></div></td>
              </tr>
              <tr class="font_text"> 
                <td valign="top" class="subtitle"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Email 
                    Address :</font></div></td>
                <td valign="top" class="title"> <div align="left" class="cell"> 
                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=strSEmail%> 
                    </font></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <br> <table width="94%" border="0" align=center cellpadding="0" cellspacing="0">
        <tr> 
          <td><%=strhtml1%></td>
        </tr>
      </table>
      <table width="94%" border="0" align=center cellpadding="0" cellspacing="0">
        <tr class=cell> 
          <form name="frmpay" method="post" action="register.asp?flag=1">
            <input name="userid" type="hidden" value="<%=UId%>">
            <td colspan="2" align="center" valign="top"   class="cell"> <input name="oNO" type="hidden" value="<%=session("oNO")%>"> 
              <input name="sid" type="hidden" value="335672"> <input name="cart_order_id" type="hidden" value="<%=session("ono")%>"> 
              <input name="total" type="hidden" value="<%=request("GT")%>"> <input name="card_holder_name" type="hidden" value="<%=strfName%>&nbsp;<%=strlName%>"> 
              <input name="city" type="hidden" value="<%=strCity%>"> <input name="street_address" type="hidden" value="<%=strAdd%>"> 
              <input name="state" type="hidden" value="<%=strCity%>"> <input name="zip" type="hidden" value="<%=strZip%>"> 
              <input name="country" type="hidden" value="<%=strCountry%>"> <input name="phone" type="hidden" value="<%=strPhone%>"> 
              <input name="email" type="hidden" value="<%=stremail%>"> <input name="ship_name" type="hidden" value="<%=strSfName%>&nbsp;<%=strSlName%>"> 
              <input name="ship_city" type="hidden" value="<%=strSCity%>"> <input name="ship_street_address" type="hidden" value="<%=strSAdd%>"> 
              <input name="ship_state" type="hidden" value="<%=strSCity%>"> <input name="ship_zip" type="hidden" value="<%=strSZip%>"> 
              <input name="ship_country" type="hidden" value="<%=strSCountry%>"> 
              <input name="ship_Phone" type="hidden" value="<%=strSPhone%>"> <input name="ship_Email" type="hidden" value="<%=strSEmail%>"> 
              <input name="txtLoginId" type="hidden" value="<%=strLoginName%>"> 
              <input name="txtLoginPass" type="hidden" value="<%=strpass%>"> <input name="txtfName" type="hidden" value="<%=strfName%>"> 
              <input name="txtLName" type="hidden" value="<%=strlName%>"> <input name="txtSfName" type="hidden" value="<%=strSfName%>"> 
              <input name="txtSLName" type="hidden" value="<%=strSlName%>"> <input name="txtGfName" type="hidden" value="<%=strGfName%>"> 
              <input name="txtGLName" type="hidden" value="<%=strGlName%>"> <input name="txtGEmail" type="hidden" value="<%=strGEmail%>"> 
              <input name="txtGMsg" type="hidden" value="<%=strGMsg%>"> </td>
          </form>
          <td align="center" valign="top"  class="cell"> <div align="center"> 
              <% if payMode="Credit Card" then%>
              <form method="post" name="frmpay"  action="https://secure2.nettally.com/fatherson/creditcard.asp">
                <%else if payMode="Paypal" then%>
              </form>
              <form method="post" name="frmpay"  action="https://www.paypal.com/cgi-bin/webscr">
                <%else%>
              </form>
              <form method="post" name="frmpay"  action="thanks.asp">
                <%end if
				end if
				   %>
                <input type="hidden" name="fname" value="<%=strfName%>">
                <input type="hidden" name="lname" value="<%=strlName%>">
                <input name="cart_order_id" type="hidden" value="<%=session("ono")%>">
                <input type="hidden" name="howkey" value="CCS">
                <input type="hidden" name="payment_method" value="By Credit Card - via Shepheard-Walwyn Secure Server">
                <input type="hidden" name="subject" value="URGENT : Order to Shepheard Walwyn">
                <input type="hidden" name="email" value="<%=stremail%>">
                <input type="hidden" name="clients_ref" value="ZE/MNBS68">
                <input type="hidden" name="OrderDate" value="<%=date()%>">
                <input type="hidden" name="chName" value="<%=strfName%>&nbsp;<%=strlName%>">
                <input type="hidden" name="chTitle" value="">
                <input type="hidden" name="chAddress" value="<%=strAdd%>">
                <input type="hidden" name="chCountry" value="<%=strCountry%>">
                <input type="hidden" name="chPostCode" value="<%=strZip%>">
                <input type="hidden" name="ShipToPlace" value="<%=strSCountry%>">
                <input type="hidden" name="ShipPostcode" value="<%=strSZip%>">
                <input type="hidden" name="delivery" value="<%=strSAdd%>">
                <input type="hidden" name="telephone" value="<%=strPhone%>">
                <input type="hidden" name="cmd" value="_xclick">
                <input type="hidden" name="item_name" value="Shopping from Fatherson.com">
                <input type="hidden" name="business" value="lance@fatherson.com">
                <input type="hidden" name="order_number" value="<%=ono%>">
                <input type="hidden" name="amount" value="<%=session("GR")%>">
                <input type="hidden" name="payment_amount" value="<%=session("GR")%>">
                <br>
                <input name="submit"  type="submit" class="formbutton"  value="Continue"  >
              </form>
            </div></td>
        </tr>
      </table></td>
    <td width="96" valign="top"> 
      <div align="right"> 
        <!--#include file="include/rightbar.asp" -->
      </div></td>
    <td width="1" valign="top" background="images/rightimages.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4" valign="top" > 
      <!--#include file=include/footer.asp-->
    </td>
  </tr>
</table>
