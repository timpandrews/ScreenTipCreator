<%@ Language=VBScript %>
<%
Option Explicit
'Response.Expires = 0 
'Response.Buffer = True    '** Performance Enhancement ??? **

'On Error Resume Next
'**** Variable Sections *****
Dim x, y, action, Height, Width
dim objDC, strDC, strSQL, objRS_bld, strRS, objRS_db, objRS_db2
dim objFSO, objFile, objFileItem, objFolder, objFolderContents
Dim strImagePath, strImagePhysicalPath, image
dim row, rowd, PageID, DivID, aspDivTag, DivStyle
  
action = Request.Form("hdnAction")
PageID = Request.Form("hdnPageID")
DivID = Request.Form("hdnDivID")
aspDivTag = Request.Form("hdnDivTag")
DivStyle = Request.Form("DivStyle")

response.write("<div style=""left: 10px; top: 512px; width: 300px; position: absolute;"">")
'response.write("Action:" & action & "<br>")
'response.write("PageID:" & PageID & "<br>")
'response.write("DivID:" & DivID & "<br>")
'response.write("aspDivTag:" & aspDivTag & "<br>")
'response.write("DivStyle:" & DivStyle & "<br>")

'****** Creating Connection Object, Opening Connection, FSO Object etc. *****
Set objDC = Server.CreateObject("ADODB.Connection")
strDC = "Provider=SQLNCLI11;Data Source=I3TIMANDREWS\SQLEXPRESS;Persist Security Info=True;User ID=sa;Initial Catalog=ScreenTipCreator;Password=Adam321"
objDC.Open strDC 

set objRS_bld = server.CreateObject("ADODB.Recordset")	
objRS_bld.CursorLocation = 3
objRS_bld.CursorType = 3

set objRS_db = server.CreateObject("ADODB.Recordset")	
objRS_db.CursorLocation = 3
objRS_db.CursorType = 3

set objRS_db2 = server.CreateObject("ADODB.Recordset")	
objRS_db2.CursorLocation = 3
objRS_db2.CursorType = 3

'***** Create File Object with content of image director
strImagePath = Request.ServerVariables("PATH_INFO")
strImagePath = replace(strImagePath, "ScreenTipCreator.asp", "screenshots/~DoNotMove.png")
strImagePhysicalPath = Server.MapPath(strImagePath)

'response.Write("strImagePath:" & strImagePath & "<br>")
'response.Write("strImagePhysicalPath:" & strImagePhysicalPath & "<br>")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strImagePhysicalPath)
Set objFolder = objFile.ParentFolder
Set objFolderContents = objFolder.Files



'**** Database Actions *****
Select Case Action
    case "SaveOne"
        strRS = "UPDATE tblAttributes SET " &_
                    "Pos1_X = '" & Request.Form("hdnLastX") & "', " &_
                    "Pos1_Y = '" & Request.Form("hdnLastY")& "' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS)   
    case "SaveTwo"
        strRS = "UPDATE tblAttributes SET " &_
                    "Pos2_X = '" & Request.Form("hdnLastX") & "', " &_
                    "Pos2_Y = '" & Request.Form("hdnLastY")& "' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "Clear"
        strRS = "UPDATE tblAttributes SET " &_
                    "Pos1_X = '" & Request.Form("hdnLastX") & "', " &_
                    "Pos1_Y = '" & Request.Form("hdnLastY")& "', " &_
                    "Pos2_X = '" & Request.Form("hdnLastX") & "', " &_
                    "Pos2_Y = '" & Request.Form("hdnLastY")& "' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS)   
    case "UpdateData"
        strRS = "UPDATE tblAttributes SET " &_
                    "Pos1_X = '" & Request.Form("Pos1_X") & "', " &_
                    "Pos1_Y = '" & Request.Form("Pos1_Y")& "', " &_
                    "Pos2_X = '" & Request.Form("Pos2_X") & "', " &_
                    "Pos2_Y = '" & Request.Form("Pos2_Y")& "' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "SaveDivAttributes"
        strRS = "UPDATE tblAttributes SET " &_
                    "tabindex = '" & trim(Request.Form("tabindex")) & "', " &_
                    "objname = '" & trim(Request.Form("objname")) & "', " &_
                    "toolposition = '" & Request.Form("toolposition") & "', " &_
                    "toolwidth = '" & trim(Request.Form("toolwidth")) & "', " &_
                    "customtop = '" & Request.Form("customtop") & "', " &_
                    "customleft = '" & trim(Request.Form("customleft")) & "', " &_
                    "title = '" & trim(Request.Form("title")) & "' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "ClearDivAttributes"
        strRS = "UPDATE tblAttributes SET " &_
                    "tabindex = '', " &_
                    "objname = '', " &_
                    "toolposition = '', " &_
                    "toolwidth = '', " &_
                    "customtop = '', " &_
                    "customleft = '', " &_
                    "title = '' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS)      
    case "SavePageAttributes"
        strRS = "UPDATE tblAttributes SET " &_
                    "pagename = '" & trim(Request.Form("pagename")) & "', " &_
                    "pagetag = '" & trim(Request.Form("pagetag")) & "', " &_
                    "image = '" & trim(Request.Form("image")) & "' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "ClearPageAttributes"
        strRS = "UPDATE tblAttributes SET " &_
                    "pagename = '', " &_
                    "pagetag = '', " &_
                    "image = '' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "AddNewPage"
        strRS = "INSERT tblPage(PageName, pagetag, ImageName) VALUES (" &_
            "'" & Request.Form("pagename") & "', " &_
            "'" & Request.Form("pagetag") & "', " &_
            "'" & Request.Form("image") & "') " 
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "UpdatePage"
        strRS = "UPDATE tblPage SET " &_
                    "pagename = '" & trim(Request.Form("dbpagename")) & "', " &_
                    "pagetag = '" & trim(Request.Form("dbpagetag")) & "', " &_
                    "imagename = '" & trim(Request.Form("dbimagename")) & "' " &_
                "WHERE (PageID = '" & trim(PageID) & "')"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "PushPageDB2Build"
        strRS = "UPDATE tblAttributes SET " &_
                    "pagename = '" & trim(Request.Form("dbpagename")) & "', " &_
                    "pagetag = '" & trim(Request.Form("dbpagetag")) & "', " &_
                    "image = '" & trim(Request.Form("dbimagename")) & "' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "AddNewDiv"
        strRS = "INSERT tblDiv(PageID, DivTag, tabindex, DivName, toolposition, toolwidth, customleft, customtop, title, " &_
                                "CurX, CurY, Pos1_X, Pos1_Y, Pos2_X, Pos2_Y, Style) " &_
                "VALUES (" &_
                    "'" & trim(PageID) & "', " &_
                    "'" & aspDivTag & "', " &_
                    "'" & Request.Form("tabindex") & "', " &_
                    "'" & Request.Form("objname") & "', " &_
                    "'" & Request.Form("toolposition") & "', " &_
                    "'" & Request.Form("toolwidth") & "', " &_
                    "'" & Request.Form("customleft") & "', " &_
                    "'" & Request.Form("customtop") & "', " &_
                    "'" & Request.Form("title") & "', " &_
                    "'" & Request.Form("CoordX") & "', " &_
                    "'" & Request.Form("CoordY") & "', " &_
                    "'" & Request.Form("Pos1_X") & "', " &_
                    "'" & Request.Form("Pos1_Y") & "', " &_
                    "'" & Request.Form("Pos2_X") & "', " &_
                    "'" & Request.Form("Pos2_Y") & "', " &_
                    "'" & Request.Form("DivStyle") & "') " 
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "DeletePage"
        strRS = "UPDATE tblPage SET " &_
                    "Active = '0' " &_
                "WHERE PageID = '" & trim(PageID) & "'"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "DeleteDiv"
        strRS = "UPDATE tblDiv SET " &_
                    "Active = '0' " &_
                "WHERE DivID = '" & trim(DivID) & "'"
        'response.write("strRS:" & strRS & "<br>")
	    set objRS_bld = objDC.Execute(strRS) 
    case "PushDivDB2Build"
        strRS = "SELECT divID, PageID, DivTag, tabindex, DivName, toolposition, toolwidth, customleft, customtop, title, " &_
                        "CurX, CurY, Pos1_X, Pos1_Y, Pos2_X, Pos2_Y " &_
                "FROM tblDiv WHERE (divID = '" & trim(divID) & "') "
        objRS_db.Open(strRS), objDC

        strRS = "UPDATE tblAttributes SET " &_
                    "tabindex = '" & trim(objRS_db.Fields("tabindex")) & "', " &_
                    "objname = '" & trim(objRS_db.Fields("DivName")) & "', " &_
                    "toolposition = '" & trim(objRS_db.Fields("toolposition")) & "', " &_
                    "toolwidth = '" & trim(objRS_db.Fields("toolwidth")) & "', " &_
                    "customtop = '" & trim(objRS_db.Fields("customtop")) & "', " &_
                    "customleft = '" & trim(objRS_db.Fields("customleft")) & "', " &_
                    "title = '" & trim(objRS_db.Fields("title")) & "', " &_
                    "CurX = '" & trim(objRS_db.Fields("CurX")) & "', " &_
                    "CurY = '" & trim(objRS_db.Fields("CurY")) & "', " &_
                    "Pos1_X = '" & trim(objRS_db.Fields("Pos1_X")) & "', " &_
                    "Pos1_Y = '" & trim(objRS_db.Fields("Pos1_Y")) & "', " &_
                    "Pos2_X = '" & trim(objRS_db.Fields("Pos2_X")) & "', " &_
                    "Pos2_Y = '" & trim(objRS_db.Fields("Pos2_Y")) & "' " &_
                "WHERE (id = 1)"
        'response.write("strRS:" & strRS & "<br>")
        objRS_db.Close
	    set objRS_bld = objDC.Execute(strRS) 
    case else
End Select

'**** Recordsets & File System Objects *****
'**** Open Build Record *****
strSQL = "SELECT Pos1_X, Pos1_Y, Pos2_X, Pos2_Y, tabindex, objname, toolposition, " &_ 
                "toolwidth, title, pagename, pagetag, image, customleft, customtop, CurX, CurY " &_
         "FROM tblAttributes"
objRS_bld.Open(strSQL), objDC

'**** Open Database Recordset *****
strSQL = "SELECT pageid, pagename, pagetag, ImageName " &_
         "FROM tblPage " &_
         "WHERE Active = '1' AND tmpView = 'Y' "
objRS_db.Open(strSQL), objDC

'*** Assign ImageName 
if action = "PushPageDB2Build" then
    image = trim(Request.Form("dbimagename"))
else
    image = trim(objRS_bld.Fields("image"))
end if
'response.Write("image:" & image & "<br>")

if len(image) > 0 then
    strImagePath = replace(strImagePath, "~DoNotMove.png", image)
end if
'response.Write("image:" & image & "<br>")

response.write("</div>")
%>

<html>
<head>

<title>Screen Tip Creator</title>

<!--<link href='https://fonts.googleapis.com/css?family=Source+Sans+Pro:300,400,600,700,900' rel='stylesheet' type='text/css'>-->
<link rel="stylesheet" href="css/font-awesome.min.css">

<link rel="stylesheet" href="css/reset_STC.css">
<link rel="stylesheet" href="css/template_STC.css">
<link rel="stylesheet" href="css/tooltip_STC.css">


<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/jgestures.min.js"></script>
<script type="text/javascript">

    var myX, myY, xyOn, myMouseX, myMouseY, numClicks
    xyOn = true;

    function getXYPosition(e) {
        document.getElementById("hdnLastX").value = document.getElementById("CoordX").value
        document.getElementById("hdnLastY").value = document.getElementById("CoordY").value
        myMouseX = (e || event).clientX;
        myMouseY = (e || event).clientY;
        if (document.documentElement.scrollTop > 0) {
            myMouseY = myMouseY + document.documentElement.scrollTop;
        }
        if (xyOn) {
            if (myMouseX < 325) {
                document.getElementById("CoordX").value = myMouseX
                document.getElementById("CoordY").value = myMouseY
            }
        }
    }

    function toggleXY() {
        xyOn = !xyOn;
        document.getElementById('xyLink').blur();
        return false;
    }

    function cmdSaveOne_onclick(PageID) {
        document.getElementById("hdnAction").value = "SaveOne";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function cmdSaveTwo_onclick(PageID) {
        document.getElementById("hdnAction").value = "SaveTwo";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function cmdClearData_onclick(PageID) {
        document.getElementById("hdnAction").value = "Clear";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function cmdUpdateData_onclick(PageID) {
        document.getElementById("hdnAction").value = "UpdateData";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function cmdClearDiv_onclick(PageID) {
        document.getElementById("hdnAction").value = "";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("DivStyle").value = "";
        document.getElementById("frmMain").submit();
    }

    function cmdSaveDivAttributes_onclick(PageID) {
        document.getElementById("hdnAction").value = "SaveDivAttributes";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    } 

    function cmdClearDivAttributes_onclick(PageID) {
        document.getElementById("hdnAction").value = "ClearDivAttributes";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("DivStyle").value = "";
        document.getElementById("frmMain").submit();
    }

    function cmdSavePageAttributes_onclick(PageID) {
        document.getElementById("hdnAction").value = "SavePageAttributes";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function cmdClearPageAttributes_onclick(PageID) {
        document.getElementById("hdnAction").value = "ClearPageAttributes";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function cmdAddNewPage_onclick() {
        document.getElementById("hdnAction").value = "AddNewPage";
        document.getElementById("frmMain").submit();
    }

    function cmdAddNewDiv_onclick(PageID) {
        var DivTag = document.getElementById("visDivTag").value
        document.getElementById("hdnAction").value = "AddNewDiv";
        document.getElementById("hdnDivTag").value = DivTag;
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function OpenPage(PageID) {
        document.getElementById("hdnAction").value = "OpenPage";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function ClosePage() {
        document.getElementById("hdnPageID").value = "";
        document.getElementById("frmMain").submit();
    }

    function DeletePage(PageID) {
        if (confirm("Are You Sure You Want to Delete ID #" + PageID)) {
            document.getElementById("hdnAction").value = "DeletePage";
            document.getElementById("hdnPageID").value = PageID;
            document.getElementById("frmMain").submit();
        }
    }

    function DeleteDiv(DivID, PageID) {
        if (confirm("Are You Sure You Want to Delete ID #" + DivID)) {
            document.getElementById("hdnAction").value = "DeleteDiv";
            document.getElementById("hdnPageID").value = PageID;
            document.getElementById("hdnDivID").value = DivID;
            document.getElementById("frmMain").submit();
        }
    }

    function UpdatePage(PageID) {
        document.getElementById("hdnAction").value = "UpdatePage";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();

    }

    function PushPageDB2Build(PageID) {
        document.getElementById("hdnAction").value = "PushPageDB2Build";
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }

    function PushDivDB2Build(DivID, PageID) {
        document.getElementById("hdnAction").value = "PushDivDB2Build";
        document.getElementById("hdnDivID").value = DivID;
        document.getElementById("hdnPageID").value = PageID;
        document.getElementById("frmMain").submit();
    }
    
    function cmdBuildBox_onclick() {
        var top = document.getElementById("Pos1_Y").value;
        var left = document.getElementById("Pos1_X").value;
        var width = document.getElementById("Pos2_X").value - document.getElementById("Pos1_X").value;
        var height = document.getElementById("Pos2_Y").value - document.getElementById("Pos1_Y").value;
        var position = document.getElementById("toolposition").value;
        var ttwidth = document.getElementById("toolwidth").value;
       // var ttheight = document.getElementById("toolheight").value;

        var StyleValue = 'position:absolute; ' +
                    'left:' + left + 'px; ' +
                    'top:' + top + 'px; ' +
                    'width:' + width + 'px; ' +
                    'height:' + height + 'px; " ' 

        var DivTag = '{div ' +
                'tabindex="' + document.getElementById("tabindex").value + '" ' +
                'Style="' + StyleValue +
                'obj-name="' + document.getElementById("objname").value + '" ' +
                'tool-position="' + document.getElementById("toolposition").value + '" ' +
                'tool-width="' + document.getElementById("toolwidth").value + '" ' +
                'custom-top="' + document.getElementById("customtop").value + '" ' +
                'custom-left="' + document.getElementById("customleft").value + '" ' +
                'title="' + document.getElementById("title").value + '"> ' +
            '{/div>'
        
        document.getElementById("BoxTop").value = top;
        document.getElementById("BoxLeft").value = left;
        document.getElementById("BoxWidth").value = width;
        document.getElementById("BoxHeight").value = height;

        document.getElementById("visDivTag").value = DivTag;
        document.getElementById("DivStyle").value = StyleValue;

        var newdiv = document.createElement('div');
        newdiv.setAttribute('id', 'Box');
        newdiv.setAttribute('tabindex', '1');
        newdiv.style.top = top;
        newdiv.style.left = left; 
        newdiv.style.width = width;
        newdiv.style.height = height;
        newdiv.style.position = "absolute"; 
        newdiv.style.background = "hsla(200,100%,90%,0.22)";
        newdiv.style.border = "2px solid #FF9900";
        document.body.appendChild(newdiv);

        var tt = document.createElement('div');
        tt.setAttribute('id', 'tt');
        tt.setAttribute('tabindex', '2');
        tt.innerHTML = document.getElementById("title").value;
        tt.style.top = top;
        tt.style.left = left;
        tt.style.width = ttwidth;
        tt.style.height = "";
        tt.style.position = "absolute";
        tt.style.background = "hsla(200,100%,90%,0.22)";
        tt.style.backgroundImage = "linear-gradient(to bottom left, #50EBCE 0%, #5355C9 100%)";
        tt.style.border = "2px solid White";
        tt.style.boxShadow = "10px 10px 20px #888888"; 
        tt.style.borderRadius = "10px";
        tt.style.padding = "8px 8px";
        tt.style.color = "#ffffff";
        tt.style.textAlign = "left";
        tt.style.fontSize = "16px";
        tt.style.margin = "0";
        document.body.appendChild(tt);

        if (position == "top") {
            //alert("top");
            tt.style.left = newdiv.offsetLeft + ( $(newdiv).outerWidth() / 2 ) - ($(tt).outerWidth() / 2 );
            tt.style.top = newdiv.offsetTop - $(tt).outerHeight() - 12
        }
        else if (position == "bottom") {
            //alert("bottom");
            tt.style.left = newdiv.offsetLeft + ($(newdiv).outerWidth() / 2) - ($(tt).outerWidth() / 2);
            tt.style.top = newdiv.offsetTop + $(newdiv).outerHeight() + 12
        }
        else if (position == "right") {
            //alert("right");
            tt.style.left = newdiv.offsetLeft + $(newdiv).outerWidth() + 12
            tt.style.top = newdiv.offsetTop + ( $(newdiv).outerHeight() / 2 ) - ( $(tt).outerHeight() / 2 )
               
        }
        else if (position == "left") {
            //alert("left");
            tt.style.left = newdiv.offsetLeft - $(tt).outerWidth() - 12;
            tt.style.top = newdiv.offsetTop + ( $(newdiv).outerHeight() / 2 ) - ( $(tt).outerHeight() / 2 )
        }
        else if (position == "custom-1") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-2") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-3") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-4") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-5") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-6") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-7") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-8") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-9") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-10") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-11") {
            //alert("custom-11");
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-12") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        
        var arrow = document.createElement('div');
        arrow.setAttribute('id', 'arrow');
        arrow.setAttribute('tabindex', '3');
        arrow.style.position = "absolute";
        arrow.style.borderWidth = "10px";
        arrow.style.borderStyle = "solid";
        arrow.style.top = "0px";
        arrow.style.left = "0px";
        document.body.appendChild(arrow);

        x = $(tt).position();

        if (position == "bottom") {
            arrow.style.left = ($(tt).outerWidth() / 2 + x.left);
            arrow.style.top = x.top - 20;
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "transparent transparent #ffffff transparent";
        }
        else if (position == "top") {
            arrow.style.left = ($(tt).outerWidth() / 2 + x.left);
            arrow.style.top = x.top + $(tt).outerHeight();
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "#ffffff transparent transparent transparent";
        }
        else if (position == "right") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = ($(tt).outerHeight() / 2 + x.top);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "left") {
            arrow.style.left = (x.left + $(tt).outerWidth());
            arrow.style.top = ($(tt).outerHeight() / 2 + x.top);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent transparent transparent #ffffff";
        }
        else if (position == "custom-1") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.85));
            arrow.style.top = x.top - 20;
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "transparent transparent #ffffff transparent";
        }
        else if (position == "custom-2") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + ($(tt).outerHeight() * 0.15));
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-3") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + ($(tt).outerHeight() * 0.5));
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-4") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + ($(tt).outerHeight() * 0.85));
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-5") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.85));
            arrow.style.top = x.top + $(tt).outerHeight();
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "#ffffff transparent transparent transparent";
        }
        else if (position == "custom-6") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.5));
            arrow.style.top = x.top + $(tt).outerHeight();
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "#ffffff transparent transparent transparent";
        }
        else if (position == "custom-7") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.15 ));
            arrow.style.top = x.top + $(tt).outerHeight();
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "#ffffff transparent transparent transparent";
        }
        else if (position == "custom-8") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + $(tt).outerHeight() * 0.85);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-9") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + $(tt).outerHeight() * 0.5);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-10") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + $(tt).outerHeight() * 0.15 );
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-11") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.15 ));
            arrow.style.top = x.top - 20;
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "transparent transparent #ffffff transparent";
        }
        else if (position == "custom-12") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.5));
            arrow.style.top = x.top - 20;
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "transparent transparent #ffffff transparent";
        }
    }

    function cmdBuildCircle_onclick() {
        var top = document.getElementById("Pos1_Y").value;
        var left = document.getElementById("Pos2_X").value;
        var radius = document.getElementById("Pos1_X").value - document.getElementById("Pos2_X").value;
        var diameter = radius * 2;
        var position = document.getElementById("toolposition").value;
        var ttwidth = document.getElementById("toolwidth").value;

        var StyleValue = 'position:absolute; ' +
                   'left:' + left + 'px; ' +
                    'top:' + top + 'px; ' +
                    'width:' + diameter + 'px; ' +
                    'height:' + diameter + 'px; ' +
                    '-webkit-border-radius:' + diameter + 'px; ' +
                    '-moz-border-radius:' + diameter + 'px; ' +
                    'border-radius:' + diameter + 'px; "' 
        
        var DivTag = '{div ' +
               'tabindex="' + document.getElementById("tabindex").value + '" ' +
               'Style="' + StyleValue +
               'obj-name="' + document.getElementById("objname").value + '" ' +
               'tool-position="' + document.getElementById("toolposition").value + '" ' +
               'tool-width="' + document.getElementById("toolwidth").value + '" ' +
               'custom-top="' + document.getElementById("customtop").value + '" ' +
               'custom-left="' + document.getElementById("customleft").value + '" ' +
               'title="' + document.getElementById("title").value + '"> ' +
           '{/div>'

        document.getElementById("CircleTop").value = top
        document.getElementById("CircleLeft").value = left
        document.getElementById("CircleRadius").value = radius
        document.getElementById("CircleDiameter").value = diameter

        document.getElementById("visDivTag").value = DivTag;
        document.getElementById("DivStyle").value = StyleValue;

        var newdiv = document.createElement('div');
        newdiv.setAttribute('id', 'Circle');
        newdiv.style.top = top;
        newdiv.style.left = left;
        newdiv.style.width = diameter;
        newdiv.style.height = diameter;
        newdiv.style.WebkitBorderRadius = radius;
        newdiv.style.MozBorderRadius = radius;
        newdiv.style.borderRadius = radius;
        newdiv.style.position = "absolute";
        newdiv.style.background = "hsla(200,100%,90%,0.22)";
        newdiv.style.border = "1px solid #000";
        document.body.appendChild(newdiv);

        var tt = document.createElement('div');
        tt.setAttribute('id', 'tt');
        tt.setAttribute('tabindex', '2');
        tt.innerHTML = document.getElementById("title").value;
        tt.style.top = top;
        tt.style.left = left;
        tt.style.width = ttwidth;
        tt.style.height = "";
        tt.style.position = "absolute";
        tt.style.background = "hsla(200,100%,90%,0.22)";
        tt.style.backgroundImage = "linear-gradient(to bottom left, #50EBCE 0%, #5355C9 100%)";
        tt.style.border = "2px solid White";
        tt.style.boxShadow = "10px 10px 20px #888888";
        tt.style.borderRadius = "10px";
        tt.style.padding = "8px 8px";
        tt.style.color = "#ffffff";
        tt.style.textAlign = "left";
        tt.style.fontSize = "16px";
        tt.style.margin = "0";
        document.body.appendChild(tt);
        
        if (position == "top") {
            //alert("top");
            tt.style.left = newdiv.offsetLeft + ($(newdiv).outerWidth() / 2) - ($(tt).outerWidth() / 2);
            tt.style.top = newdiv.offsetTop - $(tt).outerHeight() - 12
        }
        else if (position == "bottom") {
            //alert("bottom");
            tt.style.left = newdiv.offsetLeft + ($(newdiv).outerWidth() / 2) - ($(tt).outerWidth() / 2);
            tt.style.top = newdiv.offsetTop + $(newdiv).outerHeight() + 12
        }
        else if (position == "right") {
            //alert("right");
            tt.style.left = newdiv.offsetLeft + $(newdiv).outerWidth() + 12
            tt.style.top = newdiv.offsetTop + ($(newdiv).outerHeight() / 2) - ($(tt).outerHeight() / 2)

        }
        else if (position == "left") {
            //alert("left");
            tt.style.left = newdiv.offsetLeft - $(tt).outerWidth() - 12;
            tt.style.top = newdiv.offsetTop + ($(newdiv).outerHeight() / 2) - ($(tt).outerHeight() / 2)
        }
        else if (position == "custom-1") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-2") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-3") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-4") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-5") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-6") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-7") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-8") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-9") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-10") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-11") {
            //alert("custom-11");
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }
        else if (position == "custom-12") {
            tt.style.left = document.getElementById("customleft").value;
            tt.style.top = document.getElementById("customtop").value;
        }

        var arrow = document.createElement('div');
        arrow.setAttribute('id', 'arrow');
        arrow.setAttribute('tabindex', '3');
        arrow.style.position = "absolute";
        arrow.style.borderWidth = "10px";
        arrow.style.borderStyle = "solid";
        arrow.style.top = "0px";
        arrow.style.left = "0px";
        document.body.appendChild(arrow);

        x = $(tt).position();

        if (position == "bottom") {
            arrow.style.left = ($(tt).outerWidth() / 2 + x.left);
            arrow.style.top = x.top - 20;
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "transparent transparent #ffffff transparent";
        }
        else if (position == "top") {
            arrow.style.left = ($(tt).outerWidth() / 2 + x.left);
            arrow.style.top = x.top + $(tt).outerHeight();
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "#ffffff transparent transparent transparent";
        }
        else if (position == "right") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = ($(tt).outerHeight() / 2 + x.top);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "left") {
            arrow.style.left = (x.left + $(tt).outerWidth());
            arrow.style.top = ($(tt).outerHeight() / 2 + x.top);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent transparent transparent #ffffff";
        }
        else if (position == "custom-1") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.85));
            arrow.style.top = x.top - 20;
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "transparent transparent #ffffff transparent";
        }
        else if (position == "custom-2") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + ($(tt).outerHeight() * 0.15));
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-3") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + ($(tt).outerHeight() * 0.5));
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-4") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + ($(tt).outerHeight() * 0.85));
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-5") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.85));
            arrow.style.top = x.top + $(tt).outerHeight();
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "#ffffff transparent transparent transparent";
        }
        else if (position == "custom-6") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.5));
            arrow.style.top = x.top + $(tt).outerHeight();
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "#ffffff transparent transparent transparent";
        }
        else if (position == "custom-7") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.15));
            arrow.style.top = x.top + $(tt).outerHeight();
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "#ffffff transparent transparent transparent";
        }
        else if (position == "custom-8") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + $(tt).outerHeight() * 0.85);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-9") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + $(tt).outerHeight() * 0.5);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-10") {
            arrow.style.left = (x.left - 20);
            arrow.style.top = (x.top + $(tt).outerHeight() * 0.15);
            arrow.style.marginTop = "-10px";
            arrow.style.borderColor = "transparent #ffffff transparent transparent";
        }
        else if (position == "custom-11") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.15));
            arrow.style.top = x.top - 20;
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "transparent transparent #ffffff transparent";
        }
        else if (position == "custom-12") {
            arrow.style.left = (x.left + ($(tt).outerWidth() * 0.5));
            arrow.style.top = x.top - 20;
            arrow.style.marginLeft = "-10px";
            arrow.style.borderColor = "transparent transparent #ffffff transparent";
        }
    }

    function CreateJSON() {
        
        var NumDiv = document.getElementById("selNumDiv").value;
        //alert(document.getElementById("selObjName" + NumDiv).value)

        jsonWindow=window.open('','','width=1200, height=300, left=50, top=200, scrollbars=1')
        //jsonWindow.document.write("{'_id': ObjectID('*****"  + document.getElementById("hdnPageID").value + "*****'), ");
        //jsonWindow.document.write("<br>&nbsp;");
        jsonWindow.document.write("{'image': '" + document.getElementById("selImageName").value + "', ")
        //jsonWindow.document.write("<br>&nbsp;");
        jsonWindow.document.write("'items': [");

        for (var i=1; i<=NumDiv; i++) {
            //jsonWindow.document.write("<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"); 
            jsonWindow.document.write("{'obj-name': '" + document.getElementById("selObjName" + i).value + "', ");
            jsonWindow.document.write("'style': '" + document.getElementById("selStyle" + i).value + "', ");
            jsonWindow.document.write("'tabindex': '" + document.getElementById("selTabindex" + i).value + "', ");
            jsonWindow.document.write("'title': '" + document.getElementById("selTitle" + i).value + "', ");
            jsonWindow.document.write("'tool-position': '" + document.getElementById("selToolPosition" + i).value + "', ");
            jsonWindow.document.write("'tool-width': '" + document.getElementById("selToolWidth" + i).value + "', ");
            jsonWindow.document.write("'custom-left': '" + document.getElementById("selCustomLeft" + i).value + "', ");
            if (i==NumDiv) {
                jsonWindow.document.write("'custom-top': '" + document.getElementById("selCustomTop" + i).value + "'}], ");  }
            else {
                jsonWindow.document.write("'custom-top': '" + document.getElementById("selCustomTop" + i).value + "'}, ");  }
        }

        //jsonWindow.document.write("<br>&nbsp;");
        //alert(document.getElementById("selPageTag").value)
        if (document.getElementById("selPageTag").value != "") {
            //alert("if")
            jsonWindow.document.write("'view': '" + document.getElementById("selPageName").value + "', ")   
            jsonWindow.document.write("'tag': '" + document.getElementById("selPageTag").value + "'} ") }
        else {
            //alert("else")
            jsonWindow.document.write("'view': '" + document.getElementById("selPageName").value + "'} ") }

        jsonWindow.focus();

    }

    function CreateHTML() {
        
        HTMLWindow=window.open('','','width=600, height=600, left=50, top=200, location=no, menubar=no, resizable=yes, scrollbars=yes, status=no, titlebar=no, toolbar=no')
        HTMLWindow.document.write("<p>This is 'myWindow'</p>")
        HTMLWindow.focus()

    }
</script>

<script type="text/javascript">
	document.onmouseup = getXYPosition;
</script>

</head>
<body style="background-image:url('<%=strImagePath%>'); background-repeat:no-repeat; background-size:320px">

    <form id="frmMain" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" <p style="font-size:12px; font-family:'Arial Unicode MS'">
        <div style="left: 330px; top: 0px; width:auto; height:auto; position:absolute; border:1px solid" >
        <div style="background:#9494B8; border: 1px solid">
             <table width="450" border="1">
                <tr>
                    <td colspan="4" align="center" ><p style="font-size:14px; font-weight:bold; ">Page Attributes</p></td>
                </tr>
                <tr>
                    <td>Page Name:</td>
                    <td><input type="text" style="width:200" id="pagename" name="pagename" value="<%=trim(objRS_bld.Fields("pagename")) %>"/></td>
                    <td>Page Tag:</td>
                    <td><input type="text" style="width:100" id="pagetag" name="pagetag" value="<%=trim(objRS_bld.Fields("pagetag")) %>"/></td>
                </tr>
                <tr>
                    <td>Image:</td>
                    <td colspan="3">
                        <select id="image" name="image">
                            <option></option> <%
                            For Each objFileItem In objFolderContents %>
                                <option <%if trim(objRS_bld.Fields("image")) = trim(objFileItem.Name) then%>selected<%end if%>><%=objFileItem.Name%></option> <%
                            Next %>
                        </select>
                    </td>
                </tr>
                 <tr>
                    <td colspan="4" align="center">
                        <input type="button" value="Save/Update Page Attributes" onclick="cmdSavePageAttributes_onclick(<%=PageID%>)" />&nbsp;&nbsp;
                        <input type="button" value="Clear Page Attributes" onclick="cmdClearPageAttributes_onclick(<%=PageID%>)" />
                    </td>
                </tr>
            </table>
        </div>
        <div style="background:#FFE0A3; border: 1px solid">  
            <table width="450" border="0">
                <tr>
                    <td colspan="4" align="center" ><p style="font-size:14px; font-weight:bold; ">Div Attributes</p></td>
                </tr>
                <tr>
                    <td>tabindex:</td>
                    <td><input type="text" style="width:50" id="tabindex" name="tabindex" value="<%=trim(objRS_bld.Fields("tabindex")) %>"/></td>
                    <td>obj-name:</td>
                    <td><input type="text" style="width:150" id="objname" name="objname" value="<%=trim(objRS_bld.Fields("objname")) %>"/></td>
                </tr>
                <tr>
                    <td>tool-position:</td>
                    <td>
                        <select id="toolposition" name="toolposition">
                            <option value="" ></option>        
                            <option value="top" <%if trim(objRS_bld.Fields("toolposition")) = "top" then%>selected<%end if%>>top</option>
                            <option value="bottom" <%if trim(objRS_bld.Fields("toolposition")) = "bottom" then%>selected<%end if%>>bottom</option>
                            <option value="right" <%if trim(objRS_bld.Fields("toolposition")) = "right" then%>selected<%end if%>>right</option>
                            <option value="left" <%if trim(objRS_bld.Fields("toolposition")) = "left" then%>selected<%end if%>>left</option>
                            <option value="custom-1" <%if trim(objRS_bld.Fields("toolposition")) = "custom-1" then%>selected<%end if%>>custom-1</option>
                            <option value="custom-2" <%if trim(objRS_bld.Fields("toolposition")) = "custom-2" then%>selected<%end if%>>custom-2</option>
                            <option value="custom-3" <%if trim(objRS_bld.Fields("toolposition")) = "custom-3" then%>selected<%end if%>>custom-3</option>
                            <option value="custom-4" <%if trim(objRS_bld.Fields("toolposition")) = "custom-4" then%>selected<%end if%>>custom-4</option>
                            <option value="custom-5" <%if trim(objRS_bld.Fields("toolposition")) = "custom-5" then%>selected<%end if%>>custom-5</option>
                            <option value="custom-6" <%if trim(objRS_bld.Fields("toolposition")) = "custom-6" then%>selected<%end if%>>custom-6</option>
                            <option value="custom-7" <%if trim(objRS_bld.Fields("toolposition")) = "custom-7" then%>selected<%end if%>>custom-7</option>
                            <option value="custom-8" <%if trim(objRS_bld.Fields("toolposition")) = "custom-8" then%>selected<%end if%>>custom-8</option>
                            <option value="custom-9" <%if trim(objRS_bld.Fields("toolposition")) = "custom-9" then%>selected<%end if%>>custom-9</option>
                            <option value="custom-10" <%if trim(objRS_bld.Fields("toolposition")) = "custom-10" then%>selected<%end if%>>custom-10</option>
                            <option value="custom-11" <%if trim(objRS_bld.Fields("toolposition")) = "custom-11" then%>selected<%end if%>>custom-11</option>
                            <option value="custom-12" <%if trim(objRS_bld.Fields("toolposition")) = "custom-12" then%>selected<%end if%>>custom-12</option>
                        </select>
                    </td>
                    <td>tool-width:</td>
                    <td><input type="text" style="width:50" id="toolwidth" name="toolwidth" value="<%=trim(objRS_bld.Fields("toolwidth")) %>"/></td>
                </tr>
                <tr>
                    <td>custom-left*:</td>
                    <td><input type="text" style="width:50" id="customleft" name="customleft" value="<%=trim(objRS_bld.Fields("customleft")) %>"/></td>
                    <td>custom-top*:</td>
                    <td><input type="text" style="width:50" id="customtop" name="customtop" value="<%=trim(objRS_bld.Fields("customtop")) %>"/></td>
                </tr>
                <tr>
                    <td style="vertical-align:top;">title/content:</td>
                    <td colspan="3">
                        <textarea rows="4" cols="40" id="title" name="title"><% response.Write(objRS_bld.Fields("title")) %></textarea>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" align="center">
                        <input type="button" value="Save/Update Div Attributes" onclick="cmdSaveDivAttributes_onclick(<%=PageID%>)" />&nbsp;&nbsp;
                        <input type="button" value="Clear Div Attributes" onclick="cmdClearDivAttributes_onclick(<%=PageID%>)" />
                    </td>
                </tr>
            </table>
        </div>
        <div style="background:#99FFCC; border: 1px solid">  
            <table width="400">
                <tr>
                    <td colspan="4" align="center"><p style="font-size:14px; font-weight:bold; ">Click to Record XY Coordinates</p></td>
                </tr>
                <tr>
                    <td>X - Left</td>
                    <td><input type="text" value="<%=objRS_bld.Fields("CurX")%>" id="CoordX" name="CoordX" style="width:75px"/></td>
                    <td>Y - Top </td>
                    <td><input type="text" value="<%=objRS_bld.Fields("CurY")%>" id="CoordY" name="CoordY" style="width:75px"/></td>
                </tr>
                <tr>
                    <td colspan="4" align="center">
                        <input type="button" value="Save Position One" onclick="cmdSaveOne_onclick(<%=PageID%>)" />&nbsp;&nbsp;
                        <input type="button" value="Save Position Two" onclick="cmdSaveTwo_onclick(<%=PageID%>)" />
                    </td>
                </tr>
            </table>
        </div>
        <div style="background:#99E6FF; border: 1px solid">  
            <table width="400">
                <tr>
                    <td colspan="4" align="center"><p style="font-size:14px; font-weight:bold; ">Saved Coord #1 (Top-Left or Center-Top)</p></td>
                </tr>
                <tr>
                    <td>X - Left</td>
                    <td><input type="text" value="<%=objRS_bld.Fields("Pos1_X")%>" id="Pos1_X" name="Pos1_X" style="width:75px"/></td>
                    <td>Y - Top </td>
                    <td><input type="text" value="<%=objRS_bld.Fields("Pos1_Y")%>" id="Pos1_Y" name="Pos1_Y" style="width:75px"/></td>
                </tr>
                <tr>
                    <td colspan="4" align="center"><p style="font-size:14px; font-weight:bold; ">Saved Coord #2 (Bottom-Right or Left-Middle)</p></td>
                </tr>
                <tr>
                    <td>X - Left</td>
                    <td><input type="text" value="<%=objRS_bld.Fields("Pos2_X")%>" id="Pos2_X" name="Pos2_X" style="width:75px"/></td>
                    <td>Y - Top </td>
                    <td><input type="text" value="<%=objRS_bld.Fields("Pos2_Y")%>" id="Pos2_Y" name="Pos2_Y" style="width:75px"/></td> 
                </tr>
                <tr>
                    <td colspan="4" align="center">
                        <input type="button" value="Update Data" onclick="cmdUpdateData_onclick(<%=PageID%>)" />&nbsp;&nbsp;
                        <input type="button" value="Clear Data" onclick="cmdClearData_onclick(<%=PageID%>)" />&nbsp;&nbsp;
                        <input type="button" value="Clear Div" onclick="cmdClearDiv_onclick(<%=PageID%>)" />
                    </td>
                </tr>
            </table>
        </div>
        </div>
        <div style="left: 788px; top: 0px; width:auto; height:auto; position:absolute; border:1px solid" >
        <div style="background:#FFFFCC; border: 1px solid; height:90px">  
            <table width="400">
                <tr>
                    <td colspan="4" align="center"><input type="button" value="Build Box" onclick="cmdBuildBox_onclick()" /></td>
                </tr>
                <tr>
                    <td>Box Left</td>
                    <td><input type="text" value="" id="BoxLeft" name="BoxLeft" style="width:75px"/></td>
                    <td>Box Top</td>
                    <td><input type="text" value="" id="BoxTop" name="BoxTop" style="width:75px"/></td>
                </tr>
                <tr>
                    <td>Box Width</td>
                    <td><input type="text" value="" id="BoxWidth" name="BoxTop" style="width:75px"/></td>
                    <td>Box Height</td>
                    <td><input type="text" value="" id="BoxHeight" name="BoxTop" style="width:75px"/></td>
                </tr>
            </table>
        </div>
        <div style="background:#FFCCCC; border: 1px solid; height:90px;">
            <table width="400">
                <tr>
                    <td colspan="4" align="center"><input type="button" value="Build Circle" onclick="cmdBuildCircle_onclick()" /></td>
                </tr>
                <tr>
                    <td>Circle Left</td>
                    <td><input type="text" value="" id="CircleLeft" name="CircleLeft" style="width:75px"/></td>
                    <td>Circle Top</td>
                    <td><input type="text" value="" id="CircleTop" name="CircleTop" style="width:75px"/></td>
                </tr>
                <tr>
                    <td>Circle Diameter</td>
                    <td><input type="text" value="" id="CircleDiameter" name="CircleDiameter" style="width:75px"/></td>
                    <td>Circle Radius</td>
                    <td><input type="text" value="" id="CircleRadius" name="CircleRadius" style="width:75px"/></td>
                </tr>
            </table>
        </div>
        <div style="background:#F0B2FF; border: 1px solid; height:274px;">  
            <table width="400">
                <tr>
                    <td colspan="4" align="center"><p style="font-size:14px; font-weight:bold; ">Div Tag</p></td>
                </tr>
                <tr>
                    <td colspan="4" align="center">
                        <textarea rows="4" cols="55" id="visDivTag" name="visDivTag"><%=aspDivTag%></textarea>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" align="center"><p style="font-size:14px; font-weight:bold; ">Div Style</p></td>
                </tr>
                <tr>
                    <td colspan="4" align="center">
                        <textarea rows="4" cols="55" id="DivStyle" name="DivStyle"><%=DivStyle%></textarea>
                    </td>
                </tr>
            </table>
        </div>
        </div>
        <div style="left: 330px; top: 465px; width:auto; height:auto; position:absolute" >
        <div style="background:#FF9999; border: 2px solid">  
            <table width="925" >
                <tr>
                    <td class="title" colspan="3">Database</td>
                </tr>
                <tr>
                    <td class="buttonbar">Select Page ID : <%response.Write(PageID)%></td>
                    <td><input type="button" value="Add New Page" onclick="cmdAddNewPage_onclick()" /></td>
                    <td> <%
                        if (len(PageID) > 0) then %>
                            <input type="button" value="Add New Div to Page ID : <%=PageID%>" onclick="cmdAddNewDiv_onclick(<%=PageID%>)" /> <%
                        end if %>
                    </td>
                </tr> 
                <tr>
                    <td colspan="3">
                        <table class="table" style="margin-left:5px; width:98%">
                             <tr class="Header">
                                <td style="width:8%;">Page ID</td>
                                <td style="width:20%;">Page Name</td>
                                <td style="width:15%;">Page Tag</td>
                                <td style="width:30%;">Image Name</td>
                                <td style="width:27%;"></td>
                            </tr> <%
                            
                            row = 1
                            Do While not objRS_Db.EOF 
                                '** Selected **
                                if PageID = trim(objRS_db.Fields("PageID")) then %>
                                    <tr>
                                        <td>
                                            &nbsp;
                                            <input type="hidden" id="selPageName" name="selPageName" value="<%=trim(objRS_db.Fields("PageName"))%>" />
                                            <input type="hidden" id="selPageTag" name="selPageTag" value="<%=trim(objRS_db.Fields("pagetag"))%>" />
                                            <input type="hidden" id="selImageName" name="selImageName" value="<%=trim(objRS_db.Fields("ImageName"))%>" />
                                        </td>
                                    </tr>
                                    <tr class="Selected">
                                        <td>&nbsp;&nbsp;<%=objRS_db.Fields("PageID") %></td>
                                        <td><input type="text" id="dbPageName" name="dbPageName" style="width:200px" value="<%=trim(objRS_db.Fields("PageName"))%>" /></td>
                                        <td><input type="text" id="dbpagetag" name="dbpagetag" style="width:100px" value="<%=trim(objRS_db.Fields("pagetag"))%>" /></td>
                                        <td><input type="text" id="dbImageName" name="dbImageName" style="width:250px" value="<%=trim(objRS_db.Fields("ImageName"))%>" /></td>
                                        <td>
                                            <a href="javascript: CreateJSON()"><img src="images/json.png" /></a>
                                            <a href="javascript: CreateHTML()"><img src="images/html.png" /></a>
                                             &nbsp;&nbsp;
                                            <a href="javascript: PushPageDB2Build(<%=objRS_db.Fields("PageID") %>)"><img src="images/up.png" /></a>
                                            <a href="javascript: UpdatePage(<%=objRS_db.Fields("PageID") %>)"><img src="images/save.png" /></a>
                                            <a href="javascript: ClosePage()"><img src="images/close.png" /></a>
                                             &nbsp;&nbsp;
                                            <a href="javascript: DeletePage(<%=objRS_db.Fields("PageID") %>)"><img src="images/delete.png" /></a>
                                        </td>
                                    </tr> 
                                    <tr>
                                        <td colspan="5">
                                            <table style="border:2px solid; border-color:black; background-color:antiquewhite; width:95%; margin-left:5px;">
                                                <tr class="DetailTable" style="font-weight:bold">
                                                    <td style="width:5%; text-align:center;">Div ID</td>
                                                    <td style="width:10%; text-align:center;">TabIndex</td>
                                                    <td style="width:10%">Div Name</td>
                                                    <td style="width:65%">Div Tag</td>
                                                    <td style="width:10%"></td>
                                                </tr> <%
                                                strRS = "SELECT divID, PageID, DivTag, tabindex, DivName, toolposition, toolwidth, customleft, " &_
                                                                "customtop, title, CurX, CurY, Pos1_X, Pos1_Y, Pos2_X, Pos2_Y, Style " &_
                                                        "FROM tblDiv " &_
                                                        "WHERE " &_
                                                            "(active = 1) AND " &_
                                                            "(PageID = '" & trim(PageID) & "') " &_
                                                        "ORDER BY tabindex "
                                                'response.Write("strRS:" & strRS & "<br>") 
                                                objRS_db2.Open(strRS), objDC 
                                                rowd = 1

                                                if objRS_db2.eof then %>
                                                    <tr>
                                                        <td colspan="5" style="color:red; text-align:center; font-weight:bold"><br />No Div Tags Defined<br /><br /></td>
                                                    </tr> <%
                                                end if
                                                     
                                                Do While not objRS_Db2.EOF %>
                                                    <tr <%if (rowd mod 2) = 0 Then %>class="evend"<%else%>class="oddd"<%end if%>>
                                                        <td style="text-align:center"><%=trim(objRS_db2.Fields("divID"))%></td>
                                                        <td style="text-align:center"><%=trim(objRS_db2.Fields("tabindex"))%></tdstyle="text-align:center>
                                                        <td><%=trim(objRS_db2.Fields("DivName"))%></td>
                                                        <td class="divtag"><%=trim(objRS_db2.Fields("DivTag"))%></td>
                                                        <td style="vertical-align:middle;">
                                                            <a href="javascript: PushDivDB2Build(<%=objRS_db2.Fields("DivID")%>,<%=objRS_db.Fields("PageID")%>)"><img src="images/up.png" /></a>
                                                            &nbsp;&nbsp;&nbsp;
                                                            <a href="javascript: DeleteDiv(<%=objRS_db2.Fields("DivID")%>,<%=objRS_db.Fields("PageID")%>)"><img src="images/delete.png" /></a>
                                                            <!--**** Selected Hidden Variables **** -->
                                                            <input type="hidden" id="selTabindex<%=rowd%>" name="selTabindex<%=rowd%>" value="<%=trim(objRS_db2.Fields("tabindex"))%>" />
                                                            <input type="hidden" id="selObjName<%=rowd%>" name="selObjName<%=rowd%>" value="<%=trim(objRS_db2.Fields("DivName"))%>" />
                                                            <input type="hidden" id="selToolPosition<%=rowd%>" name="selToolPosition<%=rowd%>" value="<%=trim(objRS_db2.Fields("ToolPosition"))%>" />
                                                            <input type="hidden" id="selToolWidth<%=rowd%>" name="selToolWidth<%=rowd%>" value="<%=trim(objRS_db2.Fields("ToolWidth"))%>" />
                                                            <input type="hidden" id="selCustomLeft<%=rowd%>" name="selCustomLeft<%=rowd%>" value="<%=trim(objRS_db2.Fields("CustomLeft"))%>" />
                                                            <input type="hidden" id="selCustomTop<%=rowd%>" name="selCustomTop<%=rowd%>" value="<%=trim(objRS_db2.Fields("CustomTop"))%>" />
                                                            <input type="hidden" id="selTitle<%=rowd%>" name="selTitle<%=rowd%>" value="<%=trim(objRS_db2.Fields("Title"))%>" />
                                                            <input type="hidden" id="selCurX<%=rowd%>" name="selCurX<%=rowd%>" value="<%=trim(objRS_db2.Fields("CurX"))%>" />
                                                            <input type="hidden" id="selCurY<%=rowd%>" name="selCurY<%=rowd%>" value="<%=trim(objRS_db2.Fields("CurY"))%>" />
                                                            <input type="hidden" id="selPos1_X<%=rowd%>" name="selPos1_X<%=rowd%>" value="<%=trim(objRS_db2.Fields("Pos1_X"))%>" />
                                                            <input type="hidden" id="selPos1_Y<%=rowd%>" name="selPos1_Y<%=rowd%>" value="<%=trim(objRS_db2.Fields("Pos1_Y"))%>" />
                                                            <input type="hidden" id="selPos2_X<%=rowd%>" name="selPos2_X<%=rowd%>" value="<%=trim(objRS_db2.Fields("Pos2_X"))%>" />
                                                            <input type="hidden" id="selPos2_Y<%=rowd%>" name="selPos2_Y<%=rowd%>" value="<%=trim(objRS_db2.Fields("Pos2_Y"))%>" />
                                                            <input type="hidden" id="selStyle<%=rowd%>" name="selStyle<%=rowd%>" value="<%=trim(objRS_db2.Fields("Style"))%>" />
                                                        </td>
                                                    </tr> <%
                                                    objRS_db2.MoveNext
                                                    rowd = rowd + 1
                                                Loop %>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            &nbsp;<input type="hidden" id="selNumDiv" name="selNumDiv" value="<%=rowd - 1%>" />
                                        </td>
                                    </tr><%
                                '** Not Selected **
                                else %>
                                    <tr <%if (row mod 2) = 0 Then %>class="odd"<%else%>class="even"<%end if%>>
                                        <td>&nbsp;&nbsp;<%=objRS_db.Fields("PageID") %></td>
                                        <td><%=objRS_db.Fields("PageName") %></td>
                                        <td><%=objRS_db.Fields("pagetag") %></td>
                                        <td><%=objRS_db.Fields("ImageName") %></td>
                                        <td style="text-align:right">
                                            <a href="javascript: OpenPage(<%=objRS_db.Fields("PageID") %>)"><img src="images/open.png" /></a>
                                            &nbsp;&nbsp;&nbsp;
                                            <a href="javascript: DeletePage(<%=objRS_db.Fields("PageID") %>)"><img src="images/delete.png" /></a>
                                        </td>
                                    </tr> <%
                                end if
                                row = row + 1
                                objRS_db.MoveNext 
                            Loop %>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td></td>
                </tr>
               
            </table>
        </div>
        </div>

        <input type="hidden" id="hdnDivTag" name="hdnDivTag" value="">
        <input type="hidden" id="hdnAction" name="hdnAction" value="">  
        <input type="hidden" id="hdnLastX" name="hdnLastX" value="">  
        <input type="hidden" id="hdnLastY" name="hdnLastY" value="">  
        <input type="hidden" id="hdnPageID" name="hdnPageID" value="<%=PageID%>">
        <input type="hidden" id="hdnDivID" name="hdnDivID" value="<%=DivID%>"> 
        <input type="hidden" id="hdnDivStyle" name="hdnDivStyle" value=""> 
        
    </form>


    <script type="text/javascript" src="js/tooltip-mobile.js"></script>

</body>
</html>