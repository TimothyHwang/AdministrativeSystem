<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MOA10006.aspx.vb" Inherits="M_Source_10_MOA10006" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>主官在營檢視表</title>    
    
    <link href='<%#ResolveUrl("~/css/marquee/jquery.marquee.css") %>' rel="stylesheet" type="text/css" title="default" media="all"/>
    <script src='<%#ResolveUrl("~/script/jquery-1.10.2.min.js") %>' type="text/javascript"></script>
    <script src='<%#ResolveUrl("~/script/marquee/jquery.marquee.js") %>' type="text/javascript"></script>

    <style type ="text/css">
        body {
            
            background-color: #006;
            margin-top: 0;
            margin-left: 0;
            margin-right: 0;
            margin-bottom: 0;
            
        }
                  
        td.nametd {
            font-family: 華康中圓體;            
            color: #FF0;            
            width:100%;
            margin-top: 15px;
            font-size: 269%;
            word-wrap: break-word;
            line-height: 150%;
            text-align: justify;
            text-justify:inter-ideograph;
            /*
            font-size: 230%;
            
            text-align:center;
            vertical-align:middle;
            
            text-align: justify;            
            
            -ms-text-align-last:center;
            
            
            
            
            -webkit-text-justify:distribute-all-lines;
            -moz-text-justify:distribute-all-lines;
            -o-text-justify:distribute-all-lines;            
            -webkit-text-align-last:justify;
            -moz-text-align-last:justify;
            -o-text-align-last:justify;            
            height: 50px;*/
        }
        td.nametd:after 
        {            
            content: "";
            display: inline-block;
            width: 100%;    
        }
        
        td.imgtd {
            border: 1px solid #fff;
        }       
    </style>
    <script language="javascript" type="text/javascript">
        function resize() {
            window.moveTo(0, 0);
            window.resizeTo(screen.width, screen.height);
        }

        function toggleFullScreen() {
            var docElement = $('#StatusList');
            if (docElement != null) {
                if (!document.mozFullScreen && !document.webkitFullScreen) {
                    if (docElement.mozRequestFullScreen) {
                        docElement.mozRequestFullScreen();
                    } else {
                        docElement.webkitRequestFullScreen(Element.ALLOW_KEYBOARD_INPUT);
                    }
                } else {
                    if (document.mozCancelFullScreen) {
                        document.mozCancelFullScreen();
                    } else {
                        document.webkitCancelFullScreen();
                    }
                }
            }
        }        
    </script>
</head>
<body onload="resize();">
    <form id="form1" runat="server">
    <div id="StatusList">
        <asp:ScriptManager ID="sm" runat="server">
            
        </asp:ScriptManager>
        <div style="height:3px">&nbsp;</div>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <asp:Timer ID="Refresh" runat="server" Interval="5000">
                </asp:Timer>
                <div id="render" align="center">
                <asp:PlaceHolder ID="phrStatus" runat="server"></asp:PlaceHolder>        

                </div>
            </ContentTemplate>
        </asp:UpdatePanel>
        
                <div align="left">                      
                <ul id="marquee" class="marquee">
                    <asp:Literal ID="lblMarqueeMsg" runat="server" Text="Label"></asp:Literal>
                </ul>
                </div>
    </div>
    
    </form>
    <script language="javascript" type="text/javascript">
    //*************** marquee start *****************
        // on DOM ready
        $(document).ready(function () {
            $("#marquee").marquee({
                loop: -1
                // this callback runs when the marquee is initialized
			, init: function ($marquee, options) {
			    //window.debug("init", arguments);

			    // shows how we can change the options at runtime
			    options.yScroll = "bottom";
			}                
                // this callback runs when a has fully scrolled into view (from either top or bottom)
			, show: function () {
			    //window.debug("show", arguments);
			}
            });
        });

        var iNewMessageCount = 0;

        function addMessage(selector) {
            // increase counter
            iNewMessageCount++;

            // append a new message to the marquee scrolling list
            var $ul = $(selector).append("<li>New message #" + iNewMessageCount + "</li>");
            // update the marquee
            $ul.marquee("update");
        }

        function pause(selector) {
            $(selector).marquee('pause');
        }

        function resume(selector) {
            $(selector).marquee('resume');
        }
        //*************** marquee end *****************
        
 function myrefresh()
 {
      window.location.reload();
 }
 setTimeout('myrefresh()',60000); //指定1秒刷新一次
</script>
    
</body>
</html>
