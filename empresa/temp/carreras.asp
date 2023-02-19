<%@ Language=VBScript %>
<%
dim textXML

  textxml = textxml + "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
  textxml = textxml + "<channel>" 
  textxml = textxml + "<item>" 
  textxml = textxml + "<title>USAC</title>" 
  textxml = textXML + "<link>http://sistemas.ingenieria-usac.edu.gt</link>" 
  textxml = textXML + "<description>Ingenieria en ciencias y Sistemas</description>"
  textxml = textXML + "<pubDate>2004-04-02T14:02:18-08:00</pubDate>"
  textxml = textXML + "<content:encoded>"
  textxml = textXML + "<![CDATA[<TABLE BORDER=0 WIDTH=100><TR><TD><table border='0' width='100%' cellspacing='0' cellpadding='0'> "
  textxml = textXML + "<tr valign='top' align='left'> "
  textxml = textXML + "<td>gika</td>"  
  textxml = textXML + "</tr>"
  textxml = textXML + "</table></TD></TR>"
  textxml = textXML + "</TABLE>]]></content:encoded>"
  textxml = textXML + "<itms:artist>D12</itms:artist>"
  textxml = textXML + "<itms:album>My Band - Single</itms:album>"
  textxml = textXML + "<itms:albumPrice>By Song Only</itms:albumPrice>                        "
  textxml = textXML + "<itms:song>My Band</itms:song>"
  textxml = textXML + "<itms:songLink>http://phobos.apple.com/WebObjects/MZStore.woa/wa/viewAlbum?selectedItemId=6063159&amp;playListId=6063161</itms:songLink>"
  textxml = textXML + "<itms:rank>1</itms:rank>"
  textxml = textXML + "<itms:rights>2004 Shady Records</itms:rights>"
  textxml = textXML + "<itms:releasedate>03/11/2004</itms:releasedate>"
  textxml = textXML + "</item>"
  textxml = textxml + "</channel>" 
  Response.Write textxml
  Response.end

%>