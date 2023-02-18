<%
'the pie chart will throw an exception if an error occurs so enable inline error trapping
On Error Resume Next

strPath = Server.MapPath(".")  'retrieve the physical directory where ASP pages are located

Set objPieChart = Server.CreateObject("Dundas.PieChartServer.2")  'create an instance of the control

objPieChart.DirTemplate = strPath & "\Templates\"     'set the Template directory of the control objPieChart.DirTexture = strPath & "\Textures\"          'set the Textures directory of the control



'	rt = objPieChart.LoadTemplate("Textures.cuc")
'	rt = objPieChart.LoadTemplate("HumblePie.cuc")
'	rt = objPieChart.LoadTemplate("PicklePie.cuc")	
'	rt = objPieChart.LoadTemplate("PieInTheSky.cuc")
'	rt = objPieChart.LoadTemplate("Marketshare2.cuc")
'	rt = objPieChart.LoadTemplate("Portfoliomix.cuc")

'	rt = objPieChart.LoadTemplate("Texture.cuc")
	rt = objPieChart.LoadTemplate("Gustavo3.cuc")
'	rt = objPieChart.LoadTemplate("Colors.cuc")
' rt  =  objPieChart.LoadTemplate("Poll.cuc")     

'load a template (made in the Template Editor)

 'add 3 slices to the pie chart and specify values (sizes) and slice labels
 dim cont
 dim fin
 dim total
 dim descripcion
 dim votos
 
 votos  = session("total")
 descripcion = session("descripcion")
 fin = session("cont")
 cont = 0
 
 for cont=0 to fin 
    total = total + votos(cont)
 next 
 
 for cont= 0 to fin
' total = session("total")
   
     porcentaje = (votos(cont) / total) * 100
     if porcentaje <> "100" then
	 Mypos = InStr(1, porcentaje, ".", 1)
	 porcentaje = Left(porcentaje, Mypos + 2)   
	 end if
  
  	rt = objPieChart.AddData(votos(cont),porcentaje&"%")
  	'Response.Write total(cont)
  	
 next 

'now set the legend entry text
for cont= 0 to fin
rt = objPieChart.AddLabel(descripcion(cont))
next 
'output graphics as a jpeg file, specifying the width and height of the image
rt = objPieChart.SendJPEG(250,300,5)
%>



