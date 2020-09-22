<div align="center">

## Office Web Chart


</div>

### Description

This code will create an Office web chart from the data supplied from a database. It will then create a gif file and place it on a web page.
 
### More Info
 
You need a database with data in it that you want to plot

You need office web components installed on the server


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Billy Yawn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/billy-yawn.md)
**Level**          |Beginner
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/billy-yawn-office-web-chart__4-6429/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
</HEAD>
<%
dim sql
'SQL Statment to select data from a database. This will be bound to a chart.
sql = "SELECT Date, [Open], High, Low, Last, [Day Volume], OBV from symbol"
'Create ADO connection, recordset and command object for database
Set cnn = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.Recordset")
Set cmd = Server.CreateObject("ADODB.Command")
cnn.Open "DSN=Stocks;uid=sa;pwd=;"
Set cmd.ActiveConnection = cnn
cmd.CommandText = sql
rst.CursorLocation = adUseClient
rst.Open cmd, , adOpenStatic, adLockBatchOptimistic
'Define a chartspace and create an Office Web Chart object
dim oChartSpace, oConst
set oChartSpace = server.CreateObject("OWC.Chart")
set oConst = oChartSpace.Constants
With oChartSpace
    .Clear
    .Border.Color = oConst.chColorNone
    .Interior.Color = "gainsboro"
    .HasChartSpaceTitle = True
    .ChartSpaceTitle.Caption = "Open - High - Low - Close"
    .ChartSpaceTitle.Font.Name = "Tahoma"
    .ChartSpaceTitle.Font.Size = 12
    .ChartSpaceTitle.Font.Bold = True
  End With
'Bind database recordset to chart datasource
set oChartSpace.DataSource = rst
'Define first chart in chartspace
dim oChart1
set oChart1 = oChartSpace.Charts.Add
dim oSeries
set oSeries = oChart1.SeriesCollection.Add
    'Specify the Marker style of the Series
    oChartSpace.Charts(0).SeriesCollection(0).Marker.Style = oConst.chMarkerStyleNone
    oChartSpace.Charts(0).SeriesCollection(0).Marker.Size = 1
    'A OHLC chart is a stock Open - High - Low - Close chart
    oChartSpace.Charts(0).Type = oConst.chChartTypeStockOHLC
  'Add the dates and OHLC values
  With oChartSpace.Charts(0).SeriesCollection(0)
    .Caption = "Daily"
    .SetData oConst.chDimCategories, 0, 0 'Dates item 0 from sql statement
    .SetData oConst.chDimOpenValues, 0, 1 'Open item 1 from sql statement
    .SetData oConst.chDimHighValues, 0, 2 'High item 2 from sql statement
    .SetData oConst.chDimLowValues, 0, 3 'Low item 3 from sql statement
    .SetData oConst.chDimCloseValues, 0, 4 'Close item 4 from sql statement
  End With
  oChartSpace.Charts(0).Axes(oConst.chAxisPositionBottom).HasMajorGridlines = true
  'Remove the ticklabels
  oChart1.Axes(oConst.chAxisPositionBottom).HasTickLabels = False
  '--- Create a Column Chart for Volume ---
  'Add a second chart to the Chartspace for the Volume
  Set oChart2 = oChartSpace.Charts.Add
  oChart2.Type = oConst.chChartTypeColumnClustered
  oChart2.SetData oConst.chDimCategories, 0, 0 'Dates item 0 from sql statement
  oChart2.SetData oConst.chDimValues, 0, 5 'Volume item 5 from sql statement
  oChart2.SeriesCollection(0).Caption = "Volume"
  Set oChart3 = oChartSpace.Charts.Add
  oChart3.Type = oConst.chChartTypeLine
  oChart3.SetData oConst.chDimCategories, 0, 0 'Dates item 0 from sql statement
  oChart3.SetData oConst.chDimValues, 0, 6 'OBV item 6 from sql statement
  oChart3.SeriesCollection(0).Caption = "On Balance Volume"
  '--- Apply Formatting Common to Both Charts ---
  'Make the HLC chart twice as large as the Volume chart
  oChart1.HeightRatio = 200
  oChart2.HeightRatio = 100
  oChart3.HeightRatio = 100
  For each oCht in oChartSpace.Charts
   'Display the legend to the Right of the Chart, remove the legend border and change the color
   oCht.HasLegend = True
   oCht.Legend.Position = oConst.chLegendPositionRight
   oCht.Legend.Border.Color = oConst.chColorNone
   oCht.Legend.Interior.Color = "gainsboro"
   'Remove tick labels for both axes
   oCht.Axes(oConst.chAxisPositionBottom).MajorTickmarks = oConst.chTickMarkNone
   oCht.Axes(oConst.chAxisPositionLeft).MajorTickmarks = oConst.chTickMarkNone
   'Change the weight and color of the axes lines
   oCht.Axes(oConst.chAxisPositionBottom).Line.Color = "SLATEGRAY"
   oCht.Axes(oConst.chAxisPositionLeft).Line.Color = "SLATEGRAY"
   oCht.Axes(oConst.chAxisPositionBottom).Line.Weight = oConst.owcLineWeightMedium
   oCht.Axes(oConst.chAxisPositionLeft).Line.Weight = oConst.owcLineWeightMedium
   'Display the category axis gridlines
   oCht.Axes(oConst.chAxisPositionBottom).HasMajorGridlines = true
   'Change gridlines color
   oCht.Axes(oConst.chAxisPositionLeft).MajorGridlines.Line.Color = "BURLYWOOD"
   oCht.Axes(oConst.chAxisPositionBottom).MajorGridlines.Line.Color = "BURLYWOOD"
   'Change color of all series
   For each oSeries in oCht.SeriesCollection
     oSeries.Interior.Color = "SLATEGRAY"
     oSeries.Border.Color = "SLATEGRAY"
     oSeries.Line.Color = "SLATEGRAY"
   Next
   'Change plotarea color
   oCht.PlotArea.Interior.Color = "BISQUE"
   'Change Font style for the axis and the legends
   oCht.Axes(oConst.chAxisPositionBottom).Font.Name = "Tahoma"
   oCht.Axes(oConst.chAxisPositionLeft).Font.Name = "Tahoma"
   oCht.Legend.Font.Name = "Tahoma"
  Next
'This exports a gif from the chart object so the remote user does not need office to see the chart
'Use the SessionID and file count to determine a unique filename
sFilename = Session.SessionID & "_" & Session("FSO").GetTempName
'Create a GIF image of the Chart
oChartSpace.ExportPicture Server.MapPath(sFilename), "GIF", 700, 500
'Store the filename in the Session object and increment the file count variable
Session ("GIF" & Session("nGIFCount")) = Server.MapPath(".") & "\" & sFileName
Session("nGIFCount") = Session("nGIFCount") + 1
rst.Close
cnn.Close
Set oChartSpace = Nothing
Set rst = Nothing
Set cnn = Nothing
%>
<BODY BGCOLOR=WHITE BORDER=0 CELLSPACING=30 CELLPADDING=30>
<TABLE WIDTH="90%" ALIGN=CENTER>
<TR>
<TD COLSPAN=2 VALIGN=CENTER BGCOLOR=INDIANRED><FONT SIZE=5 FACE="Tahoma" COLOR=LINEN><B>Chart</B><BR></FONT></TD>
</TR>
<TR>
<TD BGCOLOR=LINEN VALIGN=TOP WIDTH=200 ALIGN=CENTER>
<BR><BR><BR>
<FONT FACE="Tahoma">
<FONT SIZE=3>
</FONT>
<BR><HR COLOR=INDIANRED WIDTH="75%"><BR>
</TD>
<TD ALIGN=CENTER VALIGN=TOP>
<IMG SRC='<%=sFilename%>'>
</TD>
</TR>
</TABLE>
<BR>
<BR>
</BODY>
</HTML>
You also need to add the following to your global.asa file
Sub Session_OnStart
	Set Session("FSO") = CreateObject("Scripting.FileSystemObject")
	Session("n")=0
	Session.Timeout=1
End Sub
Sub Session_OnEnd
	dim x
	for x = 0 to Session("n") - 1
		Session("FSO").DeleteFile Session("sTempFile" & x), True
	next
end sub
```

