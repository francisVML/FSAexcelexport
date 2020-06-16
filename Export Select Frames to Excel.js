

// create an instance of Excel and show it
excel = new ActiveXObject("Excel.Application");
excel.Visible = true;

// create a workbook with a single sheet named "FSA Export"
book = excel.Workbooks.Add();

sheets = book.Sheets;
excel.DisplayAlerts = false;



recordNumber = application.View.RecordingIndex;

frames = application.Document.Recordings(recordNumber).Frames;
row = 3;
count = frames.Count;
selectedFrames = 0;
averageReading = new ActiveXObject("FSA4.Reading");

//no_of_sensors = frames(0).Readings.Count;

regions = application.View.RecordingView.Panel("readings").Panel(0).Regions;

//Number of Frames you want to extract, for 5min intervals is 144, for 1min interval is 60
TotalFrames = 12;

no_of_regions = regions.count;

//Every 5 minutes
var a = [500,2000,3500,5000,6500,8000,9500,11000,12500,14000,15500,17000,18500,20000,21500,23000,24500,26000,27500,29000,30500,32000,33500,35000,36500,38000,39500,41000,42500,44000,45500,47000,48500,50000,51500,53000,54500,56000,57500,59000,60500,62000,63500,65000,66500,68000,69500,71000,72500,74000,75500,77000,78500,80000,81500,83000,84500,86000,87500,89000,90500,92000,93500,95000,96500,98000,99500,101000,102500,104000,105500,107000,108500,110000,111500,113000,114500,116000,117500,119000,120500,122000,123500,125000,126500,128000,129500,131000,132500,134000,135500,137000,138500,140000,141500,143000,144500,146000,147500,149000,150500,152000,153500,155000,156500,158000,159500,161000,162500,164000,165500,167000,168500,170000,171500,173000,174500,176000,177500,179000,180500,182000,183500,185000,186500,188000,189500,191000,192500,194000,195500,197000,198500,200000,201500,203000,204500,206000,207500,209000,210500,212000,213500,215000];

//Every 1 minute
//var a = [500,800,1100,1400,1700,2000,2300,2600,2900,3200,3500,3800,4100,4400,4700,5000,5300,5600,5900,6200,6500,6800,7100,7400,7700,8000,8300,8600,8900,9200,9500,9800,10100,10400,10700,11000,11300,11600,11900,12200,12500,12800,13100,13400,13700,14000,14300,14600,14900,15200,15500,15800,16100,16400,16700,17000,17300,17600,17900,18200];

//Every 5 seconds, 720 frames for 1hr
//var a = [5,30,55,80,105,130,155,180,205,230,255,280,305,330,355,380,405,430,455,480,505,530,555,580,605,630,655,680,705,730,755,780,805,830,855,880,905,930,955,980,1005,1030,1055,1080,1105,1130,1155,1180,1205,1230,1255,1280,1305,1330,1355,1380,1405,1430,1455,1480,1505,1530,1555,1580,1605,1630,1655,1680,1705,1730,1755,1780,1805,1830,1855,1880,1905,1930,1955,1980,2005,2030,2055,2080,2105,2130,2155,2180,2205,2230,2255,2280,2305,2330,2355,2380,2405,2430,2455,2480,2505,2530,2555,2580,2605,2630,2655,2680,2705,2730,2755,2780,2805,2830,2855,2880,2905,2930,2955,2980,3005,3030,3055,3080,3105,3130,3155,3180,3205,3230,3255,3280,3305,3330,3355,3380,3405,3430,3455,3480,3505,3530,3555,3580,3605,3630,3655,3680,3705,3730,3755,3780,3805,3830,3855,3880,3905,3930,3955,3980,4005,4030,4055,4080,4105,4130,4155,4180,4205,4230,4255,4280,4305,4330,4355,4380,4405,4430,4455,4480,4505,4530,4555,4580,4605,4630,4655,4680,4705,4730,4755,4780,4805,4830,4855,4880,4905,4930,4955,4980,5005,5030,5055,5080,5105,5130,5155,5180,5205,5230,5255,5280,5305,5330,5355,5380,5405,5430,5455,5480,5505,5530,5555,5580,5605,5630,5655,5680,5705,5730,5755,5780,5805,5830,5855,5880,5905,5930,5955,5980,6005,6030,6055,6080,6105,6130,6155,6180,6205,6230,6255,6280,6305,6330,6355,6380,6405,6430,6455,6480,6505,6530,6555,6580,6605,6630,6655,6680,6705,6730,6755,6780,6805,6830,6855,6880,6905,6930,6955,6980,7005,7030,7055,7080,7105,7130,7155,7180,7205,7230,7255,7280,7305,7330,7355,7380,7405,7430,7455,7480,7505,7530,7555,7580,7605,7630,7655,7680,7705,7730,7755,7780,7805,7830,7855,7880,7905,7930,7955,7980,8005,8030,8055,8080,8105,8130,8155,8180,8205,8230,8255,8280,8305,8330,8355,8380,8405,8430,8455,8480,8505,8530,8555,8580,8605,8630,8655,8680,8705,8730,8755,8780,8805,8830,8855,8880,8905,8930,8955,8980,9005,9030,9055,9080,9105,9130,9155,9180,9205,9230,9255,9280,9305,9330,9355,9380,9405,9430,9455,9480,9505,9530,9555,9580,9605,9630,9655,9680,9705,9730,9755,9780,9805,9830,9855,9880,9905,9930,9955,9980,10005,10030,10055,10080,10105,10130,10155,10180,10205,10230,10255,10280,10305,10330,10355,10380,10405,10430,10455,10480,10505,10530,10555,10580,10605,10630,10655,10680,10705,10730,10755,10780,10805,10830,10855,10880,10905,10930,10955,10980,11005,11030,11055,11080,11105,11130,11155,11180,11205,11230,11255,11280,11305,11330,11355,11380,11405,11430,11455,11480,11505,11530,11555,11580,11605,11630,11655,11680,11705,11730,11755,11780,11805,11830,11855,11880,11905,11930,11955,11980,12005,12030,12055,12080,12105,12130,12155,12180,12205,12230,12255,12280,12305,12330,12355,12380,12405,12430,12455,12480,12505,12530,12555,12580,12605,12630,12655,12680,12705,12730,12755,12780,12805,12830,12855,12880,12905,12930,12955,12980,13005,13030,13055,13080,13105,13130,13155,13180,13205,13230,13255,13280,13305,13330,13355,13380,13405,13430,13455,13480,13505,13530,13555,13580,13605,13630,13655,13680,13705,13730,13755,13780,13805,13830,13855,13880,13905,13930,13955,13980,14005,14030,14055,14080,14105,14130,14155,14180,14205,14230,14255,14280,14305,14330,14355,14380,14405,14430,14455,14480,14505,14530,14555,14580,14605,14630,14655,14680,14705,14730,14755,14780,14805,14830,14855,14880,14905,14930,14955,14980,15005,15030,15055,15080,15105,15130,15155,15180,15205,15230,15255,15280,15305,15330,15355,15380,15405,15430,15455,15480,15505,15530,15555,15580,15605,15630,15655,15680,15705,15730,15755,15780,15805,15830,15855,15880,15905,15930,15955,15980,16005,16030,16055,16080,16105,16130,16155,16180,16205,16230,16255,16280,16305,16330,16355,16380,16405,16430,16455,16480,16505,16530,16555,16580,16605,16630,16655,16680,16705,16730,16755,16780,16805,16830,16855,16880,16905,16930,16955,16980,17005,17030,17055,17080,17105,17130,17155,17180,17205,17230,17255,17280,17305,17330,17355,17380,17405,17430,17455,17480,17505,17530,17555,17580,17605,17630,17655,17680,17705,17730,17755,17780,17805,17830,17855,17880,17905,17930,17955,17980,18005]

if(no_of_regions > 0){
for(q=0; q<no_of_regions; q++){

excel.DisplayAlerts = true;
excel.StatusBar = "Filling cells with FSA data...";

		view = application.View;
		recordingView = view.RecordingView.Panel("readings");
		stats = recordingView.Panel(0).Statistics

//sheet = book.Sheets(q+1);
sheet = sheets.Add();

sheet.Name="Region # "+(q+1);

cells = sheet.Cells;

sheet.Range(cells(1, 1), cells(1, 1)).ColumnWidth = 24;
sheet.Range(cells(1, 2), cells(1, 100)).ColumnWidth = 15;
		
 cells(3, 1).Value = "Frame Index";
 cells(4, 1).Value = "Time";
 cells(5, 1).Value = stats.StatisticName(0) + " (" + stats.StatisticUnits(0) +")" ;
 cells(6, 1).Value = stats.StatisticName(1) + " (" + stats.StatisticUnits(1) +")" ;
 cells(7, 1).Value = stats.StatisticName(2) + " (" + stats.StatisticUnits(2) +")" ;
 cells(8, 1).Value = stats.StatisticName(3) + " (" + stats.StatisticUnits(3) +")" ;
 cells(9, 1).Value = stats.StatisticName(4) + " (" + stats.StatisticUnits(4) +")" ;
 cells(10, 1).Value = stats.StatisticName(5) + " (" + stats.StatisticUnits(5) +")" ;
 cells(11, 1).Value = stats.StatisticName(6) + " (" + stats.StatisticUnits(6) +")" ;
 cells(12, 1).Value = stats.StatisticName(7) + " (" + stats.StatisticUnits(7) +")" ;
 cells(13, 1).Value = stats.StatisticName(8) + " (" + stats.StatisticUnits(8) +")" ;
 cells(14, 1).Value = stats.StatisticName(9) + " (" + stats.StatisticUnits(9) +")" ;


columns = frames(0).Readings(0).Columns;
rows = frames(0).Readings(0).Rows;

	c1=regions(q).StartColumn;
	r1=regions(q).StartRow;
	width = regions(q).Columns;
	height = regions(q).Rows;
	c2=c1+width;
	r2=r1+height;	


		cnt=0;
		var character;
		var asc = 65;
		for (c = 0; c < columns; ++c){
			if(c1-1>=c || c>=c2) continue;
			for (r = 0; r < rows; ++r){
				if(r1>r || r>=r2) continue;
				cnt++;
				if (c<26){
					character = String.fromCharCode(65 + c);
				} 
				else {
					character = String.fromCharCode(97 + (c-26));
				}				
				cells(cnt+15,1).Value = character + (r+1)
			}
		}

var oldcopx = -1;
var oldcopy = -1;	
var sensors = new Array();
col=1;

for (f = 0; f < count; ++f)
	if (frames(f).Selected)
	{
		col++;
		view = application.View;
		recordingView = view.RecordingView.Panel("readings");
		stats = recordingView.Panel(0).Statistics;
		stats.FrameIndex = f;
		stats.Recalculate(2);
		cells(3,col).Value = f+1;
		//cells(4,col).Value = mydateFunction(stats.StatisticTime(0));
		cells(4,col).Value = frames(f).TimeString;

		
		i = 5;
		for(w=0; w<10;w++){
		cells(i+w, col).Value = stats.StatisticValue(0, w, q).toFixed(2);

			if(w==6){
				newcopx = stats.StatisticValue(0, w, q).toFixed(2);
			}
			if(w==7){
				newcopy = stats.StatisticValue(0, w, q).toFixed(2);
			}
		}
			if(oldcopx != -1 && oldcopy !=-1){
			//	cells(15,col).Value = Math.sqrt((Math.pow(newcopx - oldcopx,2) + Math.pow(newcopy - oldcopy,2))).toFixed(2);
			//	cells(16,col).Value = ((Math.atan((newcopy - oldcopy)/(newcopx - oldcopx)))*180 / Math.PI).toFixed(2);					
			}
				oldcopx = newcopx;
				oldcopy = newcopy;

		// copy the values into the spreadsheet

		sheet.EnableCalculation = false;
		reading = frames(f).Readings(0);
		cnt=0;
		//reading = frames(f).regions(q);
		for (c = 0; c < columns; ++c)
		{
			if(!(c1<=c && c<c2)) continue;
			//if(c1-1>=c || c>=c2) continue;
			for (r = 0; r < rows; ++r)
			{
				if(!(r1<=r && r<r2)) continue;
				
				//sensors.push(reading.Value(r, c));
				cnt++;
				cells(cnt+15, col).Value = reading.Value(c, r);

			}
		}
		sheet.EnableCalculation = true;
		selectedFrames++;
		
	}

} //end of for loof	

}else{

excel.DisplayAlerts = true;
excel.StatusBar = "Filling cells with FSA data...";

		view = application.View;
		recordingView = view.RecordingView.Panel("readings");
		stats = recordingView.Panel(0).Statistics

sheet = book.Sheets(1);

//sheet.Name="Region # "+(1);

cells = sheet.Cells;

sheet.Range(cells(1, 1), cells(1, 1)).ColumnWidth = 24;
sheet.Range(cells(1, 2), cells(1, 100)).ColumnWidth = 15;
		
 cells(3, 1).Value = "Frame Index";
 cells(4, 1).Value = "Time";
 cells(5, 1).Value = stats.StatisticName(0) + " (" + stats.StatisticUnits(0) +")" ;
 cells(6, 1).Value = stats.StatisticName(1) + " (" + stats.StatisticUnits(1) +")" ;
 cells(7, 1).Value = stats.StatisticName(2) + " (" + stats.StatisticUnits(2) +")" ;
 cells(8, 1).Value = stats.StatisticName(3) + " (" + stats.StatisticUnits(3) +")" ;
 cells(9, 1).Value = stats.StatisticName(4) + " (" + stats.StatisticUnits(4) +")" ;
 cells(10, 1).Value = stats.StatisticName(5) + " (" + stats.StatisticUnits(5) +")" ;
 cells(11, 1).Value = stats.StatisticName(6) + " (" + stats.StatisticUnits(6) +")" ;
 cells(12, 1).Value = stats.StatisticName(7) + " (" + stats.StatisticUnits(7) +")" ;
 cells(13, 1).Value = stats.StatisticName(8) + " (" + stats.StatisticUnits(8) +")" ;
 cells(14, 1).Value = stats.StatisticName(9) + " (" + stats.StatisticUnits(9) +")" ;


columns = frames(0).Readings(0).Columns;
rows = frames(0).Readings(0).Rows;

		cnt=0;
		var character;
		var asc = 65;
		for (c = 0; c < columns; ++c){
			for (r = 0; r < rows; ++r){
				cnt++;
				if (c<26){
					character = String.fromCharCode(65 + c);
				} 
				else {
					character = String.fromCharCode(97 + (c-26));
				}				
				cells(cnt+15,1).Value = character + (r+1)
			}
		}

var oldcopx = -1;
var oldcopy = -1;	
var sensors = new Array();
col=1;

//for (f = 0; f < count; ++f)
for (x = 0; x < TotalFrames; ++x) //a.length 12 for 1hr and 144 for 12hrs
	//if (frames(f).Selected)
	
	if (1)
	{
		f = a[x]-1;
		col++;
		view = application.View;
		recordingView = view.RecordingView.Panel("readings");
		stats = recordingView.Panel(0).Statistics;
		stats.FrameIndex = f;
		stats.Recalculate(2);
		cells(3,col).Value = f+1;
		//cells(4,col).Value = mydateFunction(stats.StatisticTime(0));
		cells(4,col).Value = frames(f).TimeString;

		
		i = 5;
		for(w=0; w<10;w++){
		cells(i+w, col).Value = stats.StatisticValue(0, w, 0).toFixed(2);

			if(w==6){
				newcopx = stats.StatisticValue(0, w, 0).toFixed(2);
			}
			if(w==7){
				newcopy = stats.StatisticValue(0, w, 0).toFixed(2);
			}
		}
			if(oldcopx != -1 && oldcopy !=-1){
			//	cells(15,col).Value = Math.sqrt((Math.pow(newcopx - oldcopx,2) + Math.pow(newcopy - oldcopy,2))).toFixed(2);
			//	cells(16,col).Value = ((Math.atan((newcopy - oldcopy)/(newcopx - oldcopx)))*180 / Math.PI).toFixed(2);					
			}
				oldcopx = newcopx;
				oldcopy = newcopy;

		// copy the values into the spreadsheet

		sheet.EnableCalculation = false;
		reading = frames(f).Readings(0);
		cnt=0;
		//reading = frames(f).regions(q);
		for (c = 0; c < columns; ++c)
		{
			//if(c1-1>=c || c>=c2) continue;
			for (r = 0; r < rows; ++r)
			{
				
				//sensors.push(reading.Value(r, c));
				cnt++;
				cells(cnt+15, col).Value = reading.Value(c, r);

			}
		}
		sheet.EnableCalculation = true;
		selectedFrames++;
		
	}

}//end of else

//if(no_of_regions>0)
//for (s = sheets.Count; s > no_of_regions; --s) sheets(s).Delete;

//excel.ScreenUpdating= true;
excel.StatusBar = "Ready";
