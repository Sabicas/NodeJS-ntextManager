var fs = require("fs");
var xl = require('excel4node');
var Converter = require("csvtojson").Converter;
var json2xls = require('json2xls');



var converter = new Converter({});
var wb = new xl.Workbook();

var clientName
var clientDir;
var csvDir;
var ntextDir;
var fileArr = [];
var csvArr = [];
var fileTitle = 0;

let ntext = {};

ntext.createDirVars = function(cName){
	clientName = cName;
	clientDir = "./" + clientName + "/";
	if(fs.existsSync(clientDir)){		
		csvDir = clientDir+"csvFiles/";
		ntextDir = clientDir+"NTEXT/";
		return true;
	}else{			
		return false;
	}
}

ntext.Perform =  function(){
	fs.readdirSync(csvDir).forEach(file => {
		//console.log(file);
		fileArr.push(file);
	});

	for(f = 0;f < fileArr.length; f++){
		var file = csvDir + fileArr[f];
		var csvEncoding = { encoding: 'utf16le' }; 
		var csvString = fs.readFileSync(file, csvEncoding).toString(); 
		
		csvArr.push(csvString);

		ntext.ConvertToJson(csvString).then(function(responseData) {
	       ntext.ConvertToXls(responseData);
	    }).catch(function(error) {
	        console.log("ERROR: " + error)        
	    });		
	}
}




ntext.ConvertToJson =  function(csvStr){
	console.log("STEP: ConvertToJson");
	var testblah = csvStr.length;
	return new Promise(function(resolve,reject) {
		var jsonObj;
		var converter = new Converter({noheader:true});
		converter.fromString(csvStr, function(err,result){ 	 		
	 		if(result.length > 0){
	 			jsonObj = result;	 			
	 			resolve(jsonObj);
	 		}else{
	 			console.log("ConvertToJson ERROR: " + err)
	 			reject(err);
	 		}	 		
	 	});
	})	
}

ntext.ConvertToXls = function(jsonObj){
		console.log("STEP: ConvertToXls");		
		var rows = jsonObj.length;		
		var columns = Object.keys(jsonObj[0]);

		//create new worksheet
		var ws = wb.addWorksheet(fileArr[fileTitle]);
		
		//rows
		for(r=0;r < rows;r++){
			//columns
			for(c=0;c < columns.length;c++){
				var thisCell = 	ws.cell(r+1,c+1);			
				var cellData = jsonObj[r][columns[c]];

				//.xlsx cells have a limit of 32,767 chars.  We need to create a .txt file for every NTEXT that exceeds this limit and add a reference to the cell.
				var primaryKey = jsonObj[r][columns[0]] + " - " + jsonObj[0][columns[c]];
				var dataLen = cellData.length
				if(dataLen){
					if(dataLen > 32767){
						console.log("DATA LIMIT EXCEEDED: " + dataLen);						
						
						fs.writeFileSync(ntext.CreateDir(fileArr[fileTitle],primaryKey),cellData, 'utf-8');
						thisCell.style({fill: {type: 'pattern', patternType: 'solid', fgColor: 'yellow'}});
						thisCell.string("Data limit exeeded for this cell.  See included file: " + ntextDir + fileArr[fileTitle] + "/" + primaryKey);
					}else{
						thisCell.string(cellData);
					}
				}			
			}
		}

	  	if(fileTitle == fileArr.length - 1){
	 		ntext.WriteFile();
	 	}else{
		 	fileTitle++;
		 }
	
}

ntext.CreateDir = function(fileName,primKey){
	console.log("KOMPLETE: " + ntextDir + fileName + "/" + primKey);
	//create directory for specific file
	if(!fs.existsSync(ntextDir + fileName)){
		fs.mkdirSync(ntextDir + fileName);
	}
	return ntextDir + fileName + "/" + primKey;	
}

ntext.WriteFile = function(){
	console.log("STEP: WRITING DOCUMENT")
	wb.write(clientDir + clientName + '.xlsx');
}

module.exports = ntext;
ntext.createDirVars(process.argv[2]);
ntext.Perform();





 