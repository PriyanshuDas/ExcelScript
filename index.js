

var DLOAD_PREFIX = 'https://spreadsheets.google.com/feeds/download/spreadsheets/Export?key=';

var DLOAD_SUFFIX = '&exportFormat=xlsx';
var FILE_ID = '1jzg4b2U_dX8zL4Y9GnqqzCGIDUpRNJDHIrb87SJvKeg';

var axios = require('axios');
var fs = require('fs');
var Excel = require('exceljs');
var _ = require('underscore-node');
var Promise = require('bluebird');

var input = 'input.xlsx';
var output = 'output.xlsx';
var cache = 'lastrow.txt';
var URL_1 = 'https://i.flock.com/singleFlockShortUrl?sub3=';
var URL_2 = '&baseUrl=';

var start_index = 2;
var end_index = 0;

//https://i.flock.com/singleFlockShortUrl?sub3={sub3}&baseUrl={baseUrl}

process.argv.forEach(function (val, index, array) {
	if(index === 2)
		parseURL(val);
});

// processOutput();


function parseURL(val)
{
	var p1 = val.indexOf('/d/');
	var p2 = val.indexOf('/edit');
	p1 += 3;
	FILE_ID = val.substring(p1, p2);
	// console.log('Input : ', val);
	// console.log('ID : ', FILE_ID);

}

var processInput = function()
{
	axios({
	  method:'get',
	  url:DLOAD_PREFIX+FILE_ID,
	  responseType:'stream'
	}).catch(err => console.log(err))
	  .then(function(response) {
    	//console.log('x');
	  response.data.pipe(fs.createWriteStream(input));
	  setTimeout(processOutput, 1000);
	}).catch(err => console.log(err))

	//processOutput();
}

var processOutput = function()
{
	var workbook = new Excel.Workbook();

	if (!fs.existsSync(cache)) {
	    fs.writeFile(cache, '1', function(err) {
	    	if(err)
	    	{
	    		console.log("Couldn't write cache file!");
	    		throw err;
	    	}
	    });
	    console.log('New cache file created!');
	}
	fs.readFile(cache, 'utf8', function(err, data) {
	  if (err) throw err;
	  //console.log('OK: ' + cache);
	  // console.log(data);
	  start_index = data;
	});

	workbook.xlsx.readFile(input)
	    .then(function() {
			var worksheet = workbook.getWorksheet('Google');
	    	end_index = worksheet.rowCount;
	    	var promises = [];
	    	if(start_index == end_index)
	    	{
	    		console.log('Nothing to Update!');
	    		return;
	    	}
	    	console.log('Updating rows : ', start_index, 'to ', end_index);
	    	for(var i = Number(start_index)+1; i <= end_index; i++)
	    	{
	    		var requestURL = URL_1 + worksheet.getCell('H'+ i.toString()) + URL_2 + worksheet.getCell('B' + i.toString());
	    		var outputData;
	    		(function(cntr)
	    		{
		    		var promise = axios({
					  method:'get',
					  url:requestURL,
					  responseType:'json'
					}).catch(err => console.log(err));
					promise.then(function(response) {
						outputData = response.data.shortUrl;
						worksheet.getCell('C' + cntr.toString()).value = outputData.toString();
					}).catch(err => console.log(err));
					promises.push(promise);
	    		})(i);
	    	}
	    	Promise.all(promises).then(function(response) {

	    		console.log('Success! : File is Up to Date!');
				fs.writeFile(cache, end_index.toString(), function(err) {
			    	if(err)
			    	{
			    		console.log("Couldn't write cache file!");
			    		throw err;
			    	}
			    });
    			return workbook.xlsx.writeFile(output);
	    	});
	    })
	    .catch(err => console.log(err));
};

processInput();