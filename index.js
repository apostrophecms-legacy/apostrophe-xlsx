var xlsx = require('xlsx');

module.exports = aposXlsx;

function aposXlsx(options, callback) {
  return new aposXlsx.Construct(options, callback);
}

aposXlsx.Construct = function(options, callback) {
  var self = this;
  self._apos = options.apos;


  // Convert Excel file to CSV
	self.xlsxToCsv = function(filePath) {
		var result = [];
		workbook = xlsx.readFile(filePath, {type: 'base64'});
		workbook.SheetNames.forEach( function(sheetName) {
	    var csv = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName]);
	    if(csv.length > 0){
        result.push(csv);
	    }
		});
		return result.join("\n");
	}

	// Include xlsx as a supported format
	self._apos.on('supportedDataIO' , function(fileTypes) {
	  fileTypes.dataImport.push('XLSX');
	});

	// Event for parsing excel file on import
	self._apos.on('xlsxImport' , function(context) {
	  context.csv = self.xlsxToCsv(context.path);
	});


	// Must wait at least until next tick to invoke callback!
  if (callback) {
    process.nextTick(function() { return callback(null); });
  }
}

