var xlsx = require('xlsx');

module.exports = aposXlsx;

function aposXlsx(options, callback) {
  return new aposXlsx.Construct(options, callback);
}

aposXlsx.Construct = function(options, callback) {
  var self = this;
  self._apos = options.apos;


  // --------------------------------------------------------------------------- //
	// Utilities for generating XLSX files
	// --------------------------------------------------------------------------- //

	self.datenum = function(v, date1904) {
		if(date1904) v+=1462;
		var epoch = Date.parse(v);
		return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
	}
	 
	self.arrayOfArraysToSheet = function(data, opts) {
		var ws = {};
		var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
		for(var R = 0; R != data.length; ++R) {
			for(var C = 0; C != data[R].length; ++C) {
				if(range.s.r > R) range.s.r = R;
				if(range.s.c > C) range.s.c = C;
				if(range.e.r < R) range.e.r = R;
				if(range.e.c < C) range.e.c = C;
				var cell = {v: data[R][C] };
				if(cell.v == null) continue;
				var cell_ref = xlsx.utils.encode_cell({c:C,r:R});
				
				if(typeof cell.v === 'number') cell.t = 'n';
				else if(typeof cell.v === 'boolean') cell.t = 'b';
				else if(cell.v instanceof Date) {
					cell.t = 'n'; cell.z = xslx.SSF._table[14];
					cell.v = self.datenum(cell.v);
				}
				else cell.t = 's';
				
				ws[cell_ref] = cell;
			}
		}
		if(range.s.c < 10000000) ws['!ref'] = xlsx.utils.encode_range(range);
		return ws;
	}

	function Workbook() {
		if(!(this instanceof Workbook)) return new Workbook();
		this.SheetNames = [];
		this.Sheets = {};
	}

	// --------------------------------------------------------------------------- //
	// CONVERTERS
	// --------------------------------------------------------------------------- //

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
	 
	self.arrayOfArraysToXlsx = function(data) {
		var wb = new Workbook(),
				ws = self.arrayOfArraysToSheet(data);

		wb.SheetNames.push('A');
		wb.Sheets['A'] = ws;
		return xlsx.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
	}

	// --------------------------------------------------------------------------- //
	// EVENTS
	// --------------------------------------------------------------------------- //

	// Include xlsx as a supported format.
	// This is accessed by apostrophe-snippets
	// to generate the help text in Import.html
	// and Export.html
	self._apos.on('supportedDataIO' , function(supportedDataIO) {
	  supportedDataIO.formats.push({ label: 'XLSX', value: 'xlsx'});
	});

	// Parse xlsx file to CSV on import
	self._apos.on('xlsxImport' , function(context) {
	  context.csv = self.xlsxToCsv(context.path);
	});

	// Generate xlsx file from snippets data
	self._apos.on('xlsxExport' , function(context) {
	  context.xlsx = self.arrayOfArraysToXlsx(context.data);
	});


	// Asset mixin
  self._apos.mixinModuleAssets(self, 'xlsx', __dirname, options);
	// Push fileSaver js for downloading xlsx binary
	self.pushAsset('script', 'vendor/FileSaver.min', { when: 'user' });


	// Must wait at least until next tick to invoke callback!
  if (callback) {
    process.nextTick(function() { return callback(null); });
  }
}

