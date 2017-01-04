(function($) {
	var Export = function(options) {
		this.grid = null;
		this._columnInfo = null;
	};

	function datenum(v, date1904) {
		if (date1904)
			v += 1462;
		var epoch = Date.parse(v);
		return (epoch - new Date(Date.UTC(2899, 11, 30))) / (24 * 60 * 60 * 1000);
	}

	function fillData(data, columnInfo, opts) {
		var worksheet = {};
		var range = {
			s : {
				c : 0,
				r : 0
			},
			e : {
				c : 0,
				r : 0
			}
		};

		var __colIdx = 0, __col = 0;
		for ( var idx in columnInfo) {
			var _columnDef = columnInfo[idx];
			if (_columnDef.toExport == false) {
				continue;
			}
			var cell = {
				v : _columnDef.name
			};
			__colIdx = __col++;
			if (cell.v == null) {
				continue;
			}
			var cell_ref = XLSX.utils.encode_cell({
				c : __colIdx,
				r : 0
			});
			cell.t = 's';
			worksheet[cell_ref] = cell;
		}

		var _length = data.length;
		for (var _row = 0; _row != _length; ++_row) {
			var _rowData = data[_row];
			var _colIdx = 0, _col = 0;
			for ( var idx in columnInfo) {
				var _columnDef = columnInfo[idx];
				if (_columnDef.toExport == false) {
					continue;
				}
				if (range.e.r < _row + 1) {
					range.e.r = _row + 1;
				}
				if (range.e.c < _col) {
					range.e.c = _col;
				}
				var _valTmp = null;
				if (_columnDef.formatter) {
					_valTmp = _columnDef.formatter(_row, idx, data[_row][_columnDef.field], _columnDef, data[_row]);
				} else {
					_valTmp = data[_row][_columnDef.field];
				}
				var _val;
				if (null == _valTmp || undefined == _valTmp) {
					_val = null;
				} else if (typeof _valTmp == 'string' || typeof _valTmp == 'number') {
					_val = _valTmp;
				} else {
					if (null == _valTmp.exportVal || undefined == _valTmp.exportVal) {
						_val = _valTmp.content;
					} else {
						_val = _valTmp.exportVal;
					}
				}
				var cell = {
					v : _val
				};
				_colIdx = _col++;
				if (cell.v == null) {
					continue;
				}
				var cell_ref = XLSX.utils.encode_cell({
					c : _colIdx,
					r : _row + 1
				});
				if (typeof cell.v == 'number') {
					cell.t = 'n';
				} else if (typeof cell.v === 'boolean') {
					cell.t = 'b';
				} else if (cell.v instanceof Date) {
					cell.t = 'n';
					cell.z = XLSX.SSF._table[4];
					cell.v = datanum(cell.v);
				} else {
					cell.t = 's';
				}
				worksheet[cell_ref] = cell;
			}
		}

		if (range.s.c < 10000000) {
			worksheet['!ref'] = XLSX.utils.encode_range(range);
		}

		return worksheet;
	}

	function s2ab(s) {
		var buf = new ArrayBuffer(s.length);
		var view = new Uint8Array(buf);
		for (var i = 0; i != s.length; ++i) {
			view[i] = s.charCodeAt(i) & 0xFF;
		}
		return buf;
	}

	Export.prototype.init = function(grid) {
		this.grid = grid;
		this._columnInfo = grid.getColumns();
	};

	/**
	 * { fileName : 文件名, }
	 * 
	 */
	Export.prototype.exportExcel = function(options) {
		var _self = this;
		var worksheet = fillData(_self.grid.getData(), _self._columnInfo);
		var workbook = {
			SheetNames : [],
			Sheets : {}
		};
		var sheetName = 'data';
		workbook.SheetNames.push(sheetName);
		workbook.Sheets[sheetName] = worksheet;
		var wbOuputData = XLSX.write(workbook, {
			bookType : 'xlsx',
			bookSST : true,
			type : 'binary'
		});

		var _fileName = 'data';
		if (options) {
			var __fileName = $.trim(options.fileName);
			if (__fileName.length > 0) {
				_fileName = __fileName;
			}
		}
		saveAs(new Blob([ s2ab(wbOuputData) ], {
			type : "application/octet-stream"
		}), _fileName + ".xlsx");
	};

	$.extend(true, window, {
		"Slick" : {
			"Export" : Export
		}
	});

})(jQuery);