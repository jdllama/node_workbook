var row = require("./row");
function sheet(myLabel, _this, registered) {
	if(!(this instanceof sheet)) return new sheet(myLabel, _this, registered);
	var _this = _this;
	var label = myLabel;
	this.addRow = function(arr) {
		if(!registered[label]) throw new Error("Sheet " + label + " no longer exists.");
		if(!arr) throw new Error("Row missing.");
		if(!(arr instanceof row) && Object.prototype.toString.call(arr) !== "[object Array]") throw new Error("Array expected in addRow");
		registered[label].config.rows.push(arr);
		return this;
	}
	this.removeRow = function() {
		if(!registered[label]) throw new Error("Sheet " + label + " no longer exists.");
		registered[label].config.rows.pop();
		return this;
	}
	this.parent = function() {
		return _this;
	}
	this.setActive = function() {
		activeTab = label;
		return this;
	}
	this.setLabel = function(newLabel) {
		if(registered[newLabel]) throw new Error("Sheet already exists with label " + newLabel);
		var temp = registered[label];
		delete registered[label];
		registered[newLabel] = temp;
		label = newLabel;
		return this;
	}
	this.removeSheet = function() {
		this.parent().removeSheet(label);
		delete this;
	}
	this.toRaw = function() {
		if(!registered[label]) throw new Error("Sheet " + label + " no longer exists.");
		var t = registered[label];
		return t;
	}
	//This assumes that the first row is header information, and will use that as the basis.
	this.toObjectArray = function() {
		if(!registered[label]) throw new Error("Sheet " + label + " no longer exists.");
		var out = [];
		if(registered[label].config.rows.length < 2) throw new Error("Sheet " + label + " does not have enough data to return an object.");
		var rows = registered[label].config.rows;
		var head = rows[0];
		var keys = {};
		for(var i = 0;i<head.length;i++) {
			keys[i] = head[i];
		}
		for(var i = 1;i<rows.length;i++) {
			var row = rows[i];
			var obj = {};
			for(var j in keys) {
				var myKey = keys[j];
				obj[myKey] = null;
			}
			for(var j = 0;j<row.length;j++) {
				if(keys[j]) obj[keys[j]] = row[j];
			}
			out.push(obj);
		}
		return out;
	}
	return this;
}

module.exports = sheet;