function row(options) {
	if(!(this instanceof row)) return new row(options);
	var config = extend({cells: [], style: {}}, options);
	this.addCell = function(newCell) {
		config.cells.push(newCell);
		return this;
	}
	this.getCells = function() {
		return config.cells;
	}
	this.getStyle = function() {
		return config.style;
	}
	return this;
}

module.exports = row;