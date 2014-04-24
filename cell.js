function cell(options) {		
	if(!(this instanceof cell)) return new cell(options);
	var config = extend({val: null, style: {}, type: null}, options);
	this.val = function(str) {
		if(str) {
			config.val = str;
			return this;
		}
		return config.val;
	}
	this.style = function(styling) {
		if(styling) {
			config.style = extend(config.style, styling);
			return this;
		}
		return config.style;
	}
	this.type = function(newType) {
		if(newType) {
			config.type = newType;
			return this;
		}
		return config.type;
	}
	return this;
}

module.exports = cell;