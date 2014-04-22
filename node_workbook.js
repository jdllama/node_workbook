if(typeof JSZip === "undefined" || !JSZip) var JSZip = require("jszip");
if(typeof xml2js === "undefined" || !xml2js) var xml2js = require("xml2js");

function node_workbook() {
	if(!(this instanceof node_workbook)) return new node_workbook();
	
	var registered = {};
	var labels = [];
	var activeTab = "";
	var sheetCount = 0;
	
	function compare(arr1, arr2) {
		var arr1 = arr1, arr2 = arr2;
		if(arr1.length !== arr2.length) return false;
		for(var i = 0;i<arr1.length;i++) {
			if(arr1[i] instanceof Array && arr2[i] instanceof Array) {
				if(!compare(arr1[i], arr2[i])) return false;
			}
			if(arr1[i] !== arr2[i]) return false;
		}
		return true;
	}
	
	//from http://javascriptweblog.wordpress.com/2011/08/08/fixing-the-javascript-typeof-operator/
	var toType = function(obj) {
		return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase();
	}
	
	//from http://krasimirtsonev.com/blog/article/JavaScript-dependency-free-extend-method
	var extend = function() {
		var process = function(destination, source) {
			for (var key in source) {
				if (hasOwnProperty.call(source, key)) {
					destination[key] = source[key];
				}
			}
			return destination;
		};
		
		var result = arguments[0];
		
		for(var i=1; i<arguments.length; i++) {
			result = process(result, arguments[i]);
		}
		return result;
	};
	
	//from http://stackoverflow.com/questions/1349404/generate-a-string-of-5-random-characters-in-javascript
	var rand = function(len, charSet) {
		charSet = charSet || 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
		var randomString = '';
		for (var i = 0; i < len; i++) {
			var randomPoz = Math.floor(Math.random() * charSet.length);
			randomString += charSet.substring(randomPoz,randomPoz+1);
		}
		return randomString;
	}
	
	//Care of http://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
	var colToChar = function(num) {
		var temp, letter = "";
		while(num > 0) {
			temp = (num - 1) % 26;
			letter = String.fromCharCode(temp + 65) + letter;
			num = (num - temp - 1) / 26;
		}
		return letter;
	}
	var charToCol = function(letter){
		var column = 0, length = letter.length;
		for(var i = 0;i < length;i++) {
			column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
		}
		return column;
	}
	
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
	
	this.cell = cell;
	
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
	
	this.row = row;
	
	this.generateXlsx = function(fileName) {
		var keys = Object.keys(registered);
		if(keys.length === 0) {
			throw new Error("No sheets added.");
		}

		var template = "UEsDBAoAAAAAACxgjkRfjmCboQUAAKEFAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbDw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9IlVURi04IiBzdGFuZGFsb25lPSJ5ZXMiPz4NCjxUeXBlcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3BhY2thZ2UvMjAwNi9jb250ZW50LXR5cGVzIj48RGVmYXVsdCBFeHRlbnNpb249InJlbHMiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtcGFja2FnZS5yZWxhdGlvbnNoaXBzK3htbCIvPjxEZWZhdWx0IEV4dGVuc2lvbj0ieG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24veG1sIi8+PE92ZXJyaWRlIFBhcnROYW1lPSIveGwvd29ya2Jvb2sueG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LnNwcmVhZHNoZWV0bWwuc2hlZXQubWFpbit4bWwiLz48T3ZlcnJpZGUgUGFydE5hbWU9Ii94bC93b3Jrc2hlZXRzL3NoZWV0MS54bWwiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQuc3ByZWFkc2hlZXRtbC53b3Jrc2hlZXQreG1sIi8+PE92ZXJyaWRlIFBhcnROYW1lPSIveGwvd29ya3NoZWV0cy9zaGVldDIueG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LnNwcmVhZHNoZWV0bWwud29ya3NoZWV0K3htbCIvPjxPdmVycmlkZSBQYXJ0TmFtZT0iL3hsL3dvcmtzaGVldHMvc2hlZXQzLnhtbCIgQ29udGVudFR5cGU9ImFwcGxpY2F0aW9uL3ZuZC5vcGVueG1sZm9ybWF0cy1vZmZpY2Vkb2N1bWVudC5zcHJlYWRzaGVldG1sLndvcmtzaGVldCt4bWwiLz48T3ZlcnJpZGUgUGFydE5hbWU9Ii94bC90aGVtZS90aGVtZTEueG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LnRoZW1lK3htbCIvPjxPdmVycmlkZSBQYXJ0TmFtZT0iL3hsL3N0eWxlcy54bWwiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQuc3ByZWFkc2hlZXRtbC5zdHlsZXMreG1sIi8+PE92ZXJyaWRlIFBhcnROYW1lPSIvZG9jUHJvcHMvY29yZS54bWwiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtcGFja2FnZS5jb3JlLXByb3BlcnRpZXMreG1sIi8+PE92ZXJyaWRlIFBhcnROYW1lPSIveGwvc2hhcmVkU3RyaW5ncy54bWwiIENvbnRlbnRUeXBlPSJhcHBsaWNhdGlvbi92bmQub3BlbnhtbGZvcm1hdHMtb2ZmaWNlZG9jdW1lbnQuc3ByZWFkc2hlZXRtbC5zaGFyZWRTdHJpbmdzK3htbCIgLz48T3ZlcnJpZGUgUGFydE5hbWU9Ii9kb2NQcm9wcy9hcHAueG1sIiBDb250ZW50VHlwZT0iYXBwbGljYXRpb24vdm5kLm9wZW54bWxmb3JtYXRzLW9mZmljZWRvY3VtZW50LmV4dGVuZGVkLXByb3BlcnRpZXMreG1sIi8+PC9UeXBlcz5QSwMECgAAAAAAmkSPRAAAAAAAAAAAAAAAAAYAAABfcmVscy9QSwMECgAAAAAAV4KIRJImhaa7AQAAuwEAAAsAAABfcmVscy8ucmVsczw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9IlVURi04IiBzdGFuZGFsb25lPSJ5ZXMiPz4NCjxSZWxhdGlvbnNoaXBzIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvcGFja2FnZS8yMDA2L3JlbGF0aW9uc2hpcHMiPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDMiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvZXh0ZW5kZWQtcHJvcGVydGllcyIgVGFyZ2V0PSJkb2NQcm9wcy9hcHAueG1sIi8+PFJlbGF0aW9uc2hpcCBJZD0icklkMSIgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9vZmZpY2VEb2N1bWVudCIgVGFyZ2V0PSJ4bC93b3JrYm9vay54bWwiLz48L1JlbGF0aW9uc2hpcHM+UEsDBAoAAAAAAJpEj0QAAAAAAAAAAAAAAAADAAAAeGwvUEsDBAoAAAAAAJpEj0QAAAAAAAAAAAAAAAAJAAAAeGwvX3JlbHMvUEsDBAoAAAAAANhrikRh1pT7ngEAAJ4BAAAaAAAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHM8P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJVVEYtOCIgc3RhbmRhbG9uZT0ieWVzIj8+PFJlbGF0aW9uc2hpcHMgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9wYWNrYWdlLzIwMDYvcmVsYXRpb25zaGlwcyI+PFJlbGF0aW9uc2hpcCBJZD0icklkNSIgVHlwZT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvcmVsYXRpb25zaGlwcy9zdHlsZXMiIFRhcmdldD0ic3R5bGVzLnhtbCIvPjxSZWxhdGlvbnNoaXAgSWQ9InJJZDQiIFR5cGU9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMvdGhlbWUiIFRhcmdldD0idGhlbWUvdGhlbWUxLnhtbCIvPjwvUmVsYXRpb25zaGlwcz5QSwMECgAAAAAAmkSPRAAAAAAAAAAAAAAAAAkAAAB4bC90aGVtZS9QSwMECgAAAAAAAAAhAPtipW2nGwAApxsAABMAAAB4bC90aGVtZS90aGVtZTEueG1sPD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pg0KPGE6dGhlbWUgeG1sbnM6YT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL2RyYXdpbmdtbC8yMDA2L21haW4iIG5hbWU9Ik9mZmljZSBUaGVtZSI+PGE6dGhlbWVFbGVtZW50cz48YTpjbHJTY2hlbWUgbmFtZT0iT2ZmaWNlIj48YTpkazE+PGE6c3lzQ2xyIHZhbD0id2luZG93VGV4dCIgbGFzdENscj0iMDAwMDAwIi8+PC9hOmRrMT48YTpsdDE+PGE6c3lzQ2xyIHZhbD0id2luZG93IiBsYXN0Q2xyPSJGRkZGRkYiLz48L2E6bHQxPjxhOmRrMj48YTpzcmdiQ2xyIHZhbD0iMUY0OTdEIi8+PC9hOmRrMj48YTpsdDI+PGE6c3JnYkNsciB2YWw9IkVFRUNFMSIvPjwvYTpsdDI+PGE6YWNjZW50MT48YTpzcmdiQ2xyIHZhbD0iNEY4MUJEIi8+PC9hOmFjY2VudDE+PGE6YWNjZW50Mj48YTpzcmdiQ2xyIHZhbD0iQzA1MDREIi8+PC9hOmFjY2VudDI+PGE6YWNjZW50Mz48YTpzcmdiQ2xyIHZhbD0iOUJCQjU5Ii8+PC9hOmFjY2VudDM+PGE6YWNjZW50ND48YTpzcmdiQ2xyIHZhbD0iODA2NEEyIi8+PC9hOmFjY2VudDQ+PGE6YWNjZW50NT48YTpzcmdiQ2xyIHZhbD0iNEJBQ0M2Ii8+PC9hOmFjY2VudDU+PGE6YWNjZW50Nj48YTpzcmdiQ2xyIHZhbD0iRjc5NjQ2Ii8+PC9hOmFjY2VudDY+PGE6aGxpbms+PGE6c3JnYkNsciB2YWw9IjAwMDBGRiIvPjwvYTpobGluaz48YTpmb2xIbGluaz48YTpzcmdiQ2xyIHZhbD0iODAwMDgwIi8+PC9hOmZvbEhsaW5rPjwvYTpjbHJTY2hlbWU+PGE6Zm9udFNjaGVtZSBuYW1lPSJPZmZpY2UiPjxhOm1ham9yRm9udD48YTpsYXRpbiB0eXBlZmFjZT0iQ2FtYnJpYSIvPjxhOmVhIHR5cGVmYWNlPSIiLz48YTpjcyB0eXBlZmFjZT0iIi8+PGE6Zm9udCBzY3JpcHQ9IkpwYW4iIHR5cGVmYWNlPSLvvK3vvLMg77yw44K044K344OD44KvIi8+PGE6Zm9udCBzY3JpcHQ9IkhhbmciIHR5cGVmYWNlPSLrp5HsnYAg6rOg65SVIi8+PGE6Zm9udCBzY3JpcHQ9IkhhbnMiIHR5cGVmYWNlPSLlrovkvZMiLz48YTpmb250IHNjcmlwdD0iSGFudCIgdHlwZWZhY2U9IuaWsOe0sOaYjumrlCIvPjxhOmZvbnQgc2NyaXB0PSJBcmFiIiB0eXBlZmFjZT0iVGltZXMgTmV3IFJvbWFuIi8+PGE6Zm9udCBzY3JpcHQ9IkhlYnIiIHR5cGVmYWNlPSJUaW1lcyBOZXcgUm9tYW4iLz48YTpmb250IHNjcmlwdD0iVGhhaSIgdHlwZWZhY2U9IlRhaG9tYSIvPjxhOmZvbnQgc2NyaXB0PSJFdGhpIiB0eXBlZmFjZT0iTnlhbGEiLz48YTpmb250IHNjcmlwdD0iQmVuZyIgdHlwZWZhY2U9IlZyaW5kYSIvPjxhOmZvbnQgc2NyaXB0PSJHdWpyIiB0eXBlZmFjZT0iU2hydXRpIi8+PGE6Zm9udCBzY3JpcHQ9IktobXIiIHR5cGVmYWNlPSJNb29sQm9yYW4iLz48YTpmb250IHNjcmlwdD0iS25kYSIgdHlwZWZhY2U9IlR1bmdhIi8+PGE6Zm9udCBzY3JpcHQ9Ikd1cnUiIHR5cGVmYWNlPSJSYWF2aSIvPjxhOmZvbnQgc2NyaXB0PSJDYW5zIiB0eXBlZmFjZT0iRXVwaGVtaWEiLz48YTpmb250IHNjcmlwdD0iQ2hlciIgdHlwZWZhY2U9IlBsYW50YWdlbmV0IENoZXJva2VlIi8+PGE6Zm9udCBzY3JpcHQ9IllpaWkiIHR5cGVmYWNlPSJNaWNyb3NvZnQgWWkgQmFpdGkiLz48YTpmb250IHNjcmlwdD0iVGlidCIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBIaW1hbGF5YSIvPjxhOmZvbnQgc2NyaXB0PSJUaGFhIiB0eXBlZmFjZT0iTVYgQm9saSIvPjxhOmZvbnQgc2NyaXB0PSJEZXZhIiB0eXBlZmFjZT0iTWFuZ2FsIi8+PGE6Zm9udCBzY3JpcHQ9IlRlbHUiIHR5cGVmYWNlPSJHYXV0YW1pIi8+PGE6Zm9udCBzY3JpcHQ9IlRhbWwiIHR5cGVmYWNlPSJMYXRoYSIvPjxhOmZvbnQgc2NyaXB0PSJTeXJjIiB0eXBlZmFjZT0iRXN0cmFuZ2VsbyBFZGVzc2EiLz48YTpmb250IHNjcmlwdD0iT3J5YSIgdHlwZWZhY2U9IkthbGluZ2EiLz48YTpmb250IHNjcmlwdD0iTWx5bSIgdHlwZWZhY2U9IkthcnRpa2EiLz48YTpmb250IHNjcmlwdD0iTGFvbyIgdHlwZWZhY2U9IkRva0NoYW1wYSIvPjxhOmZvbnQgc2NyaXB0PSJTaW5oIiB0eXBlZmFjZT0iSXNrb29sYSBQb3RhIi8+PGE6Zm9udCBzY3JpcHQ9Ik1vbmciIHR5cGVmYWNlPSJNb25nb2xpYW4gQmFpdGkiLz48YTpmb250IHNjcmlwdD0iVmlldCIgdHlwZWZhY2U9IlRpbWVzIE5ldyBSb21hbiIvPjxhOmZvbnQgc2NyaXB0PSJVaWdoIiB0eXBlZmFjZT0iTWljcm9zb2Z0IFVpZ2h1ciIvPjxhOmZvbnQgc2NyaXB0PSJHZW9yIiB0eXBlZmFjZT0iU3lsZmFlbiIvPjwvYTptYWpvckZvbnQ+PGE6bWlub3JGb250PjxhOmxhdGluIHR5cGVmYWNlPSJDYWxpYnJpIi8+PGE6ZWEgdHlwZWZhY2U9IiIvPjxhOmNzIHR5cGVmYWNlPSIiLz48YTpmb250IHNjcmlwdD0iSnBhbiIgdHlwZWZhY2U9Iu+8re+8syDvvLDjgrTjgrfjg4Pjgq8iLz48YTpmb250IHNjcmlwdD0iSGFuZyIgdHlwZWZhY2U9IuunkeydgCDqs6DrlJUiLz48YTpmb250IHNjcmlwdD0iSGFucyIgdHlwZWZhY2U9IuWui+S9kyIvPjxhOmZvbnQgc2NyaXB0PSJIYW50IiB0eXBlZmFjZT0i5paw57Sw5piO6auUIi8+PGE6Zm9udCBzY3JpcHQ9IkFyYWIiIHR5cGVmYWNlPSJBcmlhbCIvPjxhOmZvbnQgc2NyaXB0PSJIZWJyIiB0eXBlZmFjZT0iQXJpYWwiLz48YTpmb250IHNjcmlwdD0iVGhhaSIgdHlwZWZhY2U9IlRhaG9tYSIvPjxhOmZvbnQgc2NyaXB0PSJFdGhpIiB0eXBlZmFjZT0iTnlhbGEiLz48YTpmb250IHNjcmlwdD0iQmVuZyIgdHlwZWZhY2U9IlZyaW5kYSIvPjxhOmZvbnQgc2NyaXB0PSJHdWpyIiB0eXBlZmFjZT0iU2hydXRpIi8+PGE6Zm9udCBzY3JpcHQ9IktobXIiIHR5cGVmYWNlPSJEYXVuUGVuaCIvPjxhOmZvbnQgc2NyaXB0PSJLbmRhIiB0eXBlZmFjZT0iVHVuZ2EiLz48YTpmb250IHNjcmlwdD0iR3VydSIgdHlwZWZhY2U9IlJhYXZpIi8+PGE6Zm9udCBzY3JpcHQ9IkNhbnMiIHR5cGVmYWNlPSJFdXBoZW1pYSIvPjxhOmZvbnQgc2NyaXB0PSJDaGVyIiB0eXBlZmFjZT0iUGxhbnRhZ2VuZXQgQ2hlcm9rZWUiLz48YTpmb250IHNjcmlwdD0iWWlpaSIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBZaSBCYWl0aSIvPjxhOmZvbnQgc2NyaXB0PSJUaWJ0IiB0eXBlZmFjZT0iTWljcm9zb2Z0IEhpbWFsYXlhIi8+PGE6Zm9udCBzY3JpcHQ9IlRoYWEiIHR5cGVmYWNlPSJNViBCb2xpIi8+PGE6Zm9udCBzY3JpcHQ9IkRldmEiIHR5cGVmYWNlPSJNYW5nYWwiLz48YTpmb250IHNjcmlwdD0iVGVsdSIgdHlwZWZhY2U9IkdhdXRhbWkiLz48YTpmb250IHNjcmlwdD0iVGFtbCIgdHlwZWZhY2U9IkxhdGhhIi8+PGE6Zm9udCBzY3JpcHQ9IlN5cmMiIHR5cGVmYWNlPSJFc3RyYW5nZWxvIEVkZXNzYSIvPjxhOmZvbnQgc2NyaXB0PSJPcnlhIiB0eXBlZmFjZT0iS2FsaW5nYSIvPjxhOmZvbnQgc2NyaXB0PSJNbHltIiB0eXBlZmFjZT0iS2FydGlrYSIvPjxhOmZvbnQgc2NyaXB0PSJMYW9vIiB0eXBlZmFjZT0iRG9rQ2hhbXBhIi8+PGE6Zm9udCBzY3JpcHQ9IlNpbmgiIHR5cGVmYWNlPSJJc2tvb2xhIFBvdGEiLz48YTpmb250IHNjcmlwdD0iTW9uZyIgdHlwZWZhY2U9Ik1vbmdvbGlhbiBCYWl0aSIvPjxhOmZvbnQgc2NyaXB0PSJWaWV0IiB0eXBlZmFjZT0iQXJpYWwiLz48YTpmb250IHNjcmlwdD0iVWlnaCIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBVaWdodXIiLz48YTpmb250IHNjcmlwdD0iR2VvciIgdHlwZWZhY2U9IlN5bGZhZW4iLz48L2E6bWlub3JGb250PjwvYTpmb250U2NoZW1lPjxhOmZtdFNjaGVtZSBuYW1lPSJPZmZpY2UiPjxhOmZpbGxTdHlsZUxzdD48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48L2E6c29saWRGaWxsPjxhOmdyYWRGaWxsIHJvdFdpdGhTaGFwZT0iMSI+PGE6Z3NMc3Q+PGE6Z3MgcG9zPSIwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6dGludCB2YWw9IjUwMDAwIi8+PGE6c2F0TW9kIHZhbD0iMzAwMDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6Z3MgcG9zPSIzNTAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSIzNzAwMCIvPjxhOnNhdE1vZCB2YWw9IjMwMDAwMCIvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjxhOmdzIHBvcz0iMTAwMDAwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6dGludCB2YWw9IjE1MDAwIi8+PGE6c2F0TW9kIHZhbD0iMzUwMDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PC9hOmdzTHN0PjxhOmxpbiBhbmc9IjE2MjAwMDAwIiBzY2FsZWQ9IjEiLz48L2E6Z3JhZEZpbGw+PGE6Z3JhZEZpbGwgcm90V2l0aFNoYXBlPSIxIj48YTpnc0xzdD48YTpncyBwb3M9IjAiPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIj48YTpzaGFkZSB2YWw9IjUxMDAwIi8+PGE6c2F0TW9kIHZhbD0iMTMwMDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6Z3MgcG9zPSI4MDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNoYWRlIHZhbD0iOTMwMDAiLz48YTpzYXRNb2QgdmFsPSIxMzAwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48YTpncyBwb3M9IjEwMDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNoYWRlIHZhbD0iOTQwMDAiLz48YTpzYXRNb2QgdmFsPSIxMzUwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48L2E6Z3NMc3Q+PGE6bGluIGFuZz0iMTYyMDAwMDAiIHNjYWxlZD0iMCIvPjwvYTpncmFkRmlsbD48L2E6ZmlsbFN0eWxlTHN0PjxhOmxuU3R5bGVMc3Q+PGE6bG4gdz0iOTUyNSIgY2FwPSJmbGF0IiBjbXBkPSJzbmciIGFsZ249ImN0ciI+PGE6c29saWRGaWxsPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIj48YTpzaGFkZSB2YWw9Ijk1MDAwIi8+PGE6c2F0TW9kIHZhbD0iMTA1MDAwIi8+PC9hOnNjaGVtZUNscj48L2E6c29saWRGaWxsPjxhOnByc3REYXNoIHZhbD0ic29saWQiLz48L2E6bG4+PGE6bG4gdz0iMjU0MDAiIGNhcD0iZmxhdCIgY21wZD0ic25nIiBhbGduPSJjdHIiPjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIgdmFsPSJwaENsciIvPjwvYTpzb2xpZEZpbGw+PGE6cHJzdERhc2ggdmFsPSJzb2xpZCIvPjwvYTpsbj48YTpsbiB3PSIzODEwMCIgY2FwPSJmbGF0IiBjbXBkPSJzbmciIGFsZ249ImN0ciI+PGE6c29saWRGaWxsPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIi8+PC9hOnNvbGlkRmlsbD48YTpwcnN0RGFzaCB2YWw9InNvbGlkIi8+PC9hOmxuPjwvYTpsblN0eWxlTHN0PjxhOmVmZmVjdFN0eWxlTHN0PjxhOmVmZmVjdFN0eWxlPjxhOmVmZmVjdExzdD48YTpvdXRlclNoZHcgYmx1clJhZD0iNDAwMDAiIGRpc3Q9IjIwMDAwIiBkaXI9IjU0MDAwMDAiIHJvdFdpdGhTaGFwZT0iMCI+PGE6c3JnYkNsciB2YWw9IjAwMDAwMCI+PGE6YWxwaGEgdmFsPSIzODAwMCIvPjwvYTpzcmdiQ2xyPjwvYTpvdXRlclNoZHc+PC9hOmVmZmVjdExzdD48L2E6ZWZmZWN0U3R5bGU+PGE6ZWZmZWN0U3R5bGU+PGE6ZWZmZWN0THN0PjxhOm91dGVyU2hkdyBibHVyUmFkPSI0MDAwMCIgZGlzdD0iMjMwMDAiIGRpcj0iNTQwMDAwMCIgcm90V2l0aFNoYXBlPSIwIj48YTpzcmdiQ2xyIHZhbD0iMDAwMDAwIj48YTphbHBoYSB2YWw9IjM1MDAwIi8+PC9hOnNyZ2JDbHI+PC9hOm91dGVyU2hkdz48L2E6ZWZmZWN0THN0PjwvYTplZmZlY3RTdHlsZT48YTplZmZlY3RTdHlsZT48YTplZmZlY3RMc3Q+PGE6b3V0ZXJTaGR3IGJsdXJSYWQ9IjQwMDAwIiBkaXN0PSIyMzAwMCIgZGlyPSI1NDAwMDAwIiByb3RXaXRoU2hhcGU9IjAiPjxhOnNyZ2JDbHIgdmFsPSIwMDAwMDAiPjxhOmFscGhhIHZhbD0iMzUwMDAiLz48L2E6c3JnYkNscj48L2E6b3V0ZXJTaGR3PjwvYTplZmZlY3RMc3Q+PGE6c2NlbmUzZD48YTpjYW1lcmEgcHJzdD0ib3J0aG9ncmFwaGljRnJvbnQiPjxhOnJvdCBsYXQ9IjAiIGxvbj0iMCIgcmV2PSIwIi8+PC9hOmNhbWVyYT48YTpsaWdodFJpZyByaWc9InRocmVlUHQiIGRpcj0idCI+PGE6cm90IGxhdD0iMCIgbG9uPSIwIiByZXY9IjEyMDAwMDAiLz48L2E6bGlnaHRSaWc+PC9hOnNjZW5lM2Q+PGE6c3AzZD48YTpiZXZlbFQgdz0iNjM1MDAiIGg9IjI1NDAwIi8+PC9hOnNwM2Q+PC9hOmVmZmVjdFN0eWxlPjwvYTplZmZlY3RTdHlsZUxzdD48YTpiZ0ZpbGxTdHlsZUxzdD48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48L2E6c29saWRGaWxsPjxhOmdyYWRGaWxsIHJvdFdpdGhTaGFwZT0iMSI+PGE6Z3NMc3Q+PGE6Z3MgcG9zPSIwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6dGludCB2YWw9IjQwMDAwIi8+PGE6c2F0TW9kIHZhbD0iMzUwMDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6Z3MgcG9zPSI0MDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSI0NTAwMCIvPjxhOnNoYWRlIHZhbD0iOTkwMDAiLz48YTpzYXRNb2QgdmFsPSIzNTAwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48YTpncyBwb3M9IjEwMDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNoYWRlIHZhbD0iMjAwMDAiLz48YTpzYXRNb2QgdmFsPSIyNTUwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48L2E6Z3NMc3Q+PGE6cGF0aCBwYXRoPSJjaXJjbGUiPjxhOmZpbGxUb1JlY3QgbD0iNTAwMDAiIHQ9Ii04MDAwMCIgcj0iNTAwMDAiIGI9IjE4MDAwMCIvPjwvYTpwYXRoPjwvYTpncmFkRmlsbD48YTpncmFkRmlsbCByb3RXaXRoU2hhcGU9IjEiPjxhOmdzTHN0PjxhOmdzIHBvcz0iMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSI4MDAwMCIvPjxhOnNhdE1vZCB2YWw9IjMwMDAwMCIvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjxhOmdzIHBvcz0iMTAwMDAwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6c2hhZGUgdmFsPSIzMDAwMCIvPjxhOnNhdE1vZCB2YWw9IjIwMDAwMCIvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjwvYTpnc0xzdD48YTpwYXRoIHBhdGg9ImNpcmNsZSI+PGE6ZmlsbFRvUmVjdCBsPSI1MDAwMCIgdD0iNTAwMDAiIHI9IjUwMDAwIiBiPSI1MDAwMCIvPjwvYTpwYXRoPjwvYTpncmFkRmlsbD48L2E6YmdGaWxsU3R5bGVMc3Q+PC9hOmZtdFNjaGVtZT48L2E6dGhlbWVFbGVtZW50cz48YTpvYmplY3REZWZhdWx0cy8+PGE6ZXh0cmFDbHJTY2hlbWVMc3QvPjwvYTp0aGVtZT5QSwMECgAAAAAAAAAhABeMNyNBBQAAQQUAAA0AAAB4bC9zdHlsZXMueG1sPD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pg0KPHN0eWxlU2hlZXQgeG1sbnM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9zcHJlYWRzaGVldG1sLzIwMDYvbWFpbiIgeG1sbnM6bWM9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9tYXJrdXAtY29tcGF0aWJpbGl0eS8yMDA2IiBtYzpJZ25vcmFibGU9IngxNGFjIiB4bWxuczp4MTRhYz0iaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS9vZmZpY2Uvc3ByZWFkc2hlZXRtbC8yMDA5LzkvYWMiPjxmb250cyBjb3VudD0iMSIgeDE0YWM6a25vd25Gb250cz0iMSI+PGZvbnQ+PHN6IHZhbD0iMTEiLz48Y29sb3IgdGhlbWU9IjEiLz48bmFtZSB2YWw9IkNhbGlicmkiLz48ZmFtaWx5IHZhbD0iMiIvPjxzY2hlbWUgdmFsPSJtaW5vciIvPjwvZm9udD48L2ZvbnRzPjxmaWxscyBjb3VudD0iMiI+PGZpbGw+PHBhdHRlcm5GaWxsIHBhdHRlcm5UeXBlPSJub25lIi8+PC9maWxsPjxmaWxsPjxwYXR0ZXJuRmlsbCBwYXR0ZXJuVHlwZT0iZ3JheTEyNSIvPjwvZmlsbD48L2ZpbGxzPjxib3JkZXJzIGNvdW50PSIxIj48Ym9yZGVyPjxsZWZ0Lz48cmlnaHQvPjx0b3AvPjxib3R0b20vPjxkaWFnb25hbC8+PC9ib3JkZXI+PC9ib3JkZXJzPjxjZWxsU3R5bGVYZnMgY291bnQ9IjEiPjx4ZiBudW1GbXRJZD0iMCIgZm9udElkPSIwIiBmaWxsSWQ9IjAiIGJvcmRlcklkPSIwIi8+PC9jZWxsU3R5bGVYZnM+PGNlbGxYZnMgY291bnQ9IjIiPjx4ZiBudW1GbXRJZD0iMCIgZm9udElkPSIwIiBmaWxsSWQ9IjAiIGJvcmRlcklkPSIwIiB4ZklkPSIwIi8+PHhmIG51bUZtdElkPSIxNCIgZm9udElkPSIwIiBmaWxsSWQ9IjAiIGJvcmRlcklkPSIwIiB4ZklkPSIwIiBhcHBseU51bWJlckZvcm1hdD0iMSIvPjwvY2VsbFhmcz48Y2VsbFN0eWxlcyBjb3VudD0iMSI+PGNlbGxTdHlsZSBuYW1lPSJOb3JtYWwiIHhmSWQ9IjAiIGJ1aWx0aW5JZD0iMCIvPjwvY2VsbFN0eWxlcz48ZHhmcyBjb3VudD0iMCIvPjx0YWJsZVN0eWxlcyBjb3VudD0iMCIgZGVmYXVsdFRhYmxlU3R5bGU9IlRhYmxlU3R5bGVNZWRpdW0yIiBkZWZhdWx0UGl2b3RTdHlsZT0iUGl2b3RTdHlsZUxpZ2h0MTYiLz48ZXh0THN0PjxleHQgdXJpPSJ7RUI3OURFRjItODBCOC00M2U1LTk1QkQtNTRDQkRERjkwMjBDfSIgeG1sbnM6eDE0PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9zcHJlYWRzaGVldG1sLzIwMDkvOS9tYWluIj48eDE0OnNsaWNlclN0eWxlcyBkZWZhdWx0U2xpY2VyU3R5bGU9IlNsaWNlclN0eWxlTGlnaHQxIi8+PC9leHQ+PC9leHRMc3Q+PC9zdHlsZVNoZWV0PlBLAwQKAAAAAACaRI9EAAAAAAAAAAAAAAAACQAAAGRvY1Byb3BzL1BLAwQKAAAAAABqgYtEwyoDfkcDAABHAwAAEAAAAGRvY1Byb3BzL2FwcC54bWw8P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJVVEYtOCIgc3RhbmRhbG9uZT0ieWVzIj8+DQo8UHJvcGVydGllcyB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL29mZmljZURvY3VtZW50LzIwMDYvZXh0ZW5kZWQtcHJvcGVydGllcyIgeG1sbnM6dnQ9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L2RvY1Byb3BzVlR5cGVzIj48QXBwbGljYXRpb24+TWljcm9zb2Z0IEV4Y2VsPC9BcHBsaWNhdGlvbj48RG9jU2VjdXJpdHk+MDwvRG9jU2VjdXJpdHk+PFNjYWxlQ3JvcD5mYWxzZTwvU2NhbGVDcm9wPjxIZWFkaW5nUGFpcnM+PHZ0OnZlY3RvciBzaXplPSIyIiBiYXNlVHlwZT0idmFyaWFudCI+PHZ0OnZhcmlhbnQ+PHZ0Omxwc3RyPldvcmtzaGVldHM8L3Z0Omxwc3RyPjwvdnQ6dmFyaWFudD48dnQ6dmFyaWFudD48dnQ6aTQ+MzwvdnQ6aTQ+PC92dDp2YXJpYW50PjwvdnQ6dmVjdG9yPjwvSGVhZGluZ1BhaXJzPjxUaXRsZXNPZlBhcnRzPjx2dDp2ZWN0b3Igc2l6ZT0iMyIgYmFzZVR5cGU9Imxwc3RyIj48dnQ6bHBzdHI+U2hlZXQxPC92dDpscHN0cj48dnQ6bHBzdHI+U2hlZXQyPC92dDpscHN0cj48dnQ6bHBzdHI+U2hlZXQzPC92dDpscHN0cj48L3Z0OnZlY3Rvcj48L1RpdGxlc09mUGFydHM+PENvbXBhbnk+PC9Db21wYW55PjxMaW5rc1VwVG9EYXRlPmZhbHNlPC9MaW5rc1VwVG9EYXRlPjxTaGFyZWREb2M+ZmFsc2U8L1NoYXJlZERvYz48SHlwZXJsaW5rc0NoYW5nZWQ+ZmFsc2U8L0h5cGVybGlua3NDaGFuZ2VkPjxBcHBWZXJzaW9uPjE0LjAzMDA8L0FwcFZlcnNpb24+PC9Qcm9wZXJ0aWVzPlBLAwQKAAAAAACZbYpEOEvt4+ABAADgAQAADwAAAHhsL3dvcmtib29rLnhtbDw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9IlVURi04IiBzdGFuZGFsb25lPSJ5ZXMiPz4NCjx3b3JrYm9vayB4bWxucz0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL3NwcmVhZHNoZWV0bWwvMjAwNi9tYWluIiB4bWxuczpyPSJodHRwOi8vc2NoZW1hcy5vcGVueG1sZm9ybWF0cy5vcmcvb2ZmaWNlRG9jdW1lbnQvMjAwNi9yZWxhdGlvbnNoaXBzIj48ZmlsZVZlcnNpb24gYXBwTmFtZT0ieGwiIGxhc3RFZGl0ZWQ9IjUiIGxvd2VzdEVkaXRlZD0iNSIgcnVwQnVpbGQ9IjkzMDMiLz48d29ya2Jvb2tQciBkZWZhdWx0VGhlbWVWZXJzaW9uPSIxMjQyMjYiLz48Ym9va1ZpZXdzPjx3b3JrYm9va1ZpZXcgeFdpbmRvdz0iNDgwIiB5V2luZG93PSIzMCIgd2luZG93V2lkdGg9IjI3Nzk1IiB3aW5kb3dIZWlnaHQ9IjEzMzUwIi8+PC9ib29rVmlld3M+PHNoZWV0cyAvPjxjYWxjUHIgY2FsY0lkPSIxNDU2MjEiLz48L3dvcmtib29rPlBLAQIUAAoAAAAAACxgjkRfjmCboQUAAKEFAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAhQACgAAAAAAmkSPRAAAAAAAAAAAAAAAAAYAAAAAAAAAAAAQAAAA0gUAAF9yZWxzL1BLAQIUAAoAAAAAAFeCiESSJoWmuwEAALsBAAALAAAAAAAAAAAAAAAAAPYFAABfcmVscy8ucmVsc1BLAQIUAAoAAAAAAJpEj0QAAAAAAAAAAAAAAAADAAAAAAAAAAAAEAAAANoHAAB4bC9QSwECFAAKAAAAAACaRI9EAAAAAAAAAAAAAAAACQAAAAAAAAAAABAAAAD7BwAAeGwvX3JlbHMvUEsBAhQACgAAAAAA2GuKRGHWlPueAQAAngEAABoAAAAAAAAAAAAAAAAAIggAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzUEsBAhQACgAAAAAAmkSPRAAAAAAAAAAAAAAAAAkAAAAAAAAAAAAQAAAA+AkAAHhsL3RoZW1lL1BLAQIUAAoAAAAAAAAAIQD7YqVtpxsAAKcbAAATAAAAAAAAAAAAAAAAAB8KAAB4bC90aGVtZS90aGVtZTEueG1sUEsBAhQACgAAAAAAAAAhABeMNyNBBQAAQQUAAA0AAAAAAAAAAAAAAAAA9yUAAHhsL3N0eWxlcy54bWxQSwECFAAKAAAAAACaRI9EAAAAAAAAAAAAAAAACQAAAAAAAAAAABAAAABjKwAAZG9jUHJvcHMvUEsBAhQACgAAAAAAaoGLRMMqA35HAwAARwMAABAAAAAAAAAAAAAAAAAAiisAAGRvY1Byb3BzL2FwcC54bWxQSwECFAAKAAAAAACZbYpEOEvt4+ABAADgAQAADwAAAAAAAAAAAAAAAAD/LgAAeGwvd29ya2Jvb2sueG1sUEsFBgAAAAAMAAwAwwIAAAwxAAAAAA==";
		
		var myXlsx = JSZip(template, {base64: true, checkCRC32: true});

		var sharedStrings = [];
		var count = 0;
		for(var sheet in registered) {
			count++;
			var rows = registered[sheet].config.rows;
			if(!rows) throw new Error("Rows not defined for sheet " + sheet);

			var workbook_xml = myXlsx.file("xl/workbook.xml").asText();
			xml2js.parseString(workbook_xml, function(err, result) {
				var relations = result["workbook"]["sheets"];
				
				var temp = relations[0].sheet || [];
				var obj = {
					"$": {
						name: sheet,
						sheetId: count + "",
						"r:id": registered[sheet].id + ""
					}
				};
				temp.push(obj);
				relations[0] = {sheet: temp};
				var builder = new xml2js.Builder();
				workbook_xml = builder.buildObject(result);
			});
			myXlsx.file("xl/workbook.xml", workbook_xml);
			
			var workbook_xml_rels = myXlsx.file("xl/_rels/workbook.xml.rels").asText();
			xml2js.parseString(workbook_xml_rels, function(err, result) {
				var relations = result["Relationships"]["Relationship"];
				var obj = {
					"$": {
						Id: registered[sheet].id + "",
						Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
						Target: "worksheets/sheet" + count + ".xml"
					}
				}
				relations.push(obj);
				var builder = new xml2js.Builder();
				workbook_xml_rels = builder.buildObject(result);
			});
			myXlsx.file("xl/_rels/workbook.xml.rels", workbook_xml_rels);
			
			if(sheet === activeTab) {
				sheetText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
				var workbook_xml = myXlsx.file("xl/workbook.xml").asText();
				xml2js.parseString(workbook_xml, function(err, result) {
					result["workbook"]["bookViews"][0].workbookView[0].$.activeTab = (count-1) + "";
					var builder = new xml2js.Builder();
					workbook_xml = builder.buildObject(result);
				});
				myXlsx.file("xl/workbook.xml", workbook_xml);
			}
			else {
				sheetText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
			}
			
			var styles = [];
			
			xml2js.parseString(sheetText, function(err, result) {
				var sheetData = result["worksheet"]["sheetData"];
				rows.forEach(function(el, index) {
					var currRow = index + 1;
					var style = null;
					var whichStyle = 0;
					if(toType(el) !== "array") {
						style = el.getStyle();
						el = el.getCells();
						//Still need to finish this
						if(Object.keys(style).length) {
							styles.forEach(function(styleEl, styleIndex) {
							});
							if(whichStyle === 0) {
							}
						}
					}
					if(el.length > 0) {
						var obj = {
							"$": {
								"r": currRow + "",
								"x14ac:dyDescent": "0.25"
							},
							c:[]
						}
						if(whichStyle) obj["$"]["s"] = whichStyle;
						el.forEach(function(el2, index2) {
							var myCell = null;
							if(typeof el2 === "number") {
								myCell = {
									"$": {
										"r": colToChar(index2 + 1) + currRow
									},
									"v": [el2]
								};
							}
							else if(typeof el2 === "boolean") {
								myCell = {
									"$": {
										"r": colToChar(index2 + 1) + currRow,
										"t": "b"
									},
									"v": [(el2 ? 1 : 0)]
								};
							}
							else if(typeof el2 === "string") {
								var str = el2;
								//This...this is weird. I'm used to escaping out characters, but in every test I have done, it looks bad. Leaving this commented out.
								//Because science.
								//str = str.replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");//.replace(/&/g, "&amp;");//.replace(/'/g, "&pos;");
								if(str.charAt(0) === "=") {
									str = str.substr(1);
									myCell = {
										"$": {
											"r": colToChar(index2 + 1) + currRow
										},
										"f": [str]
									};
								}
								else {
									var ind = sharedStrings.indexOf(str);
									if(ind === -1) ind = sharedStrings.push(str) - 1;
									myCell = {
										"$": {
											"r": colToChar(index2 + 1) + currRow,
											"t": "s"
										},
										"v": [ind]
									};
								}
							}
							//Care of http://stackoverflow.com/questions/643782/how-to-know-if-an-object-is-a-date-or-not-with-javascript
							else if(Object.prototype.toString.call(el2) === "[object Date]") {
								var myDate = new Date(Date.UTC(1899,11,30));
								myCell = {
									"$": {
										"r": colToChar(index2 + 1) + currRow,
										"s": "1"
									},
									"v": [Math.floor((el2 - myDate) / (24*60*60*1000))]
								};
							}
							else if(el2 instanceof cell) {
							}
							if(myCell !== null) obj.c.push(myCell);
						})
						var row = sheetData[0].row || [];
						row.push(obj);
						sheetData[0] = {row: row};
					}
				});

				
				var builder = new xml2js.Builder();
				sheetText = builder.buildObject(result);
			});
			myXlsx.file("xl/worksheets/sheet" + count + ".xml", sheetText);
		}
		
		
		var sharedStringsStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sharedStrings.length + '" uniqueCount="' + sharedStrings.length + '">';
		sharedStrings.forEach(function(el) {
			sharedStringsStr += "<si><t>" + el + "</t></si>";
		});
		sharedStringsStr += '</sst>';
		myXlsx.file("xl/sharedStrings.xml", sharedStringsStr);
		var workbook_xml_rels = myXlsx.file("xl/_rels/workbook.xml.rels").asText();
		xml2js.parseString(workbook_xml_rels, function(err, result) {
			var relations = result["Relationships"]["Relationship"];
			var id = rand(6);
			var obj = {
				"$": {
					Id: "rId1",
					Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
					Target: "sharedStrings.xml"
				}
			}
			relations.push(obj);
			var builder = new xml2js.Builder();
			workbook_xml_rels = builder.buildObject(result);
		});
		myXlsx.file("xl/_rels/workbook.xml.rels", workbook_xml_rels);

		var results = myXlsx.generate({type: "string", compression: "DEFLATE"});
		if(fileName) {
			require("fs").writeFileSync(fileName, results, "binary");
			if(require("fs").existsSync(fileName)) return this;
			else throw new Error("Could not write to file " + fileName);
		}
		return results;
	}
	

	this.addSheet = function(userConfig, label) {
		var config = extend({rows: []}, userConfig);
		sheetCount++;
		if(registered[label]) throw new Error("Sheet already exists with label " + label);
		if(!label) label = "Sheet" + sheetCount;
		if(activeTab === "") activeTab = label;
		var id = rand(6);
		registered[label] = {
			config: config,
			id: id
		}
		labels.push(label);
		return (function sheet(myLabel, _this) {
			if(!(this instanceof sheet)) return new sheet(myLabel, _this);
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
			return this;
		})(label, this);
	}
	
	this.removeSheet = function(label) {
		if(!label) label = labels[labels.length - 1];
		if(label === "" || !registered[label]) throw new Error("Invalid sheet [" + label + "]");
		delete registered[label];
		var ind = labels.indexOf(label);
		labels.splice(ind, 1);
		if(label === activeTab) {
			if(labels[ind]) activeTab = labels[ind];
			else if(labels[ind - 1]) activeTab = labels[ind - 1];
			else activeTab = "";
		}
		return this;
	}
	
	this.readXlsx = function(filename) {
		var file = require("fs").readFileSync(filename);
		var myXlsx = JSZip(file, {base64: false, checkCRC32: true});
		var results = {
			sheets: {}
		}
		var sharedStrings = [];
		var s;
		if(s = myXlsx.file("xl/sharedStrings.xml")) {
			var sharedStringsXml = s.asText();
			xml2js.parseString(sharedStringsXml, function(err, result) {
				result["sst"].si.forEach(function(el) {
					sharedStrings.push(el.t[0]);
				});
				
			});
		}
		var idToSheet1 = {};
		if(s = myXlsx.file("xl/workbook.xml")) {
			var workbookXml = s.asText();
			xml2js.parseString(workbookXml, function(err, result) {
				result.workbook.sheets[0].sheet.forEach(function(sheet) {
					idToSheet1[sheet.$["r:id"]] = sheet.$.name;
				});
			});
		}
		var targets = {};
		if(s = myXlsx.file("xl/_rels/workbook.xml.rels")) {
			var workbook_xml_rels = s.asText();
			xml2js.parseString(workbook_xml_rels, function(err, result) {
				result["Relationships"]["Relationship"].forEach(function(ship) {
					var id = ship.$.Id;
					if(idToSheet1[id]) targets[id] = ship.$.Target;
				});
			});
		}
		for(id in targets) {
			if(s = myXlsx.file("xl/" + targets[id])) {
				var sheet = s.asText();
				var label = idToSheet1[id];
				results.sheets[label] = [];
				xml2js.parseString(sheet, function(err, result) {
					result["worksheet"]["sheetData"][0].row.forEach(function(row) {
						var newRow = [];
						row.c.forEach(function(cell) {
							var col = cell.$.r.match(/[a-zA-Z]*/g)[0];
							var rowLen = charToCol(col);
							while(newRow.length < rowLen - 1) newRow.push("");
							var type = cell.$.t;
							if(!type) newRow.push(parseFloat(cell.v[0]));
							else {
								if(type === "b") {
									newRow.push(cell.v[0] === "1");
								}
								else if(type === "s") {
									var ind = parseInt(cell.v[0]);
									newRow.push(sharedStrings[ind]);
								}
							}
						});
						results.sheets[label].push(newRow);
					});
				});
			}
		}
		for(var i in results.sheets) {
			this.addSheet({rows: results.sheets[i]}, i);
		}
		delete file;
		delete myXlsx;
		delete sheets;
		delete sharedStrings;
		delete idToSheet1;
		delete s;
		delete targets;
		delete results;
		return this;
	}
	
	this.toObject = function() {
		var t = registered;
		return t;
	}

	//This only does the Active Tab.
	this.toCSV = function(fileName) {
		var s;
		if(s = registered[activeTab]) {
			var str = [];
			s.config.rows.forEach(function(row) {
				str.push(row.join(","));
			});
			str = str.join("\r\n");
			if(fileName) {
				require("fs").writeFileSync(fileName, str, "utf8");
				if(require("fs").existsSync(fileName)) return this;
				else throw new Error("Could not write to file " + fileName);
			}
			return str;
		}
		else throw new Error("toCSV failed: tab '" + activeTab + "' does not exist.");
	}
	
	return this;
}

node_workbook.prototype.getTemplate = function() {
	var template = require("fs").readFileSync("template.zip");
	var myXlsx = JSZip();
	myXlsx.load(template);
	console.log(myXlsx.generate({base64: true}));
}

module.exports = global.node_workbook = node_workbook;