var fs = require('fs');
var Docxtemplater = require('docxtemplater');
var JSZip = require('jszip');

exports.Document = function() {
	
	this._body = [];
	this._header = [];
	this._footer = [];
	this._builder = this._body;
    this._bold = false;
	this._italic = false;
	this._underline = false;
	this._font = null;
	this._size = null;
	this._alignment = null;
	
	
	this.beginHeader = function() 
	{
		this._builder = this._header;
	}
	
	this.endHeader = function()
	{
		this._builder = this._body;
	}
	
	this.beginFooter = function() 
	{
		this._builder = this._footer;
	}
	
	this.endFooter = function()
	{
		this._builder = this._body;
	}
	
	this.setBold = function(){
		this._bold = true;
	}
	
	this.unsetBold = function(){
		this._bold = false;
	}
	
	this.setItalic = function(){
		
		this._italic = true;
	}
	
	this.unsetItalic = function(){
		
		this._italic = false;
	}
	
	this.setUnderline = function(){
		
		this._underline = true;
	}
	
	this.unsetUnderline = function(){
		
		this._underline = false;
	}
	
	this.setFont = function(font){
		this._font = font;
	}
	
	this.unsetFont = function() {
		this._font = null;
	}
	
	this.setSize = function(size){
		this._size = size;
	}
	
	this.unsetSize = function(){
		this._size = null;
	}
	
	this.rightAlign = function(){
		this._alignment = "right";
	}
	
	this.centerAlign = function(){
		this._alignment = "center";
	}
	
	this.leftAlign = function(){
		this._alignment = null;
	}
	
	this.insertPageBreak = function()
	{
		var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';
				  
		this._builder.push(pb);
	}
	
	this.beginTable = function(options){
		
		if(!options)
		{
			this._builder.push('<w:tbl>');
		}
		else
		{
			options = options || { borderSize: 4, borderColor: 'auto' };
			this._builder.push('<w:tbl><w:tblPr><w:tblBorders> \
				<w:top w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:left w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:bottom w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:right w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:insideH w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:insideV w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				</w:tblBorders>	\
			</w:tblPr>');
		}
	}
	
	this.insertRow = function(){
		
		this._builder.push('<w:tr><w:tc>');
	}
	
	this.nextColumn = function(){
		this._builder.push('</w:tc><w:tc>');
	}
	
	this.nextRow = function(){
		this._builder.push('</w:tc></w:tr><w:tr><w:tc>');
	}
	
	this.endTable = function(){
		this._builder.push('</w:tc></w:tr></w:tbl>');
	}
	
    this.insertText = function(text) {
		
		var p = '<w:p>' +
		
			(this._alignment ? ('<w:pPr><w:jc w:val="' + this._alignment + '"/></w:pPr>') : '') +
			
			'<w:r> \
				<w:rPr>' +
				
				    (this._size ? ('<w:sz w:val="' + this._size + '"/>') : "") +
					(this._bold ? '<w:b/>' : "") +
					(this._italic ? '<w:i/>' : "") +
					(this._underline ? '<w:u w:val="single"/>' : "") +
					(this._font ? ('<w:rFonts w:hAnsi="' + this._font + '" w:ascii="' + this._font + '"/>') : "")					
					
					+
				'</w:rPr> \
				<w:t>[CONTENT]</w:t> \
			</w:r> \
		</w:p>'
		
        this._builder.push(p.replace("[CONTENT]", text));
    }
	
	this.insertRaw = function(xml){
		
		this._builder.push(xml);
	}
	
	this.insertDocxSync = function(path){
		
		var zip = new JSZip(fs.readFileSync(path,"binary"));
	    var xml = Utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
		xml = xml.substring(xml.indexOf("<w:body>") + 8);
        xml = xml.substring(0, xml.indexOf("</w:body>"));
		xml = xml.substring(0, xml.indexOf("<w:sectPr"));
		this.insertRaw(xml);
	}
	
	this.insertDocx = function(path, callback){
		
		fs.readFile(path, "binary", (e, data) => {
		  if (e) callback(e);
		  else
		  {
			  
			var zip = new JSZip(data);
			var xml = Utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
			xml = xml.substring(xml.indexOf("<w:body>") + 8);
			xml = xml.substring(0, xml.indexOf("</w:body>"));
			xml = xml.substring(0, xml.indexOf("<w:sectPr"));
			this.insertRaw(xml);
			callback(null);
		  }
		});
	}
	
	this.save = function(filepath, err){
		
		var template = fs.readFileSync(__dirname + "/template.docx","binary");
		var zip = new JSZip(template);
	
		var doc = new Docxtemplater().loadZip(zip);

		doc.setData({body: this._body.join(''), header: this._header.join(''), footer: this._footer.join('') });
		doc.render();

		var buf = doc.getZip().generate({type:"nodebuffer"});
		fs.writeFile(filepath,buf, err);
	}
}


function Utf8ArrayToString(array) {
    var out, i, len, c;
    var char2, char3;

    out = "";
    len = array.length;
    i = 0;
    while(i < len) {
    c = array[i++];
    switch(c >> 4)
    { 
      case 0: case 1: case 2: case 3: case 4: case 5: case 6: case 7:
        // 0xxxxxxx
        out += String.fromCharCode(c);
        break;
      case 12: case 13:
        // 110x xxxx   10xx xxxx
        char2 = array[i++];
        out += String.fromCharCode(((c & 0x1F) << 6) | (char2 & 0x3F));
        break;
      case 14:
        // 1110 xxxx  10xx xxxx  10xx xxxx
        char2 = array[i++];
        char3 = array[i++];
        out += String.fromCharCode(((c & 0x0F) << 12) |
                       ((char2 & 0x3F) << 6) |
                       ((char3 & 0x3F) << 0));
        break;
    }
    }

    return out;
}