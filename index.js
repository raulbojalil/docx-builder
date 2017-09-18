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
	
	this.mediaFiles = [];
	this.styles = [];
	
	this.getExternalDocxRawXml = function(docxData)
	{
		var zip = new JSZip(docxData);
	    var xml = Utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
		var stylesXml = Utf8ArrayToString(zip.file("word/styles.xml")._data.getContent());
		
		stylesXml = stylesXml.substring(stylesXml.indexOf("<w:styles"));
		stylesXml = stylesXml.substring(stylesXml.indexOf(">") + 1);
		stylesXml = stylesXml.substring(0, stylesXml.indexOf("</w:styles>"));
		
		this.styles.push(stylesXml);
		
		var mediaFolderName = "word/media";
		var mediaFolder = zip.folder(mediaFolderName);
		
		var relsXml = Utf8ArrayToString(zip.file("word/_rels/document.xml.rels")._data.getContent());
		
		for(var file in mediaFolder.files)
		{
			if(file.startsWith("word/media") && file != "word/media/")
			{
				var oldRId = "";
				var newRId = "";
				var rType = "";
				var newFile = file;
		
				var indexOfOldRel = relsXml.indexOf(file.substr(5));
		        if(indexOfOldRel != -1)
				{
					var left = indexOfOldRel;
					var right = indexOfOldRel;
					
					while(left > 0) { left--; if(relsXml[left] == '<') break; }
					while(right < relsXml.length) { right++; if(relsXml[right] == '>') break; }
					
					var relTag = relsXml.substr(left);
					relTag = relTag.substr(0,right-left+1).split(' ');
					
					for(var i=0; i < relTag.length; i++)
					{
						var item = relTag[i];
						if(item.startsWith('Id="'))
						{
							oldRId = item.substr(4);
							oldRId = oldRId.substr(0, oldRId.length - 1);
						}
						else if(item.startsWith('Type="'))
						{
							rType = item.substr(6);		
							rType = rType.substr(0, rType.length - 1);
						}							
					}
										
					var hrTime = process.hrtime();
					var newId = hrTime[0] + "" + hrTime[1];
					newRId = "rId" + newId;
					
					var fileExt = "." + newFile.split('.').pop();
					newFile = newFile.substr(0, newFile.length - fileExt.length) + newId + fileExt;
					
					xml = xml.replace("\"" + oldRId  + "\"", "\"" + newRId  + "\"");
					this.mediaFiles.push({ name: newFile, data: mediaFolder.files[file]._data, rId: newRId, rType: rType });
					
				}
			}
		}

		xml = xml.substring(xml.indexOf("<w:body>") + 8);
        xml = xml.substring(0, xml.indexOf("</w:body>"));
		xml = xml.substring(0, xml.indexOf("<w:sectPr"));
		
		return xml;
	}
	
	this.insertDocxSync = function(path){
		
		var xml = this.getExternalDocxRawXml(fs.readFileSync(path,"binary"));
		this.insertRaw(xml);
	}
	
	this.insertDocx = function(path, callback){
		
		fs.readFile(path, "binary", (e, data) => {
		  if (e) callback(e);
		  else
		  {
			var xml = this.getExternalDocxRawXml(data);
			this.insertRaw(xml);
			callback(null);
		  }
		});
	}
	
	this.save = function(filepath, err){
		
		var template = fs.readFileSync(__dirname + "/template.docx","binary");
		var zip = new JSZip(template);

		
		if(this.mediaFiles.length > 0)
		{
			var relsXml = Utf8ArrayToString(zip.file("word/_rels/document.xml.rels")._data.getContent());
			
			for(var i=0; i < this.mediaFiles.length; i++)
			{
				var mediaFile = this.mediaFiles[i];
				zip.file(mediaFile.name, mediaFile.data);
			    relsXml = relsXml.replace('</Relationships>', '<Relationship Id="' + mediaFile.rId + '" Type="' + mediaFile.rType + '" Target="' + mediaFile.name.substr(5) + '"/></Relationships>');
			}
			
			zip.file("word/_rels/document.xml.rels", relsXml);
		}
		
		if(this.styles.length > 0)
		{
			var stylesXml = Utf8ArrayToString(zip.file("word/styles.xml")._data.getContent()).replace("</w:styles>", "");
			zip.file("word/styles.xml", stylesXml + this.styles.join("") + "</w:styles>");
		}
		
		
	    //zip.file("word/media/image1.png", algo._data);
		
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