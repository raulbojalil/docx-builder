# docx-builder

An NPM module for creating and/or merging DOCX files.

# How to install

Using [npm](https://www.npmjs.com/) do:

```
npm install docx-builder
```

# Example

``` js
var builder = require('docx-builder');
var docx = new builder.Document();

//SET THE HEADER

docx.beginHeader();
docx.insertText("This is the header");
docx.endHeader();

//SET THE FOOTER

docx.beginFooter();
docx.insertText("This is the footer");
docx.endFooter();

//COMMON OPERATIONS

docx.setBold();
docx.unsetBold();
docx.setItalic();
docx.unsetItalic();
docx.setUnderline();
docx.unsetUnderline();
docx.setFont("Arial");
docx.setSize(40);
docx.rightAlign();
docx.centerAlign();
docx.leftAlign();
docx.insertText("Hello");

//INSERT A TABLE

docx.beginTable({ borderColor: "red", borderSize: 20 });
docx.insertRow();
docx.setItalic();
docx.setUnderline();
docx.insertText("ID");
docx.nextColumn();
docx.insertText("Username");
docx.nextRow();
docx.insertText("1");
docx.nextColumn();
docx.insertText("raulb");
docx.endTable();

//INSERT EXTERNAL FILES

docx.insertDocxSync(__dirname+"/test.docx");
docx.insertPageBreak();
docx.insertDocx(__dirname+"/test2.docx", function(err){
	if(err) console.log(err);
});

//SAVE THE DOCX FILE

docx.save(__dirname + "/output.docx", function(err){
	if(err) console.log(err);
});

```

# Merging .docx files

A common requirement is to merge two or more .docx files into one. For that, just call
`insertDocxSync` or `insertDocx` to append the content of your .docx files.

``` js
var builder = require('docx-builder');
var docx = new builder.Document();

docx.insertDocxSync(__dirname+"/test1.docx");
docx.insertDocxSync(__dirname+"/test2.docx");
docx.insertDocxSync(__dirname+"/test3.docx");

//SAVING THE DOCX FILE

docx.save(__dirname + "/output.docx", function(err){
	if(err) console.log(err);
});

```
