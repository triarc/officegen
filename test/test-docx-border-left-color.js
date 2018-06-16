/*

	test-docx-border-left-color.js

	Example of usage of colored left border

	Author:		Ilia Ashmarin, mail[at]harduino.com

	History:	2018.06.16 - File created

*/
var assert = require('assert');
var officegen = require('../');
var fs = require('fs');
var path = require('path');

var IMAGEDIR = __dirname + "/../examples/";
var OUTDIR = __dirname + "/../tmp/";

var AdmZip = require('adm-zip');


var docxEquivalent = function (path1, path2, subdocs) {
	var left = new AdmZip(path1);
	var right = new AdmZip(path2);
	for (var i = 0; i < subdocs.length; i++) {
		if (left.readAsText(subdocs[i]) != right.readAsText(subdocs[i])) {
			return false;
		}
	}
	return true;
}

// Common error method
var onError = function (err) {
	console.log(err);
	assert(false);
	done()
};

describe("DOCX generator", function () {
	this.timeout(1000);

	it("creates a document with text and styles", function (done) {

		var docx = officegen ( 'docx' );
		docx.on ( 'error', onError );

		var pObj = docx.createP ();

		pObj.addText ( 'Simple' );
		pObj.addText ( ' with color', { color: '000088' } );
		pObj.addText ( ' and back color.', { color: '00ffff', back: '000088' } );

		var pObj = docx.createP ();

		pObj.addText ( 'Bold + underline', { bold: true, underline: true } );

		var pObj = docx.createP ( { align: 'center' } );

		pObj.addText ( 'Center this text.' );

		var pObj = docx.createP ();
		pObj.options.align = 'right';

		pObj.addText ( 'Align this text to the right.' );

		var pObj = docx.createP ();

		pObj.addText ( 'Those two lines are in the same paragraph,' );
		pObj.addLineBreak ();
		pObj.addText ( 'but they are separated by a line break.' );

		docx.putPageBreak ();

		var pObj = docx.createP ();

		pObj.addText ( 'Fonts face only.', { font_face: 'Arial' } );
		pObj.addText ( ' Fonts face and size.', { font_face: 'Arial', font_size: 40 } );

		docx.putPageBreak ();

		var pObj = docx.createListOfNumbers ();

		pObj.addText ( 'Option 1' );

		var pObj = docx.createListOfNumbers ();

		pObj.addText ( 'Option 2' );

		var FILENAME = "test-docx-border-top.docx";
		var out = fs.createWriteStream(OUTDIR + FILENAME);
		docx.generate(out);
		out.on ( 'close', function () {
			done ();
		});
	});

});
