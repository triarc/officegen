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

		

		var FILENAME = "test-docx-border-top.docx";
		var out = fs.createWriteStream(OUTDIR + FILENAME);
		docx.generate(out);
		out.on ( 'close', function () {
			done ();
		});
	});

});
