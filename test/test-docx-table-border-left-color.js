/*

	test-docx-table-border-left-color.js

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

		var table = [
		  [{
		    val: "No.",
		    opts: {
		      cellColWidth: 4261,
		      b:true,
		      sz: '48',
		      shd: {
		        fill: "7F7F7F",
		        themeFill: "text1",
		        "themeFillTint": "80"
		      },
		      fontFamily: "Avenir Book"
		    }
		  },{
		    val: "Title1",
		    opts: {
		      b:true,
		      color: "A00000",
		      align: "right",
		      shd: {
		        fill: "92CDDC",
		        themeFill: "text1",
		        "themeFillTint": "80"
		      }
		    }
		  },{
		    val: "Title2",
		    opts: {
		      align: "center",
		      vAlign: "center",
		      cellColWidth: 42,
		      b:true,
		      sz: '48',
		      shd: {
		        fill: "92CDDC",
		        themeFill: "text1",
		        "themeFillTint": "80"
		      }
		    }
		  }],
		  [1,'All grown-ups were once children',''],
		  [2,'there is no harm in putting off a piece of work until another day.',''],
		  [3,'But when it is a matter of baobabs, that always means a catastrophe.',''],
		  [4,'watch out for the baobabs!','END'],
		]

		var tableStyle = {
			tableColWidth: 4261,
			tableSize: 24,
			tableColor: "444444",
			tableAlign: "left",
			borders: true, // enable borders in table
			borderColor: "444444", // color for border
			borderSize: "12", // size of border width
			bordersInsideH:false, //do not remove horizontal borders from inside table
			bordersInsideV:true, //remove vertically borders from inside table
		}

		docx.createTable (table, tableStyle);

		var FILENAME = "test-docx-border-top.docx";
		var out = fs.createWriteStream(OUTDIR + FILENAME);
		docx.generate(out);
		out.on ( 'close', function () {
			done ();
		});
	});

});
