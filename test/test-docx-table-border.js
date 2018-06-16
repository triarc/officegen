/*

	test-docx-table-border.js

	Example of usage of table border

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

		var opts_header = {
			b:true, /* bold enabled */
			color: "FFFFFF", /* text color */
			align: "center", /* text align */
			shd: {
				fill: "98BE4D" /* background's fill color */
			}
		};
		var table = [
			[
				{
					val:	"Section/Question",
					opts:	opts_header
				},
				{
					val:	"Privacy Index",
					opts:	opts_header
				},
				{
					val:	"Risk",
					opts:	opts_header
				},
				{
					val:	"Risk Level",
					opts:	opts_header
				},
				{
					val:	"Remediation",
					opts:	opts_header
				},
				{
					val:	"No.",
					opts:	opts_header
				}
			],
			[
				{
					val: "S1/Q3"
				},
				{
					val: "w01"
				},
				{
					val: "abc"
				},
				{
					val: "medium"
				},
				{
					val: "Please contact your DPO"
				},
				{
					val: "1"
				}
			],
			[
				{
					val: "3-1"
				},
				{
					val: "3-2",
					opts: {
						align:	'center',
						vAlign: 'center'
					}
				},
				{
					val: "3-3"
				},
				{
					val: "3-4"
				},
				{
					val: "3-5"
				},
				{
					val: "3-6"
				}
			]
		];
		var tableStyle = {
			tableColWidth: 4261,
			tableSize: 24,
			tableColor: "444444",
			tableAlign: "left",
			borders: true, // enable borders in table
			borderColor: "98BE4D", // color for border
			borderSize: "12", // size of border width
			bordersInsideH:true, //do not remove horizontal borders from inside table
			bordersInsideV:true, //do not remove vertically borders from inside table
		}
		docx.createTable (table, tableStyle);

		var FILENAME = "test-docx-table-border.docx";
		var out = fs.createWriteStream(OUTDIR + FILENAME);
		docx.generate(out);
		out.on ( 'close', function () {
			done ();
		});
	});

});
