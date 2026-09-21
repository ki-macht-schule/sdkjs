/*
 * Copyright (C) KI macht Schule gGmbH, 2026
 *
 * New file in a modified version of ONLYOFFICE Docs. Licensed under the
 * GNU Affero General Public License version 3, together with the
 * additional terms in the LICENSE file of the ONLYOFFICE Document
 * Server distribution: this file is combined with that work and is
 * distributed under its terms.
 *
 * Distributed WITHOUT ANY WARRANTY. See the GNU AGPL for details:
 * https://www.gnu.org/licenses/agpl-3.0.html
 *
 * The delta against upstream, and where to get the complete source, is
 * described in KIWI-CHANGES.md next to this file.
 */
"use strict";

var assert = require("assert");
var remount = require("./PresentationThemeRemount.js");

var passed = 0;
["strictEqual", "deepStrictEqual", "ok"].forEach(function (name) {
	var orig = assert[name];
	assert[name] = function () {
		orig.apply(assert, arguments);
		passed++;
	};
});

var kimLayout = {type: 1, matchingName: "T", cSld: {name: "Title"}};
var officeLayout = {type: 1, matchingName: "T", cSld: {name: "Title"}};
var kimMaster = {
	Theme: {name: "KImS Light Theme"},
	sldLayoutLst: [kimLayout],
	getMatchingLayout: function () { return kimLayout; }
};
var officeMaster = {Theme: {name: "Office Theme"}, sldLayoutLst: [officeLayout]};
kimLayout.Master = kimMaster;
officeLayout.Master = officeMaster;
var hijacked = {Layout: officeLayout, setLayout: function (layout) { this.Layout = layout; }};
var themed = {
	Slides: [hijacked],
	slideMasters: [officeMaster, kimMaster]
};
assert.strictEqual(remount.remountOfficeHijacks(themed), 1);
assert.strictEqual(hijacked.Layout, kimLayout);
assert.strictEqual(remount.remountOfficeHijacks(themed), 0);

console.log("PresentationThemeRemount: " + passed + " checks passed");
