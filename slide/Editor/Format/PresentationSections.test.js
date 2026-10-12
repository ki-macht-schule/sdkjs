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
var sections = require("./PresentationSections.js");

var passed = 0;
function ok(cond, what) {
	assert.ok(cond, what);
	passed++;
}
function eq(a, b, what) {
	assert.deepStrictEqual(a, b, what);
	passed++;
}

// --- a PPTY writer/reader good enough to prove the roundtrip ---------

function Writer() {
	this.data = Buffer.alloc(4096);
	this.pos = 0;
	this.stack = [];
	this.depth = 0;
}
Writer.prototype.grow = function (n) {
	if (this.pos + n <= this.data.length)
		return;
	var bigger = Buffer.alloc(Math.max(this.data.length * 2, this.pos + n));
	this.data.copy(bigger);
	this.data = bigger;
};
Writer.prototype.WriteUChar = function (v) {
	this.grow(1);
	this.data[this.pos++] = v & 0xFF;
};
Writer.prototype.WriteULong = function (v) {
	this.grow(4);
	for (var i = 0; i < 4; i++)
		this.data[this.pos++] = (v >>> (8 * i)) & 0xFF;
};
Writer.prototype.WriteString2 = function (t) {
	t = String(t);
	this.WriteULong(t.length);
	this.grow(t.length * 2);
	for (var i = 0; i < t.length; i++) {
		var c = t.charCodeAt(i) & 0xFFFF;
		this.data[this.pos++] = c & 0xFF;
		this.data[this.pos++] = (c >>> 8) & 0xFF;
	}
};
Writer.prototype._WriteString2 = function (type, v) {
	if (v != null) {
		this.WriteUChar(type);
		this.WriteString2(v);
	}
};
Writer.prototype.StartRecord = function (t) {
	this.stack[this.depth++] = this.pos + 5;
	this.WriteUChar(t);
	this.WriteULong(0);
};
Writer.prototype.EndRecord = function () {
	this.depth--;
	var seek = this.pos;
	this.pos = this.stack[this.depth] - 4;
	this.WriteULong(seek - this.stack[this.depth]);
	this.pos = seek;
};
Writer.prototype.WriteRecord1 = function (t, v, f) {
	this.StartRecord(t);
	f(v);
	this.EndRecord();
};
Writer.prototype.WriteRecordArray = function (t, st, arr, f) {
	this.StartRecord(t);
	this.WriteULong(arr.length);
	for (var i = 0; i < arr.length; i++)
		this.WriteRecord1(st, arr[i], f);
	this.EndRecord();
};

function Reader(buf) {
	this.data = buf;
	this.cur = 0;
}
Reader.prototype.GetUChar = function () { return this.data[this.cur++]; };
Reader.prototype.GetULong = function () {
	var v = this.data[this.cur] | (this.data[this.cur + 1] << 8) |
		(this.data[this.cur + 2] << 16) | (this.data[this.cur + 3] << 24);
	this.cur += 4;
	return v >>> 0;
};
Reader.prototype.GetString2 = function () {
	var len = this.GetULong() * 2;
	var t = "";
	for (var i = 0; i < len; i += 2) {
		var c = this.data[this.cur + i] | (this.data[this.cur + i + 1] << 8);
		if (c === 0)
			break;
		t += String.fromCharCode(c);
	}
	this.cur += len;
	return t;
};
Reader.prototype.Seek2 = function (p) { this.cur = p; return 0; };
Reader.prototype.Skip2 = function (n) { return n < 0 ? 1 : this.Seek2(this.cur + n); };
Reader.prototype.SkipRecord = function () { this.Skip2(this.GetULong()); };

// --- helpers --------------------------------------------------------

// The two upstream pieces addSection() leans on. Stubbing them here is
// the point: if either signature moves, this test says so.
function CPrSection() {
	this.name = null;
	this.guid = null;
	this.startIndex = null;
}
CPrSection.prototype.setName = function (v) { this.name = v; };
CPrSection.prototype.setGuid = function (v) { this.guid = v; };
CPrSection.prototype.setStartIndex = function (v) { this.startIndex = v; };
global.AscCommonSlide = global.AscCommonSlide || {};
global.AscCommonSlide.CPrSection = CPrSection;

function deck(n, cuts) {
	var slides = [];
	for (var i = 0; i < n; i++)
		slides.push({i: i});
	var secs = (cuts || []).map(function (c) {
		var s = new CPrSection();
		s.name = c.name;
		s.guid = "{" + c.name + "}";
		s.startIndex = c.at;
		return s;
	});
	return {
		Slides: slides,
		Sections: secs,
		addSection: function (pos, pr) { this.Sections.splice(pos, 0, pr); },
		removeSection: function (pos) { this.Sections.splice(pos, 1); }
	};
}

function shape(p) {
	return p.Sections.map(function (s) {
		return s.name + "@" + s.startIndex + ":" + sections.membersOf(p, s).length;
	}).join(" ");
}

function roundtrip(p) {
	var w = new Writer();
	sections.write(w, p);
	if (w.pos === 0)
		return null;
	var r = new Reader(w.data);
	assert.strictEqual(r.GetUChar(), 7);
	return sections.read(r);
}

// --- record 7 -------------------------------------------------------

var p = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
var pending = roundtrip(p);
eq(pending.length, 2, "two sections survive the roundtrip");
eq(pending[0].sldIds, ["256", "257"], "A owns the first two ids");
eq(pending[1].sldIds, ["258", "259"], "B owns the last two");
eq(pending[0].name, "A", "the name survives");
ok(pending[0].guid === "{A}", "the guid survives");

var bound = sections.bind(p.Slides, pending);
eq(bound.sections.map(function (s) { return s.startIndex; }), [0, 2], "bind recovers the cuts");
eq(bound.mapped, 4, "every id bound");

// ids that are not 256-based (a deck PowerPoint renumbered)
var ppt = sections.bind(p.Slides, [
	{name: "A", guid: "{A}", sldIds: ["300", "301"]},
	{name: "B", guid: "{B}", sldIds: ["302", "303"]}
]);
eq(ppt.mapped, 0, "nothing binds");
var unbound = deck(4, []);
unbound.pendingSections = [{name: "A", guid: "{A}", sldIds: ["300", "301"]}];
sections.apply(unbound);
ok(unbound.sectionsUnbound === true, "a file we cannot read is flagged");
ok(roundtrip(unbound) === null, "and it is not written back");

var partial = deck(4, []);
partial.pendingSections = [
	{name: "A", guid: "{A}", sldIds: ["256", "257"]},
	{name: "B", guid: "{B}", sldIds: ["260", "261"]}
];
sections.apply(partial);
ok(partial.sectionsPartial === true, "a half-read file is flagged");
ok(roundtrip(partial) === null, "and it is not written back either");

// --- normalise ------------------------------------------------------

var messy = deck(4, [{name: "B", at: 2}, {name: "A", at: 0}]);
sections.normalise(messy);
eq(shape(messy), "A@0:2 B@2:2", "cuts are sorted");

// An emptied section is kept, so the user can fill it again. The LAST
// cut at an index owns the slides that follow.
var dup = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}, {name: "C", at: 2}]);
sections.normalise(dup);
eq(shape(dup), "A@0:2 B@2:0 C@2:2", "an empty section keeps its place");

var far = deck(3, [{name: "A", at: 0}, {name: "B", at: 9}]);
sections.normalise(far);
eq(shape(far), "A@0:3 B@3:0", "a cut past the end clamps to the end and empties");

var floating = deck(3, [{name: "A", at: 1}]);
sections.normalise(floating);
eq(shape(floating), "A@0:3", "the first section always owns slide 0");

// --- derived fixups -------------------------------------------------

var ins = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
ins.Slides.splice(2, 0, {i: 9});
sections.onInsert(ins, 2, 1, ins.Sections[0]);
eq(shape(ins), "A@0:3 B@3:2", "naming A as the owner keeps the new slide in A");

ins = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
ins.Slides.splice(2, 0, {i: 9});
sections.onInsert(ins, 2, 1, ins.Sections[1]);
eq(shape(ins), "A@0:2 B@2:3", "naming B starts B with it");

// no owner named: the slide joins the section it follows. This is the
// duplicate case -- a copy belongs with its original, even when the
// original is the last slide of its section.
ins = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
ins.Slides.splice(2, 0, {i: 9});
sections.onInsert(ins, 2, 1);
eq(shape(ins), "A@0:3 B@3:2", "a duplicate stays with its original");

var lone = deck(3, [{name: "A", at: 0}, {name: "B", at: 1}]);
lone.Slides.splice(1, 0, {i: 9});
sections.onInsert(lone, 1, 1);
eq(shape(lone), "A@0:2 B@2:2", "even when the section held exactly one slide");

var del = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
del.Slides.splice(3, 1);
sections.onDelete(del, 3);
eq(shape(del), "A@0:2 B@2:1", "a delete below a cut moves it");

// the special case: a section with exactly one slide
var solo = deck(3, [{name: "A", at: 0}, {name: "B", at: 1}, {name: "C", at: 2}]);
eq(shape(solo), "A@0:1 B@1:1 C@2:1", "three sections of one slide each");
solo.Slides.splice(1, 1);
sections.onDelete(solo, 1);
eq(shape(solo), "A@0:1 B@1:0 C@1:1", "deleting B's only slide leaves B empty, not gone");

var move = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
var moved = move.Slides.splice(0, 1)[0];
move.Slides.push(moved);
sections.shiftMove(move, [0], 3, 1, move.Sections[1]);
eq(shape(move), "A@0:1 B@1:3", "the moved slide joins B at the end");

var soloMove = deck(4, [{name: "A", at: 0}, {name: "B", at: 1}]);
var gone = soloMove.Slides.splice(0, 1)[0];
soloMove.Slides.push(gone);
sections.shiftMove(soloMove, [0], 3, 1, soloMove.Sections[1]);
eq(shape(soloMove), "A@0:0 B@0:4", "A loses its only slide but keeps its bar");
// and can be filled again
var back = soloMove.Slides.splice(3, 1)[0];
soloMove.Slides.splice(0, 0, back);
sections.shiftMove(soloMove, [3], 0, 1, soloMove.Sections[0]);
eq(shape(soloMove), "A@0:1 B@1:3", "dropping a slide into the empty bar refills it");

var seam = deck(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
var one = seam.Slides.splice(0, 1)[0];
seam.Slides.splice(2, 0, one);
sections.shiftMove(seam, [0], sections.adjustInsertPos(3, [0]), 1, seam.Sections[0]);
eq(shape(seam), "A@0:3 B@3:3", "a drop at the end of A stays in A");

seam = deck(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
one = seam.Slides.splice(0, 1)[0];
seam.Slides.splice(2, 0, one);
sections.shiftMove(seam, [0], sections.adjustInsertPos(3, [0]), 1, seam.Sections[1]);
eq(shape(seam), "A@0:2 B@2:4", "the same drop below the bar joins B");

eq(sections.adjustInsertPos(3, [0]), 2, "upstream decrements the insert index");
eq(sections.adjustInsertPos(1, [4]), 1, "and only for indexes below it");

// --- which section a drop joins --------------------------------------

var ask = deck(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
eq(sections.sectionOfInsert(ask, 3, ask.Sections[0]).name, "A", "the named owner wins");
eq(sections.sectionOfInsert(ask, 3, ask.Sections[1]).name, "B", "either way");
eq(sections.sectionOfInsert(ask, 3).name, "A", "unnamed follows the slide above");
eq(sections.sectionOfInsert(ask, 6).name, "B", "past the end is the last section");
eq(sections.sectionOfInsert(ask, 0).name, "A", "index 0 is always the first");
eq(sections.cutsAt(ask, 3).length, 1, "one cut on the seam");

// --- deliberate edits ------------------------------------------------

var add = deck(6, [{name: "A", at: 0}]);
ok(sections.canAddSection(add, 0) === false, "no cut before the first slide");
ok(sections.canAddSection(add, 6) === false, "and none past the last");
ok(sections.canAddSection(add, 3) === true, "but one in the middle");
var made = sections.addSection(add, 3);
eq(shape(add), "A@0:3 Neuer Abschnitt@3:3", "the new section takes the rest");
ok(sections.canAddSection(add, 3) === false, "the same cut cannot be made twice");
ok(sections.addSection(add, 3) === null, "and is refused");

var second = sections.addSection(add, 5);
eq(second.name, "Neuer Abschnitt 2", "names stay unique");

var units = ["A"];
ok(sections.canRenameSection(made, units, add, 1) === true, "a section the user added may be renamed");
ok(sections.canRenameSection(add.Sections[0], units, add, 1) === false, "a unit's section may not");
ok(sections.canRenameSection(made, null, add, 1) === false, "with no id list, nothing is renameable");
ok(sections.canRenameSection(made, units, add) === false, "without kiwiProtocol, nothing is renameable");
ok(sections.canRenameSection(made, units, add, 1, null) === false, "a missing reserved list locks every name");
var reserved = deck(2, [{name: "_kurs", at: 0}, {name: "u1", at: 1}]);
var hold = ["_kurs"];
ok(sections.canRenameSection(reserved.Sections[0], ["u1"], reserved, 1, []) === true,
	"the name is free until the list says otherwise");
ok(sections.canRenameSection(reserved.Sections[0], ["u1"], reserved, 1, hold) === false,
	"a listed name is not renameable");
ok(sections.canMoveSection(reserved.Sections[0], 1, reserved, hold) === false, "a listed name is not movable");
ok(sections.moveSection(reserved, reserved.Sections[0], 1, hold) === false, "moving it is refused");
ok(sections.canMoveSection(reserved.Sections[1], 0, reserved, hold) === false,
	"nothing moves in front of a reserved first section");
ok(sections.renameSection(reserved, reserved.Sections[0], "Kursfolien", hold) === false,
	"renameSection honours the list");
ok(sections.renameSection(add, made, "Vertiefung") === true, "renaming works");
eq(made.name, "Vertiefung", "and lands");
ok(sections.renameSection(add, second, "Vertiefung") === false, "a duplicate name is refused");

// --- moving a whole section ------------------------------------------

var order = deck(6, [{name: "A", at: 0}, {name: "B", at: 2}, {name: "C", at: 4}]);
order.moveSlides = function (indexes, at) {
	var taken = [];
	for (var i = indexes.length - 1; i >= 0; i--)
		taken.unshift(this.Slides.splice(indexes[i], 1)[0]);
	var dest = at;
	for (i = 0; i < indexes.length; i++)
		if (indexes[i] < at)
			dest--;
	for (i = 0; i < taken.length; i++)
		this.Slides.splice(dest + i, 0, taken[i]);
};
sections.moveSection(order, order.Sections[0], 3);
eq(order.Sections.map(function (s) { return s.name; }), ["B", "C", "A"], "A moved to the end");
eq(shape(order), "B@0:2 C@2:2 A@4:2", "and the cuts were rebuilt");
eq(order.Slides.map(function (s) { return s.i; }), [2, 3, 4, 5, 0, 1], "the slides travelled with it");

// --- undo / redo, through the real History-change path ---------------

/** The shape of CChangesDrawingsContentPresentation that we read. */
function slideChange(presentation, items, pos, isAdd) {
	return {
		Class: presentation,
		Items: items,
		Pos: pos,
		IsAdd: function () { return isAdd; },
		private_GetChangedArray: function () { return presentation.Slides; }
	};
}

// Deleting the FIRST slide of a section and undoing it is the one
// insert whose owner cannot be derived: the slide comes back onto an
// index a cut already sits on. Deriving picks the section above, and
// because derived writes are not in History, the drift is permanent.
var undoDeck = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
var lost = undoDeck.Slides.splice(2, 1)[0];
var del = slideChange(undoDeck, [lost], 2, false);
sections.syncFromChange(del, true);
eq(shape(undoDeck), "A@0:2 B@2:1", "the delete moves nothing, B just gets shorter");
undoDeck.Slides.splice(2, 0, lost);
sections.syncFromChange(del, false);
eq(shape(undoDeck), "A@0:2 B@2:2", "and undo gives B its first slide back, not A");

// redo takes it away again (History removes it, then the hook runs)
undoDeck.Slides.splice(2, 1);
sections.syncFromChange(del, true);
eq(shape(undoDeck), "A@0:2 B@2:1", "redo is the delete once more");

// an ordinary slide, away from any cut: derivation is right
var plainDeck = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
var plain = plainDeck.Slides.splice(3, 1)[0];
var delPlain = slideChange(plainDeck, [plain], 3, false);
sections.syncFromChange(delPlain, true);
eq(shape(plainDeck), "A@0:2 B@2:1", "B loses its second slide");
plainDeck.Slides.splice(3, 0, plain);
sections.syncFromChange(delPlain, false);
eq(shape(plainDeck), "A@0:2 B@2:2", "and gets it back");

// a multi-slide delete moves the cuts by as many places as it removed
var manyDeck = deck(5, [{name: "A", at: 0}, {name: "B", at: 3}]);
var many = manyDeck.Slides.splice(0, 2);
sections.syncFromChange(slideChange(manyDeck, many, 0, false), true);
eq(shape(manyDeck), "A@0:1 B@1:2", "two slides gone, the cut moved two");

// an add (paste, duplicate) arrives with no recorded owner
var addDeck = deck(4, [{name: "A", at: 0}, {name: "B", at: 2}]);
var fresh = {i: 9};
addDeck.Slides.splice(2, 0, fresh);
sections.syncFromChange(slideChange(addDeck, [fresh], 2, true), true);
eq(shape(addDeck), "A@0:3 B@3:2", "it joins the section of the slide it follows");

// a change on some other array is none of our business
var other = deck(3, [{name: "A", at: 0}]);
sections.syncFromChange({
	Class: other,
	Items: [{}],
	Pos: 0,
	IsAdd: function () { return false; },
	private_GetChangedArray: function () { return other.Sections; }
}, true);
eq(shape(other), "A@0:3", "a Sections change is not a slide change");

// --- the prototype wrappers ------------------------------------------

/** Enough of CPresentation for hookPresentation to wrap. */
function CPresentation(n, cuts) {
	var base = deck(n, cuts);
	this.Slides = base.Slides;
	this.Sections = base.Sections;
	this.editable = true;
	this.mixed = false;
}
CPresentation.prototype.CanEdit = function () { return this.editable; };
CPresentation.prototype.IsMasterMode = function () { return false; };
CPresentation.prototype.addSection = function (pos, pr) { this.Sections.splice(pos, 0, pr); };
CPresentation.prototype.removeSection = function (pos) { this.Sections.splice(pos, 1); };
CPresentation.prototype.shiftSlides = function (pos, array) {
	// upstream's own two bail-outs, which leave Slides untouched
	if (!this.CanEdit() || this.mixed)
		return;
	var sorted = array.slice().sort(function (a, b) { return a - b; });
	var taken = [];
	var i;
	for (i = sorted.length - 1; i >= 0; i--)
		taken.unshift(this.Slides.splice(sorted[i], 1)[0]);
	var dest = pos;
	for (i = 0; i < sorted.length; i++)
		if (sorted[i] < pos)
			dest--;
	for (i = 0; i < taken.length; i++)
		this.Slides.splice(dest + i, 0, taken[i]);
};
global.AscCommonSlide.CPresentation = CPresentation;
ok(sections.install() !== false, "the prototype wrappers install");

// The view names the section a drop means; the model only asks.
var asked = [];
var refreshed = [];
var chosen = null;
sections.setDropResolver({
	ownerFor: function (p, pos) { asked.push(pos); return chosen; },
	afterShift: function (p, pos) { refreshed.push(pos); }
});

var wrapped = new CPresentation(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
chosen = wrapped.Sections[0];
wrapped.shiftSlides(3, [0]);
eq(shape(wrapped), "A@0:3 B@3:3", "a real move fixes the cuts");
eq(asked, [3], "the owner was asked for, once, with the insert index");
eq(refreshed, [3], "and the view was told to re-lay out");

// A bail-out leaves the slide count unchanged too, so length alone
// cannot tell it from a move -- and a wrongly "fixed" cut is a derived
// write, which undo cannot take back.
var locked = new CPresentation(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
locked.editable = false;
chosen = locked.Sections[1];
locked.shiftSlides(3, [0]);
eq(shape(locked), "A@0:3 B@3:3", "a refused move leaves the cuts alone");
eq(locked.Slides.map(function (s) { return s.i; }), [0, 1, 2, 3, 4, 5], "and the slides");

var mixed = new CPresentation(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
mixed.mixed = true;
chosen = null;
mixed.shiftSlides(3, [0]);
eq(shape(mixed), "A@0:3 B@3:3", "so does a mixed selection");

// and no edit rights means no deliberate edits either
var ro = new CPresentation(6, [{name: "A", at: 0}]);
ro.editable = false;
ok(sections.canAddSection(ro, 3) === false, "no new section without edit rights");
ok(sections.addSection(ro, 3) === null, "and the call is refused");
ok(sections.renameSection(ro, ro.Sections[0], "X") === false, "nor a rename");

console.log("PresentationSections: " + passed + " checks passed");
