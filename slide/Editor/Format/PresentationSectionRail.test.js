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
var rail = require("./PresentationSectionRail.js");

function CPrSection() {
	this.name = null;
	this.guid = null;
	this.startIndex = null;
}
CPrSection.prototype.setName = function (v) { this.name = v; };
CPrSection.prototype.setGuid = function (v) { this.guid = v; };
CPrSection.prototype.setStartIndex = function (v) { this.startIndex = v; };

global.AscCommonSlide = {
	PresentationSections: sections,
	PresentationSectionRail: rail,
	CPrSection: CPrSection
};

var passed = 0;
function ok(cond, what) {
	assert.ok(cond, what);
	passed++;
}
function eq(a, b, what) {
	assert.deepStrictEqual(a, b, what);
	passed++;
}

var THUMB_H = 72;
var THUMB_W = 128;
var BORDER = 7;        // const_border_w at ratio 1
var STEP = THUMB_H + 3 * BORDER;
var RAIL_W = 180;
var OFFSET_X = 24;
var START = 8;

function deck(n, cuts) {
	var slides = [];
	for (var i = 0; i < n; i++)
		slides.push({i: i});
	return {
		Slides: slides,
		Sections: cuts.map(function (c) {
			var s = new CPrSection();
			s.name = c.name;
			s.guid = "{" + c.name + "}";
			s.startIndex = c.at;
			return s;
		}),
		collapsedSections: {},
		addSection: function (pos, pr) { this.Sections.splice(pos, 0, pr); },
		removeSection: function (pos) { this.Sections.splice(pos, 1); }
	};
}

function names(list) {
	return list.map(function (g) { return g.owner ? g.owner.name : "-"; });
}

function metrics(n) {
	return {
		slideCount: n,
		vertical: true,
		rtl: false,
		startOffset: START,
		currentScroll: 0,
		thumbW: THUMB_W,
		thumbH: THUMB_H,
		offsetX: OFFSET_X,
		offsetY: 8,
		offsetR: 8,
		border: BORDER,
		ratio: 1,
		slideStep: STEP,
		railW: RAIL_W
	};
}

/** What CalculatePlaces does with the plan, including hidden slides. */
function place(plan, n) {
	var pages = [];
	var cursor = START;
	for (var i = 0; i < n; i++) {
		var slot = plan.slots[i];
		cursor += slot.advance;
		var page = {pageIndex: i, left: OFFSET_X, right: OFFSET_X + THUMB_W};
		if (slot.hidden) {
			// the fix: keep x, collapse the height. No -10000 sentinel,
			// so every upstream reader of m_arrPages[0] stays sane.
			page.top = cursor;
			page.bottom = cursor;
			page.sectionHidden = true;
		} else {
			page.top = cursor;
			page.bottom = cursor + THUMB_H;
			page.sectionHidden = false;
			cursor += STEP;
		}
		pages.push(page);
	}
	return pages;
}

// --- layout ---------------------------------------------------------

var p = deck(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
var plan = rail.layout(p, metrics(6));
eq(plan.headers.map(function (h) { return h.section.name; }), ["A", "B"], "one bar per section");
eq(plan.slots.map(function (s) { return s.advance > 0 ? 1 : 0; }), [1, 0, 0, 1, 0, 0],
	"only the first slide of a section carries the bar");
ok(plan.slots.every(function (s) { return !s.hidden; }), "nothing hidden");
eq(plan.extra, 2 * rail.headerAdvance(1), "the bars are the extra strip length");
// The strip has to grow by exactly the space the bars take, or the
// scrollbar and the thumbnails disagree about how long the strip is.
eq(plan.extra, plan.slots.reduce(function (n, s) { return n + s.advance; }, 0),
	"the extra length is exactly the advances the bars asked for");
eq(rail.extraLength(p, 6, STEP, 1), plan.extra, "and CheckSizes is told the same number");

var pages = place(plan, 6);
ok(pages[0].left === OFFSET_X, "thumbs keep their column");
ok(pages[0].top > plan.headers[0].bottom, "the first thumb sits under its bar");
ok(pages[3].top > plan.headers[1].bottom, "and so does B's first");

// --- drop targets ---------------------------------------------------

eq(plan.gaps.length, 8, "seven insert positions, and the seam counts twice");
var seam = plan.gaps.filter(function (g) { return g.insertAt === 3; });
eq(seam.length, 2, "a cut offers two targets");
eq(names(seam), ["A", "B"], "above the bar stays in A, below it joins B");
ok(seam[0].along < plan.headers[1].top, "the 'stay in A' line is above the bar");
ok(seam[1].along > plan.headers[1].bottom, "the 'join B' line is below it");

function lands(y) {
	var gap = rail.dropTarget(plan.gaps, 40, y, true);
	var dest = sections.sectionOfInsert(p, gap.insertAt, gap.owner);
	return {at: gap.insertAt, name: dest ? dest.name : "-", line: Math.round(gap.along)};
}

var endOfA = pages[2].bottom;
var barTop = plan.headers[1].top;
var barBottom = plan.headers[1].bottom;
eq(lands(endOfA - 4).name, "A", "just above the end of A stays in A");
eq(lands(endOfA + 2).name, "A", "in the gap under A's last slide too");
eq(lands(barTop + 2).name, "A", "on the upper half of B's bar still A");
eq(lands(barBottom - 2).name, "B", "past the middle of the bar it is B");
eq(lands(barBottom + 4).name, "B", "and below the bar of course");
eq(lands(endOfA + 2).at, 3, "the index is the seam either way");
eq(lands(barBottom + 4).at, 3, "the index is the seam either way");

// one line per outcome: no jumping between equivalent positions
var seen = {};
for (var y = endOfA - 20; y < barBottom + 20; y++) {
	var got = lands(y);
	var key = got.at + ":" + got.name;
	if (!seen[key])
		seen[key] = {};
	seen[key][got.line] = true;
}
eq(Object.keys(seen["3:A"]).length, 1, "'stays in A' has exactly one line position");
eq(Object.keys(seen["3:B"]).length, 1, "'joins B' has exactly one line position");

// --- a section with a single slide -----------------------------------

var solo = deck(3, [{name: "A", at: 0}, {name: "B", at: 1}, {name: "C", at: 2}]);
var soloPlan = rail.layout(solo, metrics(3));
eq(soloPlan.headers.length, 3, "three bars for three one-slide sections");
eq(soloPlan.slots.map(function (s) { return s.advance > 0 ? 1 : 0; }), [1, 1, 1],
	"every slide carries a bar");
var soloGaps = soloPlan.gaps.filter(function (g) { return g.insertAt === 1; });
eq(soloGaps.length, 2, "the seam before a one-slide section still offers both sides");
var soloPages = place(soloPlan, 3);
ok(soloPages[1].top > soloPlan.headers[1].bottom, "the lone slide sits under its own bar");

// --- collapsed --------------------------------------------------------

rail.setCollapsed(p, p.Sections[0], true);
var collapsed = rail.layout(p, metrics(6));
eq(collapsed.slots.map(function (s) { return s.hidden ? 1 : 0; }), [1, 1, 1, 0, 0, 0],
	"A's slides are hidden");
eq(collapsed.headers.length, 2, "both bars stay");
eq(collapsed.extra, 2 * rail.headerAdvance(1) - 3 * STEP, "the strip gets shorter");
var collapsedPages = place(collapsed, 6);
ok(collapsedPages[0].left === OFFSET_X, "a hidden page keeps its x, so page 0 stays usable");
ok(collapsedPages[0].top === collapsedPages[0].bottom, "it just has no height");
ok(collapsedPages[3].top > collapsed.headers[1].bottom, "B still sits under its bar");
var intoCollapsed = rail.dropTarget(collapsed.gaps, 40, collapsed.headers[0].bottom + 1, true);
eq(sections.sectionOfInsert(p, intoCollapsed.insertAt, intoCollapsed.owner).name, "A",
	"a collapsed bar is still a drop target");
rail.setCollapsed(p, p.Sections[0], false);

// --- the "new section" chip ------------------------------------------

plan = rail.layout(p, metrics(6));
pages = place(plan, 6);
ok(rail.chipFor(p, pages, 0, 1, plan) === null, "no cut before the first slide");
ok(rail.chipFor(p, pages, 3, 1, plan) === null, "and none where a cut already is");
var chip = rail.chipFor(p, pages, 4, 1, plan);
ok(chip !== null, "but one in the middle of a section");
// The bars and the chip are drawn on the canvas UNDER the thumbnails,
// so anything over a slide is invisible. The chip lives in the gap.
ok(chip.bottom <= pages[4].top, "the chip sits above the slide, not on it");
ok(chip.top >= pages[3].bottom, "and below the one before -- inside the gap");
var chipMid = (chip.left + chip.right) / 2;
eq(chipMid, (pages[4].left + pages[4].right) / 2, "centred over the slide");
eq(chip.lineFrom, pages[4].left, "the cut line spans the thumbnail column");
eq(chip.lineTo, pages[4].right, "exactly");
plan.headers.forEach(function (h) {
	ok(!(chip.top < h.bottom && chip.bottom > h.top), "and never inside a bar");
});
ok(rail.hitChip(chip, chip.left + 2, chip.top + 2) === true, "the chip is hittable");
ok(rail.hitChip(chip, chip.left + 400, chip.top) === false, "next to it is not");
eq(rail.chipAt(p, pages, pages[4].left + 40, pages[4].top + 30, 1, plan).insertAt, 4,
	"hovering a slide offers its chip");
eq(rail.chipAt(p, pages, chipMid, chip.top + 1, 1, plan).insertAt, 4,
	"and the chip's own spot keeps it alive -- otherwise it is unclickable");
eq(rail.chipAt(p, pages, chipMid, pages[3].bottom + 1, 1, plan).insertAt, 4,
	"the gap belongs to the slide below it");
eq(rail.chipAt(p, pages, chipMid, pages[4].bottom + 1, 1, plan).insertAt, 5,
	"so the next gap already offers the next slide's chip");
ok(rail.chipAt(p, pages, pages[3].left + 40, pages[3].top + 30, 1, plan) === null,
	"a slide that already starts a section offers none");

// --- header hit areas -------------------------------------------------

var header = plan.headers[1];
ok(rail.hitHeader(plan.headers, 40, header.top + 2) === header, "the bar is hittable");
ok(rail.hitHeader(plan.headers, 40, header.bottom + 8) === null, "the gap below is not");
ok(rail.hitChevron(header, header.left + 10, header.top + 5, 1) === true, "chevron on the left");
ok(rail.hitTitle(header, header.left + 60, header.top + 5, 1) === true, "title next to it");
ok(rail.hitChevron(header, header.left + 60, header.top + 5, 1) === false, "and they do not overlap");

// --- an emptied section keeps its bar ---------------------------------

// A dragged its only slide away: its cut now sits on B's, and B (the
// later cut) owns the slides. The bar must stay, and it must be a drop
// target, or the section can never be filled again.
var emptied = deck(4, [{name: "A", at: 0}, {name: "B", at: 0}]);
var emptyPlan = rail.layout(emptied, metrics(4));
eq(emptyPlan.headers.map(function (h) { return h.section.name; }), ["A", "B"],
	"both bars are drawn, in section order");
eq(emptyPlan.headers.map(function (h) { return h.count; }), [0, 4], "A holds nothing");
eq(emptyPlan.slots[0].advance, 2 * rail.headerAdvance(1),
	"the first slide makes room for both bars");
var intoEmpty = emptyPlan.gaps.filter(function (g) { return g.insertAt === 0; });
eq(names(intoEmpty), ["A", "B"], "and each bar offers its own drop target");
var emptyPages = place(emptyPlan, 4);
ok(emptyPages[0].top > emptyPlan.headers[1].bottom, "slide 0 sits under the lower bar");

// a trailing empty section: the last slide was dragged out of C
var trailing = deck(3, [{name: "A", at: 0}, {name: "C", at: 3}]);
var trailPlan = rail.layout(trailing, metrics(3));
eq(trailPlan.headers.length, 2, "the bar past the last slide is still drawn");
eq(trailPlan.headers[1].count, 0, "and says it is empty");
eq(names(trailPlan.gaps.filter(function (g) { return g.insertAt === 3; })), ["A", "C"],
	"both ends of the deck are reachable");

// --- who may be renamed ------------------------------------------------

ok(rail.canRename({name: "Neuer Abschnitt"}) === true, "a user section may be renamed");
ok(rail.canRename({name: null}) === false, "a nameless one may not");

// --- headers carry their slide count ---------------------------------

eq(plan.headers.map(function (h) { return h.count; }), [3, 3], "three slides each");
var uneven = rail.layout(deck(5, [{name: "A", at: 0}, {name: "B", at: 1}]), metrics(5));
eq(uneven.headers.map(function (h) { return h.count; }), [1, 4], "and a lone slide counts one");

// --- skin -------------------------------------------------------------

var skin = rail.colors({
	BackgroundColorThumbnails: "#F4F4F4",
	ThumbnailsPageNumberText: "#000000",
	ThumbnailsPageOutlineActive: "#848484"
});
ok(skin.fill !== "#F4F4F4", "the bar is not the background");
eq(skin.text, "#000000", "text follows the skin");
var dark = rail.colors({
	BackgroundColorThumbnails: "#404040",
	ThumbnailsPageNumberText: "#FFFFFF",
	ThumbnailsPageOutlineActive: "#A0A0A0"
});
ok(dark.fill !== skin.fill, "dark mode differs");
var odd = rail.colors({BackgroundColorThumbnails: "rgb(1,2,3)", ThumbnailsPageNumberText: "red"});
ok(odd.fill !== "rgb(1,2,3)", "a non-hex skin still yields a visible bar");
ok(odd.text.charAt(0) === "#", "and readable text");

// --- section drag -----------------------------------------------------

var track = rail.beginTrack(p.Sections[0], 0, 0);
eq(rail.moveTrack(track, 2, 2).dragging, false, "a small move is still a click");
eq(rail.moveTrack(track, 40, 0).dragging, true, "past the threshold it is a drag");
var barB = plan.headers[1];
var ended = rail.endTrack(track, 40, barB.top + 2, plan.headers, p, true);
eq(ended.action, "move", "dropping on the upper half of another bar reorders");
eq(ended.destIndex, 1, "before that section");
// the lower half means "after it" -- without that, dropping A on B
// (the obvious "one further down") is a silent no-op
eq(rail.endTrack(track, 40, barB.bottom - 2, plan.headers, p, true).destIndex, 2,
	"the lower half puts it after");
eq(rail.endTrack(rail.beginTrack(p.Sections[0], 0, 0), 0, 0, plan.headers, p, true).action, "none",
	"a click without movement does not reorder");
// A slip used to move the whole section to the end of the deck.
eq(rail.endTrack(track, 40, pages[4].top + 10, plan.headers, p, true).action, "none",
	"releasing over a slide does nothing");
eq(rail.endTrack(track, 40, plan.headers[0].top + 2, plan.headers, p, true).action, "none",
	"and releasing on its own bar does nothing");

// --- no rail for a deck we could not bind ----------------------------

ok(rail.railable(p) === true, "an ordinary deck gets a rail");
p.sectionsUnbound = true;
ok(rail.railable(p) === false, "an unbound one does not -- its edits would be lost on save");
delete p.sectionsUnbound;
p.sectionsPartial = true;
ok(rail.railable(p) === false, "nor a half-bound one");
delete p.sectionsPartial;

// --- the host adapter ------------------------------------------------
//
// This is the seam between DrawingDocument.js and the rail, and it
// needs no canvas and no DOM: the host is a plain object with the four
// fields the adapter actually reads.

function host(presentation, n) {
	var h = {
		m_oWordControl: {m_oLogicDocument: presentation},
		IsMasterMode: function () { return false; },
		refreshed: 0,
		KiwiRefresh: function () { this.refreshed++; }
	};
	h.sectionPlan = rail.hostLayout(h, metrics(n));
	h.m_arrPages = h.sectionPlan ? place(h.sectionPlan, n) : [];
	return h;
}

var hp = deck(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
hp.CanEdit = function () { return true; };
var h = host(hp, 6);
ok(h.sectionPlan !== null, "a usable deck gets a plan");

// what a click means, by where it lands
var barA = h.sectionPlan.headers[0];
eq(rail.hostMouseDown(h, barA.left + 10, barA.top + 5).action, "collapsed",
	"the chevron collapses");
rail.setCollapsed(hp, hp.Sections[0], false);
eq(rail.hostMouseDown(h, barA.left + 60, barA.top + 5).action, "rename",
	"the title opens the name field");
eq(rail.hostMouseDown(h, barA.left + 60, barA.bottom - 1).action, "rename",
	"still the title");
hp.CanEdit = function () { return false; };
eq(rail.hostMouseDown(h, barA.left + 60, barA.top + 5).action, "track",
	"but without edit rights it is only a drag handle");
hp.CanEdit = function () { return true; };
h.HeaderTrack = null;
var hchip = rail.chipFor(hp, h.m_arrPages, 4, 1, h.sectionPlan);
eq(rail.hostMouseDown(h, (hchip.left + hchip.right) / 2, (hchip.top + hchip.bottom) / 2).action,
	"added", "the chip adds a section");
ok(hp.Sections.length === 3, "and it is really there");
h = host(hp, 6);
ok(rail.hostMouseDown(h, 40, h.m_arrPages[5].top + 20) === null,
	"a click on a slide is not ours");

// hover is derived from the pointer, never stored as rectangles --
// otherwise it drifts along when the strip scrolls underneath it
hp = deck(6, [{name: "A", at: 0}, {name: "B", at: 3}]);
hp.CanEdit = function () { return true; };
h = host(hp, 6);
var over = rail.hostMouseMove(h, 40, h.m_arrPages[4].top + 20, 1);
eq(over.repaint, true, "moving onto a slide with a chip asks for a repaint");
eq(rail.hostMouseMove(h, 40, h.m_arrPages[4].top + 21, 1).repaint, false,
	"moving within it does not");
var seenChip = h.sectionPointer;
ok(seenChip.x === 40, "only the pointer is remembered");
rail.hostClearHover(h);
ok(h.sectionPointer === null, "and leaving forgets it");

// the drop target, as ConvertCoords2 asks for it
var gapAt = rail.hostDropIndex(h, 40, h.sectionPlan.headers[1].bottom + 4, true);
eq(gapAt.insertAt, 3, "below B's bar is the seam");
eq(gapAt.owner.name, "B", "and it means B");
eq(h.sectionGap, gapAt, "the answer stays on the host for the drop to read");

// a collapsed destination opens up when something is dropped into it
rail.setCollapsed(hp, hp.Sections[0], true);
ok(rail.hostRevealDrop(h, 0, hp.Sections[0]) === true, "dropping into a collapsed section opens it");
ok(rail.hostRevealDrop(h, 0, hp.Sections[0]) === false, "an open one needs nothing");

// a deck we could not bind gets no plan at all
var bad = deck(4, [{name: "A", at: 0}]);
bad.sectionsUnbound = true;
var badHost = host(bad, 4);
ok(badHost.sectionPlan === null, "no plan for an unbound deck");
ok(rail.hostMouseDown(badHost, 40, 40) === null, "so nothing in the strip is ours");

console.log("PresentationSectionRail: " + passed + " checks passed");
