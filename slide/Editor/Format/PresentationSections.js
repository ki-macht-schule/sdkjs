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
/*
 * Kiwi: PowerPoint p14:sectionLst as PPTY record 7.
 *
 * MODEL. A section is a cut in slide order: {guid, name, startIndex}.
 * Membership is never stored -- sectionOf(i) is the last cut with
 * start <= i. Every mutation ends in normalise(), which is what makes
 * the model total: cuts are sorted, clamped into [0, slideCount] and
 * the first cut is pinned at 0. Nothing is dropped: a section emptied
 * by dragging its slides away keeps its cut and can be refilled, which
 * is what PowerPoint stores too (an empty <p14:section>).
 *
 * An insert index is therefore ambiguous in general -- several cuts can
 * sit on it -- so a caller that knows which section it means passes it
 * (noteDropOwner); anything else derives "the section of the slide I
 * follow".
 *
 * TWO KINDS OF WRITE, and they must not be mixed:
 *
 *   derived  -- a slide was inserted, removed or moved. The slide change
 *               is already in History; we only re-derive the cuts.
 *               setStart() assigns the field DIRECTLY. Using upstream's
 *               CPrSection.setStartIndex() here would (a) add a second
 *               History entry that undo applies on top of our own
 *               syncFromChange, and (b) throw "Fehler bei der
 *               Bearbeitung" whenever no action is open (load, save,
 *               the finally of shiftSlides).
 *   deliberate -- the user added, renamed or moved a section. Wrapped in
 *               StartAction/FinalizeAction and written through the
 *               upstream setters, so it is undoable.
 *
 * Hooks (rebase on a version bump) -- see KIWI-CHANGES.md:
 *   configs/slide.json                 this file after Presentation.js
 *   Format.js CPres.fromStream case 7  PresentationSections.read(s)
 *   SerializeWriter WritePresentation  PresentationSections.write
 *   Serialize.js Load                  PresentationSections.apply
 *   CPresentation.insertSlide / removeSlideByObject / shiftSlides
 *   CChangesDrawingsContentPresentation.Load / Redo / Undo
 *     -- the last two are prototype wrappers installed from here, so
 *        they appear in no git diff of an upstream file.
 *
 * Tests: node slide/Editor/Format/PresentationSections.test.js
 */
(function (root) {
	"use strict";

	var SLD_ID_BASE = 256;
	var ATTR_START = 0xFA;
	var ATTR_END = 0xFB;

	// ---------------------------------------------------------------
	// PPTY record 7
	// ---------------------------------------------------------------

	function attrStart() {
		if (root.AscCommon && root.AscCommon.g_nodeAttributeStart != null)
			return root.AscCommon.g_nodeAttributeStart;
		return ATTR_START;
	}

	function attrEnd() {
		if (root.AscCommon && root.AscCommon.g_nodeAttributeEnd != null)
			return root.AscCommon.g_nodeAttributeEnd;
		return ATTR_END;
	}

	function readAttrs(s, handlers) {
		s.Skip2(1);
		for (;;) {
			var at = s.GetUChar();
			if (at === attrEnd())
				break;
			if (handlers[at])
				handlers[at]();
		}
	}

	function readXmlId(s) {
		var end = s.cur + s.GetULong() + 4;
		var id = null;
		readAttrs(s, {
			0: function () { id = s.GetString2(); },
			1: function () { s.GetString2(); }
		});
		s.Seek2(end);
		return id;
	}

	function readSection(s) {
		var end = s.cur + s.GetULong() + 4;
		var guid = null;
		var name = null;
		readAttrs(s, {
			0: function () { guid = s.GetString2(); },
			1: function () { name = s.GetString2(); }
		});
		var sldIds = [];
		while (s.cur < end) {
			var rec = s.GetUChar();
			if (rec === 0) {
				s.Skip2(4);
				var n = s.GetULong();
				for (var i = 0; i < n; i++) {
					s.Skip2(1);
					var id = readXmlId(s);
					if (id != null)
						sldIds.push(id);
				}
			} else {
				s.SkipRecord();
			}
		}
		s.Seek2(end);
		return {guid: guid, name: name, sldIds: sldIds};
	}

	function read(s) {
		var end = s.cur + s.GetULong() + 4;
		var pending = [];
		while (s.cur < end) {
			var rec = s.GetUChar();
			if (rec === 0) {
				s.Skip2(4);
				var n = s.GetULong();
				for (var i = 0; i < n; i++) {
					s.Skip2(1);
					pending.push(readSection(s));
				}
			} else {
				s.SkipRecord();
			}
		}
		s.Seek2(end);
		return pending;
	}

	// ---------------------------------------------------------------
	// the cut model
	// ---------------------------------------------------------------

	function startOf(section) {
		if (!section)
			return 0;
		return section.startIndex == null ? 0 : section.startIndex;
	}

	/** Derived write: no History. See the header comment. */
	function setStart(section, value) {
		if (section)
			section.startIndex = value;
	}

	function slideCountOf(presentation) {
		return (presentation && presentation.Slides && presentation.Slides.length) || 0;
	}

	/**
	 * Put the cut list back in order: sorted, clamped, first cut at 0.
	 *
	 * Nothing is dropped. A section whose slides have all been dragged
	 * away keeps its cut and simply owns nothing -- PowerPoint stores
	 * exactly that (a <p14:section> with an empty sldIdLst), and it is
	 * the only way a section the user emptied by accident can be filled
	 * again instead of having to be recreated by name.
	 *
	 * Ties keep their array order, so the sections at one index stay in
	 * the order the user put them in, and the LAST of them owns the
	 * slides that follow (sectionOf takes the last cut <= index).
	 */
	function normalise(presentation) {
		var sections = presentation && presentation.Sections;
		if (!sections || !sections.length)
			return;
		var count = slideCountOf(presentation);
		var i;
		var order = [];
		for (i = 0; i < sections.length; i++) {
			var start = startOf(sections[i]);
			if (start < 0)
				start = 0;
			if (start > count)
				start = count;
			setStart(sections[i], start);
			order.push({section: sections[i], at: i});
		}
		// Array.prototype.sort is not stable everywhere this runs, so
		// the original position is the tie-breaker.
		order.sort(function (a, b) {
			return startOf(a.section) - startOf(b.section) || a.at - b.at;
		});
		for (i = 0; i < order.length; i++)
			sections[i] = order[i].section;
		// the first section always owns slide 0
		setStart(sections[0], 0);
	}

	function nextStart(sections, index, count) {
		for (var i = index + 1; i < sections.length; i++) {
			if (sections[i])
				return startOf(sections[i]);
		}
		return count == null ? 0 : count;
	}

	function sectionOf(presentation, index) {
		var sections = presentation && presentation.Sections;
		if (!sections || !sections.length || index == null || index < 0)
			return null;
		var dest = null;
		for (var i = 0; i < sections.length; i++) {
			if (startOf(sections[i]) <= index)
				dest = sections[i];
			else
				break;
		}
		return dest;
	}

	function rangeOf(presentation, section) {
		var sections = (presentation && presentation.Sections) || [];
		var count = slideCountOf(presentation);
		var at = sections.indexOf(section);
		if (at === -1)
			return {start: 0, end: 0};
		var start = startOf(section);
		var end = nextStart(sections, at, count);
		if (end < start)
			end = start;
		return {start: start, end: Math.min(end, count)};
	}

	function membersOf(presentation, section) {
		var slides = (presentation && presentation.Slides) || [];
		var range = rangeOf(presentation, section);
		return slides.slice(range.start, range.end);
	}

	function isCut(presentation, pos) {
		var sections = (presentation && presentation.Sections) || [];
		for (var i = 0; i < sections.length; i++) {
			if (startOf(sections[i]) === pos)
				return true;
		}
		return false;
	}

	/** Every cut that sits exactly at *pos*, in array order. */
	function cutsAt(presentation, pos) {
		var sections = (presentation && presentation.Sections) || [];
		var out = [];
		for (var i = 0; i < sections.length; i++) {
			if (startOf(sections[i]) === pos)
				out.push(sections[i]);
		}
		return out;
	}

	/**
	 * Which section a slide inserted at *insertAt* would join.
	 *
	 * An insert index is ambiguous whenever a cut sits on it: it means
	 * both "last slide of the section above" and "first slide of any
	 * section starting here" -- and with empty sections there can be
	 * several of those. The caller therefore names the section it means
	 * (the rail's drop targets carry it); with no owner named, a new
	 * slide joins the section of the slide it follows, which is what
	 * duplicate, paste and insert all want.
	 */
	function ownerForInsert(presentation, insertAt, owner) {
		if (owner)
			return owner;
		if (insertAt <= 0)
			return (presentation.Sections || [])[0] || null;
		return sectionOf(presentation, insertAt - 1);
	}

	// ---------------------------------------------------------------
	// derived fixups
	// ---------------------------------------------------------------

	function busy(presentation) {
		return !presentation || presentation._sectionMove || presentation._sectionBatch
			|| !presentation.Sections || !presentation.Sections.length;
	}

	/**
	 * Push the cuts aside for *n* slides landing at *pos*.
	 *
	 * A cut strictly after pos always moves. A cut exactly ON pos moves
	 * only if it comes after the owner in the section order -- that is
	 * what puts the new slide inside the owner and not in the section
	 * that starts on the same index.
	 */
	function shiftForInsert(presentation, pos, n, owner) {
		var sections = presentation.Sections;
		var ownerAt = sections.indexOf(ownerForInsert(presentation, pos, owner));
		for (var i = 0; i < sections.length; i++) {
			var start = startOf(sections[i]);
			if (start > pos || (start === pos && i > ownerAt))
				setStart(sections[i], start + n);
		}
	}

	function onInsert(presentation, pos, count, owner) {
		if (busy(presentation))
			return;
		shiftForInsert(presentation, pos, count == null ? 1 : count, owner);
		normalise(presentation);
	}

	function onDelete(presentation, pos) {
		if (busy(presentation))
			return;
		var sections = presentation.Sections;
		for (var i = 0; i < sections.length; i++) {
			if (startOf(sections[i]) > pos)
				setStart(sections[i], startOf(sections[i]) - 1);
		}
		normalise(presentation);
	}

	/**
	 * One move: n slides leave *removed* and land at *insertAt*.
	 *
	 * Deletes first (descending), then one insert. A section that loses
	 * its last slide keeps its cut and simply owns nothing until
	 * something is dragged back into it.
	 */
	function shiftMove(presentation, removed, insertAt, count, owner) {
		if (!presentation || !presentation.Sections || !presentation.Sections.length)
			return;
		var sections = presentation.Sections;
		var sorted = (removed || []).slice().sort(function (a, b) { return b - a; });
		var i;
		var j;
		for (i = 0; i < sorted.length; i++) {
			for (j = 0; j < sections.length; j++) {
				if (startOf(sections[j]) > sorted[i])
					setStart(sections[j], startOf(sections[j]) - 1);
			}
		}
		shiftForInsert(presentation, insertAt, count, owner);
		normalise(presentation);
	}

	/** Upstream decrements pos once per removed index below it. */
	function adjustInsertPos(pos, array) {
		var list = (array || []).slice().sort(function (a, b) { return a - b; });
		var next = pos;
		for (var i = 0; i < list.length; i++) {
			if (list[i] < next)
				next--;
			else
				break;
		}
		return next;
	}

	// ---------------------------------------------------------------
	// load / save
	// ---------------------------------------------------------------

	function indexFromSldId(id) {
		var n = parseInt(id, 10);
		return isNaN(n) ? -1 : n - SLD_ID_BASE;
	}

	function pendingClaimed(pending) {
		var n = 0;
		for (var i = 0; i < (pending || []).length; i++)
			n += ((pending[i] && pending[i].sldIds) || []).length;
		return n;
	}

	/**
	 * Turn read() output into cuts. First id wins a slide; a section's
	 * cut is its lowest claimed index. Reports how many ids bound so
	 * the caller can refuse to write a file it did not understand.
	 */
	function bind(slides, pending) {
		var count = slides ? slides.length : 0;
		var out = [];
		if (!pending || !pending.length)
			return {sections: out, claimed: pendingClaimed(pending), mapped: 0};
		var taken = {};
		var mapped = 0;
		var i;
		for (i = 0; i < pending.length; i++) {
			var ids = pending[i].sldIds || [];
			var first = null;
			for (var j = 0; j < ids.length; j++) {
				var idx = indexFromSldId(ids[j]);
				if (idx >= 0 && idx < count && !taken[idx]) {
					taken[idx] = true;
					mapped++;
					if (first == null || idx < first)
						first = idx;
				}
			}
			out.push({name: pending[i].name, guid: pending[i].guid, startIndex: first});
		}
		// a section that bound nothing inherits the following cut
		var next = count;
		for (i = out.length - 1; i >= 0; i--) {
			if (out[i].startIndex == null)
				out[i].startIndex = next;
			else
				next = out[i].startIndex;
		}
		return {sections: out, claimed: pendingClaimed(pending), mapped: mapped};
	}

	function apply(presentation) {
		if (!presentation)
			return;
		// The module may have loaded before CPresentation was exported.
		// Load is the first moment a real document exists, so retry here.
		install();
		var pending = presentation.pendingSections;
		presentation.pendingSections = null;
		if (pending == null)
			return;
		var bound = bind(presentation.Slides, pending);
		presentation.sectionsUnbound = bound.claimed > 0 && bound.mapped === 0;
		presentation.sectionsPartial = bound.claimed > bound.mapped && bound.mapped > 0;
		var Ctor = root.AscCommonSlide && root.AscCommonSlide.CPrSection;
		var counter = root.AscCommon && root.AscCommon.g_oIdCounter;
		var wasLoad = counter && counter.m_bLoad;
		if (counter && counter.Set_Load)
			counter.Set_Load(false);
		var sections = [];
		for (var i = 0; i < bound.sections.length; i++) {
			var raw = bound.sections[i];
			var section = Ctor ? new Ctor() : {};
			section.name = raw.name;
			section.guid = raw.guid;
			section.startIndex = raw.startIndex;
			sections.push(section);
		}
		presentation.Sections = sections;
		if (!presentation.sectionsUnbound && !presentation.sectionsPartial)
			normalise(presentation);
		if (counter && counter.Set_Load)
			counter.Set_Load(wasLoad);
	}

	function idsForWrite(presentation, section) {
		var range = rangeOf(presentation, section);
		var ids = [];
		for (var i = range.start; i < range.end; i++)
			ids.push(String(SLD_ID_BASE + i));
		return ids;
	}

	function writeOne(writer, presentation, section) {
		writer.WriteUChar(attrStart());
		writer._WriteString2(0, section.guid);
		writer._WriteString2(1, section.name);
		writer.WriteUChar(attrEnd());
		writer.WriteRecordArray(0, 0, idsForWrite(presentation, section), function (sldId) {
			writer.WriteUChar(attrStart());
			writer._WriteString2(0, sldId);
			writer.WriteUChar(attrEnd());
		});
	}

	function write(writer, presentation) {
		warnIfUnhooked();
		var sections = presentation && presentation.Sections;
		if (!sections || !sections.length || !presentation.Slides)
			return;
		// A file we did not fully understand stays untouched rather than
		// losing the membership we could not read (PowerPoint sldIds).
		if (presentation.sectionsUnbound || presentation.sectionsPartial)
			return;
		if (!slideCountOf(presentation))
			return;
		writer.StartRecord(7);
		writer.WriteRecordArray(0, 0, sections, function (section) {
			writeOne(writer, presentation, section);
		});
		writer.EndRecord();
	}

	// ---------------------------------------------------------------
	// deliberate edits -- undoable
	// ---------------------------------------------------------------

	/**
	 * A deliberate, undoable change.
	 *
	 * The catch is not decoration. An exception that escapes a mouse
	 * handler reaches the editor's window.onerror, which shows "Bei der
	 * Arbeit mit dem Dokument ist ein Fehler aufgetreten", disconnects
	 * co-editing and drops the editor into view-only. Losing a section
	 * is recoverable; losing the session is not. The stack goes to the
	 * console with a prefix so it is one search away.
	 */
	function inAction(presentation, description, body) {
		var started = false;
		if (presentation && typeof presentation.StartAction === "function") {
			presentation.StartAction(description || 0);
			started = true;
		}
		try {
			return body();
		} catch (e) {
			if (typeof console !== "undefined" && console.error)
				console.error("kiwi sections: action failed", e);
			return null;
		} finally {
			if (started && typeof presentation.FinalizeAction === "function") {
				try {
					presentation.FinalizeAction();
				} catch (e2) {
					if (typeof console !== "undefined" && console.error)
						console.error("kiwi sections: finalize failed", e2);
				}
			}
		}
	}

	/** Upstream checks this before every mutation; so must we. */
	function canEdit(presentation) {
		if (!presentation)
			return false;
		return typeof presentation.CanEdit !== "function" || presentation.CanEdit();
	}

	function newGuid() {
		var hex = "0123456789ABCDEF";
		var out = [];
		for (var i = 0; i < 32; i++)
			out.push(hex.charAt((Math.random() * 16) | 0));
		return "{" + out.slice(0, 8).join("") + "-" + out.slice(8, 12).join("") + "-" +
			out.slice(12, 16).join("") + "-" + out.slice(16, 20).join("") + "-" +
			out.slice(20).join("") + "}";
	}

	function uniqueName(sections, base) {
		var taken = {};
		for (var i = 0; i < (sections || []).length; i++) {
			if (sections[i] && sections[i].name)
				taken[sections[i].name] = true;
		}
		if (!taken[base])
			return base;
		var n = 2;
		while (taken[base + " " + n])
			n++;
		return base + " " + n;
	}

	/** A cut may sit before any slide but the first, and only once. */
	function canAddSection(presentation, atIndex) {
		if (!presentation || !presentation.Sections || atIndex == null)
			return false;
		if (!canEdit(presentation))
			return false;
		if (atIndex <= 0 || atIndex >= slideCountOf(presentation))
			return false;
		return !isCut(presentation, atIndex);
	}

	/** Where a cut at atIndex belongs in the section order. */
	function insertPosFor(sections, atIndex) {
		for (var i = 0; i < sections.length; i++) {
			if (startOf(sections[i]) > atIndex)
				return i;
		}
		return sections.length;
	}

	function addSection(presentation, atIndex, name) {
		if (!canAddSection(presentation, atIndex))
			return null;
		return inAction(presentation, 0, function () {
			var sections = presentation.Sections;
			var section = new root.AscCommonSlide.CPrSection();
			section.setName(uniqueName(sections, name || "Neuer Abschnitt"));
			section.setGuid(newGuid());
			section.setStartIndex(atIndex);
			// Upstream's addSection is the only History-backed way into
			// the array. Its change type ships unregistered, which
			// install() fixes -- see registerSectionChanges().
			presentation.addSection(insertPosFor(sections, atIndex), section);
			normalise(presentation);
			return section;
		});
	}

	/**
	 * A section whose name is a kiwi unit id is that unit's; renaming it
	 * would break the title lookup and the deck's own identity. Anything
	 * else -- a section the user added here -- is free.
	 *
	 * The caller supplies the unit ids, so the model stays ignorant of
	 * what kiwi puts in a name.
	 */
	function canRenameSection(section, unitIds, presentation) {
		if (!section || !section.name)
			return false;
		if (presentation && !canEdit(presentation))
			return false;
		return !(unitIds && unitIds[section.name]);
	}

	function renameSection(presentation, section, name) {
		if (!section || !name || !canEdit(presentation))
			return false;
		var trimmed = String(name).replace(/^\s+|\s+$/g, "");
		if (!trimmed)
			return false;
		var sections = (presentation && presentation.Sections) || [];
		for (var i = 0; i < sections.length; i++) {
			if (sections[i] !== section && sections[i].name === trimmed)
				return false;
		}
		inAction(presentation, 0, function () {
			if (typeof section.setName === "function")
				section.setName(trimmed);
			else
				section.name = trimmed;
		});
		return true;
	}

	/**
	 * Move a whole section before the section now at destIndex.
	 *
	 * The slides travel with it; the cut list is rebuilt from the run
	 * lengths so no index has to be patched by hand.
	 */
	function moveSection(presentation, section, destIndex) {
		var sections = presentation && presentation.Sections;
		var slides = presentation && presentation.Slides;
		if (!sections || !slides || !section || !canEdit(presentation))
			return false;
		var from = sections.indexOf(section);
		if (from === -1)
			return false;
		var dest = destIndex;
		if (dest < 0)
			dest = 0;
		if (dest > sections.length)
			dest = sections.length;
		if (from === dest || from + 1 === dest)
			return false;
		var count = slides.length;
		var lengths = [];
		var i;
		for (i = 0; i < sections.length; i++)
			lengths.push(nextStart(sections, i, count) - startOf(sections[i]));
		var start = startOf(section);
		var run = lengths[from];
		var destStart = dest < sections.length ? startOf(sections[dest]) : count;
		var indexes = [];
		for (i = start; i < start + run; i++)
			indexes.push(i);
		return inAction(presentation, 0, function () {
			presentation._sectionMove = true;
			try {
				if (indexes.length && presentation.moveSlides)
					presentation.moveSlides(indexes, destStart);
				// Every step goes through History. The array move used to
				// be a raw splice, which undo could not see: the slides
				// went back and the names stayed put, so the names ended
				// up on the wrong runs.
				presentation.removeSection(from);
				var movedLen = lengths.splice(from, 1)[0];
				var at = dest > from ? dest - 1 : dest;
				if (at < 0)
					at = 0;
				if (at > sections.length)
					at = sections.length;
				presentation.addSection(at, section);
				lengths.splice(at, 0, movedLen);
				var cursor = 0;
				for (var n = 0; n < sections.length; n++) {
					sections[n].setStartIndex(cursor);
					cursor += lengths[n];
				}
			} finally {
				presentation._sectionMove = false;
			}
			return true;
		});
	}

	// ---------------------------------------------------------------
	// upstream wrappers -- no lines in Presentation.js
	// ---------------------------------------------------------------

	var historyHooked = false;
	var presentationHooked = false;
	var warned = false;

	function warnIfUnhooked() {
		if (warned || (historyHooked && presentationHooked))
			return;
		warned = true;
		if (root.document && typeof console !== "undefined" && console.warn)
			console.warn("kiwi sections: hooks missing (history=" + historyHooked +
				", presentation=" + presentationHooked + "); cuts will drift");
	}

	/**
	 * The rail says which section the drop means, just before the move.
	 *
	 * The index alone cannot say it: a cut sits on that index and the
	 * user picked one side of it (or, with empty sections, one of
	 * several bars). Cleared by the wrapper, so an insert that arrives
	 * from anywhere else -- duplicate, paste, undo -- derives its owner
	 * instead of inheriting a stale answer.
	 */
	/*
	 * How a drop names its section.
	 *
	 * The index alone cannot: a cut sits on it, and the user picked one
	 * side of it (or, with empty sections, one of several bars). Only
	 * the view knows which. So the view registers a resolver here --
	 * the model never learns what a thumbnail strip is, and there is no
	 * module-level "the last drop was..." to go stale.
	 */
	var dropResolver = null;

	function setDropResolver(resolver) {
		dropResolver = resolver || null;
	}

	function resolveDropOwner(presentation, pos) {
		if (dropResolver && dropResolver.ownerFor)
			return dropResolver.ownerFor(presentation, pos) || null;
		return null;
	}

	function afterDrop(presentation, pos, owner) {
		if (dropResolver && dropResolver.afterShift)
			dropResolver.afterShift(presentation, pos, owner);
	}

	/*
	 * Which section a slide will rejoin when it comes back.
	 *
	 * Deleting the FIRST slide of a section is the one insert whose
	 * owner cannot be derived: on undo the slide returns to an index a
	 * cut already sits on, and "the section of the slide I follow"
	 * picks the wrong side. So the answer is recorded while it is still
	 * knowable -- at the removal -- against the slide object, which is
	 * the same object History puts back.
	 */
	var reopenOwner = typeof WeakMap === "function" ? new WeakMap() : null;

	function noteRemoved(presentation, slide, at) {
		if (!reopenOwner || !slide)
			return;
		var here = cutsAt(presentation, at);
		if (here.length)
			reopenOwner.set(slide, here[here.length - 1]);
		else
			reopenOwner["delete"](slide);
	}

	function ownerForReopen(slide) {
		return (reopenOwner && slide && reopenOwner.get(slide)) || null;
	}

	function isSlideContentChange(change) {
		var presentation = change && change.Class;
		if (!presentation || !presentation.Slides || !presentation.Sections)
			return false;
		if (typeof change.private_GetChangedArray === "function")
			return change.private_GetChangedArray() === presentation.Slides;
		var DFH = root.AscDFH;
		if (DFH)
			return change.Type === DFH.historyitem_Presentation_AddSlide
				|| change.Type === DFH.historyitem_Presentation_RemoveSlide;
		return false;
	}

	function syncFromChange(change, forward) {
		if (!isSlideContentChange(change))
			return;
		var presentation = change.Class;
		if (presentation._sectionMove || presentation._sectionBatch)
			return;
		var items = change.Items || [];
		var isAdd = change.IsAdd ? change.IsAdd() : !!change.Add;
		var adding = (isAdd && forward) || (!isAdd && !forward);
		var i;
		if (adding) {
			for (i = 0; i < items.length; i++) {
				var at = presentation.Slides.indexOf(items[i]);
				if (at !== -1)
					onInsert(presentation, at, 1, ownerForReopen(items[i]));
			}
			return;
		}
		var pos = change.Pos != null ? change.Pos : null;
		if (pos == null && items[0])
			pos = presentation.Slides.indexOf(items[0]);
		if (pos == null || pos < 0)
			return;
		// one delete per slide, not one per change: a multi-slide delete
		// moves the cuts below it by as many places as it removed
		for (i = 0; i < Math.max(1, items.length); i++) {
			noteRemoved(presentation, items[i], pos);
			onDelete(presentation, pos);
		}
	}

	function hookHistoryChanges() {
		var DFH = root.AscDFH;
		if (!DFH || !DFH.CChangesDrawingsContentPresentation)
			return false;
		var proto = DFH.CChangesDrawingsContentPresentation.prototype;
		historyHooked = true;
		if (proto._kiwiSectionHook)
			return true;
		proto._kiwiSectionHook = true;
		var origLoad = proto.Load;
		var origRedo = proto.Redo;
		var origUndo = proto.Undo;
		proto.Load = function (Color) {
			origLoad.call(this, Color);
			syncFromChange(this, true);
		};
		proto.Redo = function () {
			origRedo.call(this);
			syncFromChange(this, true);
		};
		proto.Undo = function () {
			origUndo.call(this);
			syncFromChange(this, false);
		};
		return true;
	}

	function moved(before, after) {
		if (before.length !== after.length)
			return true;
		for (var i = 0; i < before.length; i++) {
			if (before[i] !== after[i])
				return true;
		}
		return false;
	}

	function hookPresentation() {
		var Slide = root.AscCommonSlide;
		if (!Slide || !Slide.CPresentation)
			return false;
		var proto = Slide.CPresentation.prototype;
		presentationHooked = true;
		if (proto._kiwiSectionHook)
			return true;
		proto._kiwiSectionHook = true;

		var origInsert = proto.insertSlide;
		if (origInsert) {
			proto.insertSlide = function (pos, slide) {
				var r = origInsert.apply(this, arguments);
				onInsert(this, pos, 1);
				return r;
			};
		}

		var origRemove = proto.removeSlideByObject;
		if (origRemove) {
			proto.removeSlideByObject = function (oSlide, bNoCheck, pos) {
				if (!oSlide)
					return origRemove.apply(this, arguments);
				var DFH = root.AscDFH;
				var type = typeof oSlide.getObjectType === "function" ? oSlide.getObjectType() : null;
				var isSlide = !DFH || type === DFH.historyitem_type_Slide;
				var at = typeof pos === "number"
					? pos
					: (this.Slides ? this.Slides.indexOf(oSlide) : -1);
				if (isSlide && at >= 0)
					noteRemoved(this, oSlide, at);
				var r = origRemove.apply(this, arguments);
				if (isSlide && at >= 0)
					onDelete(this, at);
				return r;
			};
		}

		var origShift = proto.shiftSlides;
		if (origShift) {
			proto.shiftSlides = function (pos, array, bCopy) {
				var owner = resolveDropOwner(this, pos);
				if (!this.Sections || !this.Sections.length || !array || !array.length)
					return origShift.call(this, pos, array, bCopy);
				if (this.IsMasterMode && this.IsMasterMode())
					return origShift.call(this, pos, array, bCopy);
				var before = this.Slides ? this.Slides.slice() : [];
				this._sectionBatch = true;
				var result;
				try {
					result = origShift.call(this, pos, array, bCopy);
				} finally {
					this._sectionBatch = false;
				}
				/*
				 * Upstream bails out on its own guards -- !CanEdit(),
				 * CheckIsMixedSelection() -- without touching Slides, and
				 * a bail-out leaves the LENGTH unchanged too. Comparing
				 * lengths therefore rewrote every cut below the drop
				 * point while the slides stayed exactly where they were,
				 * and being a derived write, undo could not take it back.
				 * Identity is the only honest test.
				 */
				var after = this.Slides || [];
				if (after.length === before.length && !moved(before, after))
					return result;
				var mouse = root.AscCommon && root.AscCommon.global_mouseEvent;
				var copying = (bCopy === true) || !!(mouse && mouse.CtrlKey);
				if (copying) {
					if (after.length > before.length)
						onInsert(this, pos, after.length - before.length, owner);
				} else if (after.length === before.length) {
					shiftMove(this, array, adjustInsertPos(pos, array), array.length, owner);
				}
				// The strip has to re-lay out: the cuts moved, so the bars
				// and the per-slide advances are stale. Doing it here
				// rather than in a DrawingDocument hook keeps the hook
				// count down and covers every caller of shiftSlides,
				// including DublicateSlide.
				afterDrop(this, pos, owner);
				return result;
			};
		}
		return true;
	}

	/**
	 * Upstream declares historyitem_Presentation_AddSection and writes
	 * CPresentation.addSection against it, but never registers the
	 * change: no changesFactory entry, no drawingContentChanges entry.
	 * The array it touches therefore comes back undefined on undo, redo
	 * and on another user's change. Two lines make the upstream method
	 * work as written -- cheaper and safer than a second way of adding
	 * a section.
	 */
	function registerSectionChanges() {
		var DFH = root.AscDFH;
		var Slide = root.AscCommonSlide;
		if (!DFH || !DFH.CChangesDrawingsContent)
			return false;

		/*
		 * History.Get_RecalcData does
		 *
		 *     Item.Class.Refresh_RecalcData(Item.Data)
		 *
		 * with no guard (word/Editor/History.js). CPrSection has no such
		 * method -- upstream declared the setters but nothing ever called
		 * them -- so the first deliberate section edit threw out of
		 * FinalizeAction and into window.onerror: "Bei der Arbeit mit dem
		 * Dokument ist ein Fehler aufgetreten", co-editing disconnected,
		 * editor read-only.
		 *
		 * A cut carries no geometry, so there is genuinely nothing to
		 * recalculate. The honest implementation is the empty one.
		 */
		if (Slide && Slide.CPrSection && !Slide.CPrSection.prototype.Refresh_RecalcData)
			Slide.CPrSection.prototype.Refresh_RecalcData = function () {};

		/*
		 * And CPresentation.removeSection is broken as written
		 * (Presentation.js:983): `this.Sections.splice(pos, 0)` removes
		 * nothing, and it records an ADD change for a removal. Replaced
		 * here rather than in Presentation.js, like the other wrappers.
		 */
		var CPresentation = Slide && Slide.CPresentation;
		if (CPresentation && !CPresentation.prototype._kiwiRemoveSection) {
			CPresentation.prototype._kiwiRemoveSection = true;
			CPresentation.prototype.removeSection = function (pos) {
				var gone = this.Sections[pos];
				if (!gone)
					return;
				root.History.Add(new DFH.CChangesDrawingsContent(
					this, DFH.historyitem_Presentation_AddSection, pos, [gone], false));
				this.Sections.splice(pos, 1);
			};
		}

		var type = DFH.historyitem_Presentation_AddSection;
		if (type == null)
			return false;
		if (!DFH.changesFactory[type])
			DFH.changesFactory[type] = DFH.CChangesDrawingsContent;
		if (!DFH.drawingContentChanges[type]) {
			DFH.drawingContentChanges[type] = function (oClass) {
				return oClass.Sections;
			};
		}
		return true;
	}

	function install() {
		registerSectionChanges();
		hookHistoryChanges();
		hookPresentation();
	}

	var api = {
		read: read,
		write: write,
		bind: bind,
		apply: apply,
		normalise: normalise,
		startOf: startOf,
		sectionOf: sectionOf,
		sectionOfInsert: ownerForInsert,
		membersOf: membersOf,
		cutsAt: cutsAt,
		onInsert: onInsert,
		onDelete: onDelete,
		shiftMove: shiftMove,
		adjustInsertPos: adjustInsertPos,
		addSection: addSection,
		canAddSection: canAddSection,
		renameSection: renameSection,
		canRenameSection: canRenameSection,
		moveSection: moveSection,
		setDropResolver: setDropResolver,
		syncFromChange: syncFromChange,
		install: install
	};

	if (typeof module !== "undefined" && module.exports)
		module.exports = api;
	root.AscCommonSlide = root.AscCommonSlide || {};
	root.AscCommonSlide.PresentationSections = api;
	install();
})(typeof window !== "undefined" ? window : globalThis);
