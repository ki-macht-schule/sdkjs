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
 * Kiwi: the section rail in the PE thumbnail strip.
 *
 * VIEW. One pass -- layout() -- produces everything geometric:
 * the header bars, the per-slide advance, and the drop targets. Nothing
 * else in this file or in DrawingDocument.js re-derives a position from
 * mouse coordinates, so the line the user sees and the index the model
 * gets are the same number by construction.
 *
 *   headers[]  bar rectangles, title already resolved
 *   slots[i]   {advance, hidden} for slide i
 *   gaps[]     drop targets {insertAt, owner, along, lo, hi}
 *
 * A collapsed section hides its slides but keeps its bar, and its bar
 * stays a drop target (the host expands it on drop).
 *
 * "Neuer Abschnitt" is a chip in the gap between two thumbnails, centred
 * over the slide that would start the section, with a tick across the
 * thumbnail column showing where the cut would land. It cannot live ON
 * the slide: the rail draws on m_oThumbnailsBack, which sits UNDER the
 * thumbnail canvas, so anything over a slide is hidden by its image.
 *
 * Tests: node slide/Editor/Format/PresentationSectionRail.test.js
 */
(function (root) {
	"use strict";

	var GAP_ABOVE = 12;
	var BAR_H = 20;
	// the bar must not crowd the slide it heads
	var GAP_BELOW = 12;
	var INSET = 4;
	var RADIUS = 5;
	var DRAG = 10;
	var CHIP = 16;

	function sectionsApi() {
		return root.AscCommonSlide && root.AscCommonSlide.PresentationSections;
	}

	function startOf(section) {
		var api = sectionsApi();
		return api ? api.startOf(section) : (section && section.startIndex) || 0;
	}

	function px(value, ratio) {
		return Math.round(value * (ratio || 1));
	}

	function headerAdvance(ratio) {
		return px(GAP_ABOVE + BAR_H + GAP_BELOW, ratio);
	}

	// ---------------------------------------------------------------
	// titles and collapse state
	// ---------------------------------------------------------------

	/**
	 * The one channel from the host that reaches sdkjs.
	 *
	 * DocsAPI has two places a host can put its own data, and they do
	 * NOT both arrive here: `editorConfig.customization` is consumed by
	 * the web-apps front end and never reaches the SDK -- asc_CDocInfo
	 * has no Customization field at all. `document.options` does: the
	 * API copies it into DocInfo.Options / documentOpenOptions.
	 *
	 * kiwi sends the map as `document.options.kiwiUnitTitles`.
	 */
	function openOptions() {
		var api = (root.Asc && root.Asc.editor) || root.editor;
		if (!api)
			return null;
		if (api.DocInfo && typeof api.DocInfo.asc_getOptions === "function")
			return api.DocInfo.asc_getOptions();
		return api.documentOpenOptions || null;
	}

	var titlesWarned = false;

	function unitTitles() {
		var options = openOptions();
		var map = options && options.kiwiUnitTitles;
		if (!map && !titlesWarned && root.document) {
			titlesWarned = true;
			if (typeof console !== "undefined" && console.info)
				console.info("kiwi sections: no document.options.kiwiUnitTitles;"
					+ " bars fall back to the section name");
		}
		return map || {};
	}

	function unitIds() {
		var options = openOptions();
		var ids = options && options.kiwiUnitIds;
		return (ids && Object.prototype.toString.call(ids) === "[object Array]")
			? ids
			: null;
	}

	function kiwiProtocol() {
		var options = openOptions();
		return options ? options.kiwiProtocol : undefined;
	}

	/**
	 * Names the host marked reserved. Missing or not a list is ``null``:
	 * the model then treats every name as reserved.
	 */
	function reservedNames() {
		var options = openOptions();
		if (!options || !Object.prototype.hasOwnProperty.call(options, "kiwiReserved"))
			return null;
		var names = options.kiwiReserved;
		if (!names || Object.prototype.toString.call(names) !== "[object Array]")
			return null;
		return names;
	}

	/**
	 * A section whose name is a unit id belongs to that unit: renaming
	 * it would break the title lookup and the deck's identity. A section
	 * the user added here has a free-text name and can be renamed -- and
	 * still can after a reload, which a "freshly created" flag could not
	 * have managed. No id list, or the wrong protocol: nothing is
	 * renameable.
	 */
	function canRename(section, presentation) {
		var api = sectionsApi();
		return !!api && api.canRenameSection(
			section, unitIds(), presentation, kiwiProtocol(), reservedNames()
		);
	}

	/** A reserved bar is not a drag handle. View-only still is. */
	function canDrag(section) {
		var api = sectionsApi();
		return !!api && !api.isReservedName(section && section.name, reservedNames());
	}

	function titleOf(section, titles) {
		if (!section)
			return "";
		var id = section.name || "";
		var map = titles || unitTitles();
		return (id && map[id]) || id;
	}

	function collapseKey(section) {
		return (section && (section.guid || section.name)) || "";
	}

	function collapsedSet(presentation) {
		if (!presentation.collapsedSections)
			presentation.collapsedSections = {};
		return presentation.collapsedSections;
	}

	function isCollapsed(presentation, section) {
		var key = collapseKey(section);
		return !!(presentation && key && collapsedSet(presentation)[key]);
	}

	function setCollapsed(presentation, section, value) {
		var key = collapseKey(section);
		if (!presentation || !key)
			return;
		collapsedSet(presentation)[key] = !!value;
	}

	function toggleCollapse(presentation, section) {
		setCollapsed(presentation, section, !isCollapsed(presentation, section));
	}

	// ---------------------------------------------------------------
	// colours, from the editor skin
	// ---------------------------------------------------------------

	function parseHex(value) {
		if (typeof value !== "string" || value.charAt(0) !== "#")
			return null;
		var hex = value.slice(1);
		if (hex.length === 3)
			hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
		if (hex.length !== 6)
			return null;
		var n = parseInt(hex, 16);
		return isNaN(n) ? null : {r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255};
	}

	function toHex(c) {
		return "#" + ((1 << 24) + (c.r << 16) + (c.g << 8) + c.b).toString(16).slice(1);
	}

	function mix(a, b, t, fallback) {
		var A = parseHex(a);
		var B = parseHex(b);
		if (!A || !B)
			return fallback;
		return toHex({
			r: Math.round(A.r + (B.r - A.r) * t),
			g: Math.round(A.g + (B.g - A.g) * t),
			b: Math.round(A.b + (B.b - A.b) * t)
		});
	}

	// kiwi's own accent (schule.kiwi --primary). The bar is tinted with
	// it rather than painted in it: the strip belongs to the editor, and
	// a section header is a label, not a call to action.
	var KIWI = "#F29C55";

	function colors(skin) {
		var s = skin || (root.AscCommon && root.AscCommon.GlobalSkin) || {};
		var bg = s.BackgroundColorThumbnails || "#F4F4F4";
		var text = s.ThumbnailsPageNumberText || "#000000";
		var accent = s.ThumbnailsPageOutlineActive || "#848484";
		// A skin that is not plain hex (a host may register its own) must
		// not collapse the bar into the background -- hence the fallbacks.
		return {
			fill: mix(bg, KIWI, 0.12, "#F7E6D6"),
			hover: mix(bg, KIWI, 0.24, "#F2D3B5"),
			line: mix(bg, KIWI, 0.38, "#EBBF95"),
			kiwi: KIWI,
			text: parseHex(text) ? text : "#202020",
			accent: parseHex(accent) ? accent : "#848484",
			muted: mix(text, bg, 0.35, "#5A5A5A")
		};
	}

	// ---------------------------------------------------------------
	// layout: bars, slot advances and drop targets in one pass
	// ---------------------------------------------------------------

	function barRect(cursor, metrics, ratio) {
		var scroll = metrics.currentScroll || 0;
		var top = cursor + px(GAP_ABOVE, ratio) - scroll;
		var bottom = top + px(BAR_H, ratio);
		if (metrics.vertical) {
			return {
				left: 0,
				right: metrics.railW || 0,
				top: top,
				bottom: bottom
			};
		}
		return {
			left: top,
			right: bottom,
			top: 0,
			bottom: metrics.railW || 0
		};
	}

	function along(rect, vertical, which) {
		if (vertical)
			return which === "start" ? rect.top : rect.bottom;
		return which === "start" ? rect.left : rect.right;
	}

	/**
	 * Walk the slides once. Returns bars, per-slide advances and the
	 * complete list of drop targets.
	 */
	function layout(presentation, metrics) {
		var empty = {headers: [], slots: [], gaps: [], extra: 0};
		var api = sectionsApi();
		if (!presentation || !metrics || !api)
			return empty;
		var sections = presentation.Sections;
		if (!sections || !sections.length)
			return empty;

		var ratio = metrics.ratio || 1;
		var vertical = !!metrics.vertical;
		var count = metrics.slideCount || 0;
		var step = metrics.slideStep || 0;
		var advance = headerAdvance(ratio);
		var titles = unitTitles();
		var scroll = metrics.currentScroll || 0;

		var headers = [];
		var slots = [];
		var gaps = [];
		var extra = 0;
		var cursor = metrics.startOffset || 0;
		var prevEnd = null;   // end edge of the last laid-out thumbnail
		var i;

		// A drop target names BOTH the index and the section that will
		// own the slide -- the index alone is ambiguous wherever a cut
		// sits, and with empty sections several cuts can share one.
		function addGap(insertAt, owner, lo, hi) {
			if (lo > hi) {
				var swap = lo;
				lo = hi;
				hi = swap;
			}
			gaps.push({
				insertAt: insertAt,
				owner: owner || null,
				along: (lo + hi) / 2
			});
		}

		// Cuts per index -- a list, not a single entry: a section the
		// user emptied keeps its cut, so several bars can share an index.
		var barAt = {};
		for (i = 0; i < sections.length; i++) {
			var at = startOf(sections[i]);
			if (!barAt[at])
				barAt[at] = [];
			barAt[at].push(sections[i]);
		}

		// i runs one past the last slide so trailing bars get drawn too
		for (i = 0; i <= count; i++) {
			var here = barAt[i] || [];
			var k;
			for (k = 0; k < here.length; k++) {
				var section = here[k];
				var rect = barRect(cursor, metrics, ratio);
				rect.section = section;
				rect.startIndex = i;
				rect.title = titleOf(section, titles);
				rect.collapsed = isCollapsed(presentation, section);
				rect.inset = px(INSET, ratio);
				headers.push(rect);
				// Above the first bar at this index: the drop stays in
				// the section that owns the slide before it. Later bars
				// need no "above" target -- the previous bar's "below"
				// already names the same section at the same index.
				if (k === 0 && i > 0 && prevEnd != null)
					addGap(i, api.sectionOf(presentation, i - 1), prevEnd, along(rect, vertical, "start"));
				// below the bar: the drop joins this section
				addGap(i, section, along(rect, vertical, "end"),
					along(rect, vertical, "end") + px(GAP_BELOW, ratio));
				cursor += advance;
				extra += advance;
				prevEnd = along(rect, vertical, "end");
			}

			if (i === count)
				break;

			if (!here.length && i > 0 && prevEnd != null)
				addGap(i, api.sectionOf(presentation, i - 1), prevEnd, cursor - scroll);

			var holder = api.sectionOf(presentation, i);
			var hidden = holder ? isCollapsed(presentation, holder) : false;
			var thumbStart = cursor - scroll;
			slots[i] = {advance: here.length * advance, hidden: hidden};
			if (hidden) {
				extra -= step;
			} else {
				cursor += step;
				prevEnd = thumbStart + (vertical ? metrics.thumbH : metrics.thumbW);
			}
		}

		// after the last slide, if no trailing bar already covered it
		if (count > 0 && prevEnd != null && !(barAt[count] || []).length)
			addGap(count, api.sectionOf(presentation, count - 1), prevEnd, prevEnd + px(GAP_BELOW, ratio));

		for (i = 0; i < headers.length; i++) {
			var to = i + 1 < headers.length ? headers[i + 1].startIndex : count;
			headers[i].count = to - headers[i].startIndex;
		}

		return {
			headers: headers,
			slots: slots,
			gaps: gaps,
			extra: extra,
			// chipFor needs the strip's own geometry: the chip lives in
			// the gap between two thumbnails, not on one.
			vertical: vertical,
			railW: metrics.railW || 0,
			slideGap: Math.max(0, step - (vertical ? metrics.thumbH : metrics.thumbW) || 0)
		};
	}

	/** Extra strip length the bars add, for the scrollbar. */
	function extraLength(presentation, slideCount, slideStep, ratio) {
		var plan = layout(presentation, {
			slideCount: slideCount,
			slideStep: slideStep,
			ratio: ratio,
			vertical: true,
			startOffset: 0,
			currentScroll: 0,
			thumbW: 0,
			thumbH: slideStep,
			railW: 0
		});
		return plan.extra;
	}

	// ---------------------------------------------------------------
	// hit testing -- everything against the one layout
	// ---------------------------------------------------------------

	/** Nearest drop target. The only place a position becomes an index. */
	function dropTarget(gaps, x, y, vertical) {
		var pos = vertical ? y : x;
		var best = null;
		for (var i = 0; i < (gaps || []).length; i++) {
			var d = Math.abs(pos - gaps[i].along);
			if (!best || d < best.d)
				best = {d: d, gap: gaps[i]};
		}
		return best ? best.gap : null;
	}

	function inRect(rect, x, y) {
		return !!rect && x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom;
	}

	function hitHeader(headers, x, y) {
		for (var i = (headers || []).length - 1; i >= 0; i--) {
			if (inRect(headers[i], x, y))
				return headers[i];
		}
		return null;
	}

	function headerParts(header, ratio) {
		var inset = header.inset != null ? header.inset : px(INSET, ratio);
		var left = header.left + inset;
		var right = header.right - inset;
		var chevronRight = left + px(22, ratio);
		return {
			chevron: {left: left, top: header.top, right: chevronRight, bottom: header.bottom},
			title: {left: chevronRight, top: header.top, right: right, bottom: header.bottom}
		};
	}

	function hitChevron(header, x, y, ratio) {
		return !!header && inRect(headerParts(header, ratio).chevron, x, y);
	}

	function hitTitle(header, x, y, ratio) {
		return !!header && inRect(headerParts(header, ratio).title, x, y);
	}

	/**
	 * The chip marks the cut it would make, so it sits IN the gap
	 * between two thumbnails, centred on the slide below it.
	 *
	 * It cannot sit on the slide: the rail draws on m_oThumbnailsBack,
	 * which is under the thumbnail canvas, so the slide's own image
	 * would cover it. The gap is 3 * const_border_w and empty.
	 */
	function chipFor(presentation, pages, index, ratio, plan) {
		var api = sectionsApi();
		if (!api || !plan || !api.canAddSection(presentation, index))
			return null;
		var page = pages && pages[index];
		if (!page || page.sectionHidden)
			return null;
		var room = plan.slideGap || 0;
		var size = Math.min(px(CHIP, ratio), room - px(2, ratio));
		if (size < px(8, ratio))
			return null;
		var half = size / 2;
		var edge = (plan.vertical ? page.top : page.left) - room / 2;
		var mid = plan.vertical
			? (page.left + page.right) / 2
			: (page.top + page.bottom) / 2;
		var box = plan.vertical
			? {left: mid - half, right: mid + half, top: edge - half, bottom: edge + half}
			: {left: edge - half, right: edge + half, top: mid - half, bottom: mid + half};
		box.insertAt = index;
		// the cut line runs the width of the thumbnail column
		box.lineFrom = plan.vertical ? page.left : page.top;
		box.lineTo = plan.vertical ? page.right : page.bottom;
		box.at = edge;
		return box;
	}

	function hitChip(chip, x, y) {
		if (!chip)
			return false;
		// a couple of pixels of slack: a 16 px target is small enough
		var slack = (chip.right - chip.left) * 0.2;
		return x >= chip.left - slack && x <= chip.right + slack
			&& y >= chip.top - slack && y <= chip.bottom + slack;
	}

	/**
	 * Which slide's chip the pointer is offering.
	 *
	 * The zone is the slide's row across the strip PLUS the gap above
	 * it, because that gap is where the chip is drawn: a zone that
	 * stopped at the slide's edge would hide the chip the moment the
	 * pointer moved onto it, which makes it unclickable. The gap belongs
	 * to the slide below it, so the zones stay adjacent and disjoint.
	 */
	function chipAt(presentation, pages, x, y, ratio, plan) {
		if (!plan)
			return null;
		var pos = plan.vertical ? y : x;
		var cross = plan.vertical ? x : y;
		if (cross < 0 || cross > plan.railW)
			return null;
		var reach = plan.slideGap || 0;
		for (var i = 0; i < (pages || []).length; i++) {
			var page = pages[i];
			if (!page || page.sectionHidden)
				continue;
			var lo = (plan.vertical ? page.top : page.left) - reach;
			var hi = plan.vertical ? page.bottom : page.right;
			if (pos < lo || pos > hi)
				continue;
			return chipFor(presentation, pages, i, ratio, plan);
		}
		return null;
	}

	// ---------------------------------------------------------------
	// drawing
	// ---------------------------------------------------------------

	function roundRect(context, x, y, w, h, r) {
		var rad = Math.min(r, w / 2, h / 2);
		context.beginPath();
		context.moveTo(x + rad, y);
		context.lineTo(x + w - rad, y);
		context.quadraticCurveTo(x + w, y, x + w, y + rad);
		context.lineTo(x + w, y + h - rad);
		context.quadraticCurveTo(x + w, y + h, x + w - rad, y + h);
		context.lineTo(x + rad, y + h);
		context.quadraticCurveTo(x, y + h, x, y + h - rad);
		context.lineTo(x, y + rad);
		context.quadraticCurveTo(x, y, x + rad, y);
		context.closePath();
	}

	function drawChevron(context, cx, cy, collapsed, color, ratio) {
		var s = px(3.5, ratio);
		context.beginPath();
		if (collapsed) {
			context.moveTo(cx - s * 0.4, cy - s);
			context.lineTo(cx + s * 0.7, cy);
			context.lineTo(cx - s * 0.4, cy + s);
		} else {
			context.moveTo(cx - s, cy - s * 0.3);
			context.lineTo(cx, cy + s * 0.6);
			context.lineTo(cx + s, cy - s * 0.3);
		}
		context.closePath();
		context.fillStyle = color;
		context.fill();
	}

	function ellipsize(context, label, maxWidth) {
		if (maxWidth <= 8 || context.measureText(label).width <= maxWidth)
			return label;
		var text = label;
		while (text.length && context.measureText(text + "…").width > maxWidth)
			text = text.slice(0, -1);
		return text + "…";
	}

	function draw(context, plan, presentation, skin, state) {
		if (!context || !plan)
			return;
		var ratio = (root.AscCommon && root.AscCommon.AscBrowser
			&& root.AscCommon.AscBrowser.retinaPixelRatio) || 1;
		var col = colors(skin);
		var hoverSection = state && state.hoverSection;
		var chip = state && state.chip;
		var chipHot = state && state.chipHot;
		var headers = plan.headers || [];
		var i;

		for (i = 0; i < headers.length; i++) {
			var h = headers[i];
			var inset = h.inset != null ? h.inset : px(INSET, ratio);
			var x = h.left + inset;
			var y = h.top;
			var w = h.right - h.left - 2 * inset;
			var ht = h.bottom - h.top;
			if (w < 2 || ht < 2)
				continue;
			var hot = hoverSection === h.section;
			context.save();
			context.textBaseline = "middle";
			var mid = y + ht / 2;

			// the pill
			roundRect(context, x, y, w, ht, px(RADIUS, ratio));
			context.fillStyle = hot ? col.hover : col.fill;
			context.fill();
			context.strokeStyle = col.line;
			context.lineWidth = Math.max(1, px(1, ratio));
			context.stroke();

			// a kiwi tick on the leading edge: brand, not decoration
			context.save();
			roundRect(context, x, y, px(3, ratio), ht, px(1.5, ratio));
			context.fillStyle = col.kiwi;
			context.fill();
			context.restore();

			// count on the right, so the title never has to compete.
			// An emptied section says so -- otherwise its bar looks like
			// a bug rather than a place to drop something.
			var label = h.count > 0 ? String(h.count) : "leer";
			var tail = px(10, ratio);
			if (label) {
				context.font = (h.count > 0 ? "" : "italic ") + px(11, ratio)
					+ "px \"Segoe UI\", Arial, sans-serif";
				context.fillStyle = col.muted;
				context.textAlign = "right";
				context.fillText(label, x + w - px(10, ratio), mid);
				tail = px(14, ratio) + context.measureText(label).width;
			}

			drawChevron(context, x + px(13, ratio), mid, h.collapsed, col.muted, ratio);

			context.font = "600 " + px(12, ratio) + "px \"Segoe UI\", Arial, sans-serif";
			context.fillStyle = col.text;
			context.textAlign = "left";
			var textX = x + px(24, ratio);
			context.fillText(ellipsize(context, h.title || "", x + w - tail - textX), textX, mid);
			context.restore();
		}

		if (chip) {
			var cx = (chip.left + chip.right) / 2;
			var cy = (chip.top + chip.bottom) / 2;
			var rad = (chip.right - chip.left) / 2;
			var arm = rad * 0.42;
			context.save();
			// the cut this button would make, drawn where it would land
			context.strokeStyle = chipHot ? col.kiwi : col.line;
			context.lineWidth = Math.max(1, px(chipHot ? 2 : 1, ratio));
			context.beginPath();
			if (plan.vertical) {
				context.moveTo(chip.lineFrom, chip.at);
				context.lineTo(chip.lineTo, chip.at);
			} else {
				context.moveTo(chip.at, chip.lineFrom);
				context.lineTo(chip.at, chip.lineTo);
			}
			context.stroke();
			// a round button in the margin, so it reads as "click me"
			// rather than as part of the slide next to it
			context.beginPath();
			context.arc(cx, cy, rad, 0, 2 * Math.PI);
			context.fillStyle = chipHot ? col.kiwi : col.fill;
			context.fill();
			context.strokeStyle = chipHot ? col.kiwi : col.line;
			context.lineWidth = Math.max(1, px(1, ratio));
			context.stroke();
			context.strokeStyle = chipHot ? "#FFFFFF" : col.muted;
			context.lineWidth = Math.max(1, px(1.5, ratio));
			context.lineCap = "round";
			context.beginPath();
			context.moveTo(cx - arm, cy);
			context.lineTo(cx + arm, cy);
			context.moveTo(cx, cy - arm);
			context.lineTo(cx, cy + arm);
			context.stroke();
			context.restore();
		}
	}

	// ---------------------------------------------------------------
	// header drag (reordering whole sections)
	// ---------------------------------------------------------------

	function beginTrack(section, x, y) {
		return {section: section, X: x, Y: y, simple: true};
	}

	function grabbingCursor() {
		if (root.AscCommon && root.AscCommon.AscBrowser && root.AscCommon.AscBrowser.isWebkit)
			return "-webkit-grabbing";
		return "move";
	}

	function moveTrack(track, x, y) {
		if (!track)
			return {cursor: "default", dragging: false};
		if (track.simple && (Math.abs(track.X - x) > DRAG || Math.abs(track.Y - y) > DRAG))
			track.simple = false;
		return {cursor: track.simple ? "pointer" : grabbingCursor(), dragging: !track.simple};
	}

	/**
	 * A section drag ends only over a bar.
	 *
	 * Releasing anywhere else used to mean "move it to the end", so a
	 * slip during a drag relocated a whole section and its slides. A
	 * miss is a miss. On a bar, the half decides the side, which is what
	 * makes "one further down" -- dropping on the next bar -- possible
	 * at all.
	 */
	function endTrack(track, x, y, headers, presentation, vertical) {
		if (!track)
			return null;
		if (track.simple)
			return {action: "none", section: track.section};
		var sections = (presentation && presentation.Sections) || [];
		var dest = hitHeader(headers, x, y);
		if (!dest || !dest.section)
			return {action: "none", section: track.section};
		var at = sections.indexOf(dest.section);
		if (at < 0 || dest.section === track.section)
			return {action: "none", section: track.section};
		var down = vertical === false
			? x > (dest.left + dest.right) / 2
			: y > (dest.top + dest.bottom) / 2;
		var destIndex = down ? at + 1 : at;
		var api = sectionsApi();
		if (api && !api.canMoveSection(track.section, destIndex, presentation, reservedNames()))
			return {action: "none", section: track.section};
		return {
			action: "move",
			section: track.section,
			destIndex: destIndex
		};
	}

	// ---------------------------------------------------------------
	// inline rename -- any section the rail allows to rename
	// ---------------------------------------------------------------

	function endRename(host) {
		var input = host && host._kiwiRenameInput;
		if (!input)
			return;
		host._kiwiRenameInput = null;
		if (input.parentNode)
			input.parentNode.removeChild(input);
	}

	/**
	 * An input positioned over the title. Enter commits, Escape and
	 * blur cancel: the value is a unit id, so an accidental click
	 * elsewhere must not rewrite it.
	 */
	function startRename(host, canvas, header, presentation, ratio, onDone) {
		var api = sectionsApi();
		if (!host || !canvas || !canvas.parentNode || !header || !api)
			return false;
		if (!canRename(header.section, presentation))
			return false;
		endRename(host);
		var parent = canvas.parentNode;
		var rect = headerParts(header, ratio).title;
		var skin = colors(root.AscCommon && root.AscCommon.GlobalSkin);
		var computed = root.getComputedStyle ? root.getComputedStyle(parent) : null;
		if (computed && computed.position === "static")
			parent.style.position = "relative";
		var canvasBox = canvas.getBoundingClientRect();
		var parentBox = parent.getBoundingClientRect();
		var scaleX = canvas.width ? canvasBox.width / canvas.width : 1;
		var scaleY = canvas.height ? canvasBox.height / canvas.height : 1;
		var input = root.document.createElement("input");
		input.type = "text";
		input.value = header.section.name || "";
		input.setAttribute("aria-label", "Abschnittsname");
		// text_input.js reads this on document focus: "true" means the
		// element takes the keyboard, so the editor hands it over
		// (asc_enableKeyEvents(false)) instead of swallowing keystrokes.
		input.setAttribute("oo_editor_input", "true");
		input.style.cssText = [
			"position:absolute",
			"left:" + (canvasBox.left - parentBox.left + rect.left * scaleX) + "px",
			"top:" + (canvasBox.top - parentBox.top + rect.top * scaleY) + "px",
			"width:" + Math.max(60, (rect.right - rect.left) * scaleX) + "px",
			"height:" + Math.max(18, (rect.bottom - rect.top) * scaleY) + "px",
			"border:0",
			"padding:0 4px",
			"font:12px \"Segoe UI\",Arial,sans-serif",
			"color:" + skin.text,
			"background:" + skin.fill,
			"outline:1px solid " + skin.accent,
			"box-sizing:border-box",
			"z-index:5"
		].join(";");
		var editor = (root.Asc && root.Asc.editor) || root.editor;
		function keyEvents(enabled) {
			if (editor && typeof editor.asc_enableKeyEvents === "function")
				editor.asc_enableKeyEvents(enabled);
		}

		var done = false;
		function finish(commit) {
			if (done)
				return;
			done = true;
			var value = input.value;
			root.document.removeEventListener("mousedown", outside, true);
			root.removeEventListener("wheel", away, true);
			keyEvents(true);
			endRename(host);
			if (commit)
				api.renameSection(presentation, header.section, value, reservedNames());
			if (onDone)
				onDone(commit);
		}

		/*
		 * Focus, the hard way.
		 *
		 * The editor takes focus back to its own hidden input whenever a
		 * mouse event reaches the strip -- including the mouseup that
		 * ends the very click that opened this field, which is why
		 * typing only worked while the button was held. blur is
		 * therefore NOT a signal that the user is done with the field;
		 * it is noise. So: never commit on blur, take the focus straight
		 * back, and end the rename on the things that really mean it --
		 * Enter, Escape, a click somewhere else, or the strip scrolling
		 * out from under the field.
		 */
		function take() {
			if (done || host._kiwiRenameInput !== input)
				return;
			input.focus();
			input.select();
			keyEvents(false);
		}
		input.addEventListener("blur", function () {
			root.setTimeout(take, 0);
		});

		function swallow(e) {
			if (e.stopPropagation)
				e.stopPropagation();
		}
		["mousedown", "mouseup", "mousemove", "click", "dblclick",
			"keypress", "keyup", "touchstart", "pointerdown"].forEach(function (name) {
			input.addEventListener(name, swallow, true);
		});
		input.addEventListener("keydown", function (e) {
			if (e.key === "Enter" || e.keyCode === 13) {
				e.preventDefault();
				finish(true);
			} else if (e.key === "Escape" || e.keyCode === 27) {
				e.preventDefault();
				finish(false);
			}
			swallow(e);
		});

		// A click anywhere but the field commits -- that is what blur was
		// supposed to mean.
		function outside(e) {
			if (e.target !== input)
				finish(true);
		}
		// The field is positioned absolutely against the bar; once the
		// strip scrolls, that position is a lie.
		function away() {
			finish(true);
		}

		parent.appendChild(input);
		host._kiwiRenameInput = input;
		// registered a tick later so the mousedown that opened the field
		// does not immediately close it again
		root.setTimeout(function () {
			if (done)
				return;
			root.document.addEventListener("mousedown", outside, true);
			root.addEventListener("wheel", away, true);
			take();
		}, 0);
		return true;
	}

	// ---------------------------------------------------------------
	// host adapter
	//
	// The ONLY part that knows CThumbnailsManager. DrawingDocument.js
	// calls these six and nothing else, so the patch in that file stays
	// a handful of one-line hooks.
	// ---------------------------------------------------------------

	function useRatio(ratio) {
		return (typeof ratio === "number" && isFinite(ratio) && ratio > 0) ? ratio : 1;
	}

	function planOf(host) {
		return (host && host.sectionPlan) || {headers: [], slots: [], gaps: [], extra: 0};
	}

	function presentationOf(host) {
		return host && host.m_oWordControl && host.m_oWordControl.m_oLogicDocument;
	}

	/**
	 * Whether this deck gets a rail at all.
	 *
	 * An unbound or half-bound deck is the important case: bind() then
	 * has no idea which slides a section holds, write() refuses to emit
	 * record 7, and a rail would offer edits that are silently dropped
	 * on save. Better no rail than a lying one.
	 */
	function railable(presentation, host) {
		return !!(presentation
			&& presentation.Sections && presentation.Sections.length
			&& !presentation.sectionsUnbound && !presentation.sectionsPartial
			&& !(host && host.IsMasterMode && host.IsMasterMode())
			&& !(presentation.IsVisioEditor && presentation.IsVisioEditor()));
	}

	/** Called from CalculatePlaces. Stores the plan on the host. */
	function hostLayout(host, metrics) {
		var presentation = presentationOf(host);
		host.sectionPlan = railable(presentation, host) ? layout(presentation, metrics) : null;
		return host.sectionPlan;
	}

	/** Called from OnUpdateOverlay. */
	function hostDraw(host, context, skin) {
		if (!host.sectionPlan)
			return;
		// derived here, not remembered: see hoverAt()
		var hover = hoverAt(host);
		draw(context, host.sectionPlan, presentationOf(host), skin, {
			hoverSection: hover.section,
			chip: hover.chip,
			chipHot: hover.chipHot
		});
	}

	/** Called from ConvertCoords2. Null lets upstream decide. */
	function hostDropIndex(host, x, y, vertical) {
		var plan = host && host.sectionPlan;
		if (!plan || !plan.gaps.length)
			return null;
		var gap = dropTarget(plan.gaps, x, y, vertical);
		host.sectionGap = gap;
		return gap;
	}

	/**
	 * Mouse down in the strip. Returns what the host must do, or null
	 * when the event is not ours.
	 */
	function hostMouseDown(host, x, y, ratio) {
		var api2 = sectionsApi();
		var plan = host && host.sectionPlan;
		if (!plan || !api2)
			return null;
		var r = useRatio(ratio);
		var presentation = presentationOf(host);
		// resolve the chip here, not from the hover state: a click can
		// arrive without a preceding move (touch, or a fast click).
		var chip = chipAt(presentation, host.m_arrPages, x, y, r, plan);
		if (chip && hitChip(chip, x, y)) {
			var made = api2.addSection(presentation, chip.insertAt);
			return made ? {action: "added", section: made} : null;
		}
		var header = hitHeader(plan.headers, x, y);
		if (!header)
			return null;
		if (hitChevron(header, x, y, r)) {
			toggleCollapse(presentation, header.section);
			return {action: "collapsed", section: header.section};
		}
		if (hitTitle(header, x, y, r) && canRename(header.section, presentation))
			return {action: "rename", header: header};
		if (!canDrag(header.section))
			return {action: "locked"};
		host.HeaderTrack = beginTrack(header.section, x, y);
		return {action: "track"};
	}

	/**
	 * What the pointer is over, derived on demand from the CURRENT
	 * layout.
	 *
	 * Only the pointer is remembered, never the rectangles it hit. A
	 * stored chip rectangle keeps its absolute coordinates while the
	 * strip scrolls underneath it, so the chip appears to drift along in
	 * the background -- the hover has to be re-derived instead.
	 */
	function hoverAt(host, ratio) {
		var plan = host && host.sectionPlan;
		var at = host && host.sectionPointer;
		if (!plan || !at)
			return {section: null, chip: null, chipHot: false};
		var header = hitHeader(plan.headers, at.x, at.y);
		var chip = header
			? null
			: chipAt(presentationOf(host), host.m_arrPages, at.x, at.y, useRatio(ratio), plan);
		return {
			section: header ? header.section : null,
			header: header,
			chip: chip,
			chipHot: chip ? hitChip(chip, at.x, at.y) : false
		};
	}

	/**
	 * Mouse move. Returns {cursor, tooltip, repaint}; the caller only
	 * repaints when asked, so hovering does not redraw every frame.
	 */
	function hostMouseMove(host, x, y, ratio) {
		var plan = host && host.sectionPlan;
		if (!plan)
			return {cursor: null, tooltip: "", repaint: false};
		var r = useRatio(ratio);
		if (host.HeaderTrack) {
			host.sectionPointer = {x: x, y: y};
			var state = moveTrack(host.HeaderTrack, x, y);
			return {cursor: state.cursor, tooltip: "", repaint: false};
		}
		var before = hoverAt(host, r);
		host.sectionPointer = {x: x, y: y};
		var now = hoverAt(host, r);
		var changed = before.section !== now.section
			|| (before.chip && before.chip.insertAt) !== (now.chip && now.chip.insertAt)
			|| before.chipHot !== now.chipHot;
		var tooltip = "";
		if (now.chipHot)
			tooltip = "Neuen Abschnitt ab dieser Folie";
		else if (now.header && hitChevron(now.header, x, y, r))
			tooltip = now.header.collapsed ? "Abschnitt ausklappen" : "Abschnitt einklappen";
		else if (now.header && canRename(now.header.section, presentationOf(host)))
			tooltip = "Abschnitt benennen";
		var draggable = now.header && canDrag(now.header.section);
		var onChevron = now.header && hitChevron(now.header, x, y, r);
		return {
			cursor: (draggable || now.chipHot || onChevron) ? "pointer" : null,
			tooltip: tooltip,
			repaint: changed
		};
	}

	/** Mouse up. Only meaningful while a header is being dragged. */
	function hostMouseUp(host, x, y) {
		var track = host && host.HeaderTrack;
		if (!track)
			return null;
		host.HeaderTrack = null;
		var plan = planOf(host);
		var presentation = presentationOf(host);
		var ended = endTrack(track, x, y, plan.headers, presentation, plan.vertical);
		if (!ended || ended.action !== "move")
			return null;
		var api2 = sectionsApi();
		if (api2 && api2.moveSection(presentation, ended.section, ended.destIndex, reservedNames()))
			return {action: "moved"};
		return null;
	}

	/** A drop landed in a collapsed section: show it. */
	function hostRevealDrop(host, insertAt, owner) {
		var api2 = sectionsApi();
		var presentation = presentationOf(host);
		if (!api2 || !presentation)
			return false;
		var dest = api2.sectionOfInsert(presentation, insertAt, owner);
		if (!dest || !isCollapsed(presentation, dest))
			return false;
		setCollapsed(presentation, dest, false);
		return true;
	}

	/** The bar that currently heads *section*, after a relayout. */
	function hostHeaderFor(host, section) {
		var headers = planOf(host).headers;
		for (var i = 0; i < headers.length; i++) {
			if (headers[i].section === section)
				return headers[i];
		}
		return null;
	}

	/*
	 * The drop resolver, handed to the model at load.
	 *
	 * This is the only channel from the view back into the model, and
	 * it points the right way: the rail knows what a thumbnail strip is,
	 * the model does not. It also covers every caller of shiftSlides --
	 * a drag, but also DublicateSlide -- which a hook in one mouse
	 * handler never could.
	 */
	function thumbnailsOf(presentation) {
		var api = presentation && presentation.Api;
		var control = api && api.WordControl;
		return (control && control.Thumbnails) || null;
	}

	function installResolver() {
		var api = sectionsApi();
		if (!api || !api.setDropResolver)
			return false;
		api.setDropResolver({
			ownerFor: function (presentation, pos) {
				var host = thumbnailsOf(presentation);
				var gap = host && host.sectionGap;
				return (gap && gap.insertAt === pos) ? gap.owner : null;
			},
			afterShift: function (presentation, pos, owner) {
				var host = thumbnailsOf(presentation);
				if (!host || !host.KiwiRefresh)
					return;
				hostRevealDrop(host, pos, owner);
				host.KiwiRefresh();
			}
		});
		return true;
	}

	function hostClearHover(host) {
		host.sectionPointer = null;
		host.sectionGap = null;
		host.HeaderTrack = null;
	}

	var api = {
		headerAdvance: headerAdvance,
		setCollapsed: setCollapsed,
		colors: colors,
		canRename: canRename,
		canDrag: canDrag,
		layout: layout,
		extraLength: extraLength,
		dropTarget: dropTarget,
		hitHeader: hitHeader,
		hitChevron: hitChevron,
		hitTitle: hitTitle,
		chipFor: chipFor,
		chipAt: chipAt,
		hitChip: hitChip,
		beginTrack: beginTrack,
		moveTrack: moveTrack,
		endTrack: endTrack,
		startRename: startRename,
		railable: railable,
		hostLayout: hostLayout,
		hostDraw: hostDraw,
		hostDropIndex: hostDropIndex,
		hostMouseDown: hostMouseDown,
		hostMouseMove: hostMouseMove,
		hostMouseUp: hostMouseUp,
		installResolver: installResolver,
		hostHeaderFor: hostHeaderFor,
		hostRevealDrop: hostRevealDrop,
		hostClearHover: hostClearHover
	};

	if (typeof module !== "undefined" && module.exports)
		module.exports = api;
	root.AscCommonSlide = root.AscCommonSlide || {};
	root.AscCommonSlide.PresentationSectionRail = api;
	installResolver();
})(typeof window !== "undefined" ? window : globalThis);
