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
 * Kiwi: remount stock Office Theme / Blank slides onto the first
 * course master. Not a section concern and not a serializer concern.
 * WriteDocument2 must not call this - a writer does not change the
 * document. The caller is BinaryPPTYLoader.Load in common/Shapes/
 * Serialize.js, next to PresentationSections.apply: both are repairs
 * on the document that was just read.
 *
 * Tests: node slide/Editor/Format/PresentationThemeRemount.test.js
 */
(function (root) {
	"use strict";

	function themeName(master) {
		return (master && master.Theme && master.Theme.name) || "";
	}

	function isStockTheme(name) {
		return name === "Office Theme" || name === "Blank";
	}

	function firstCourseMaster(presentation) {
		var masters = presentation && presentation.slideMasters;
		if (!masters)
			return null;
		for (var i = 0; i < masters.length; i++) {
			if (!isStockTheme(themeName(masters[i])))
				return masters[i];
		}
		return null;
	}

	function remountOfficeHijacks(presentation) {
		var kim = firstCourseMaster(presentation);
		var slides = presentation && presentation.Slides;
		if (!kim || !slides)
			return 0;
		var n = 0;
		for (var i = 0; i < slides.length; i++) {
			var slide = slides[i];
			var master = slide.Layout && slide.Layout.Master;
			if (!master || !isStockTheme(themeName(master)))
				continue;
			var layout = slide.Layout;
			var next = kim.getMatchingLayout
				? kim.getMatchingLayout(
					layout.type,
					layout.matchingName,
					layout.cSld && layout.cSld.name,
					true
				)
				: (kim.sldLayoutLst && kim.sldLayoutLst[0]);
			if (!next || next === layout)
				continue;
			if (slide.setLayout)
				slide.setLayout(next);
			else
				slide.Layout = next;
			n++;
		}
		return n;
	}

	var api = {
		remountOfficeHijacks: remountOfficeHijacks,
		isStockTheme: isStockTheme
	};

	if (typeof module !== "undefined" && module.exports)
		module.exports = api;
	root.AscCommonSlide = root.AscCommonSlide || {};
	root.AscCommonSlide.PresentationThemeRemount = api;
})(typeof window !== "undefined" ? window : globalThis);
