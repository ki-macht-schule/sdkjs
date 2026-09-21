#!/usr/bin/env python3
#
# Copyright (C) KI macht Schule gGmbH, 2026
#
# New file in a modified version of ONLYOFFICE Docs. Licensed under the
# GNU Affero General Public License version 3, together with the
# additional terms in the LICENSE file of the ONLYOFFICE Document
# Server distribution: this file is combined with that work and is
# distributed under its terms.
#
# Distributed WITHOUT ANY WARRANTY. See the GNU AGPL for details:
# https://www.gnu.org/licenses/agpl-3.0.html
#
# The delta against upstream, and where to get the complete source, is
# described in KIWI-CHANGES.md next to this file.
#
"""Apply the kiwi section rail hooks to a pristine DrawingDocument.js.

Every hook is a one-line delegation into PresentationSectionRail.js.
Each anchor must appear exactly once, or the script refuses -- that is
the rebase check: after an upstream bump, run it and it tells you which
anchor moved.
"""
import sys

HOOKS = []


def hook(name, old, new):
    HOOKS.append((name, old, new))


# --- 3. mouse down: chevron, title, chip, header drag -----------------

hook(
    "onMouseDown",
    """		oThis.SetFocusElement(FOCUS_OBJECT_THUMBNAILS);

		var pos = oThis.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);""",
    """		oThis.SetFocusElement(FOCUS_OBJECT_THUMBNAILS);

		if (0 === global_mouseEvent.Button && oThis.KiwiRail())
		{
			var railPos = oThis.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
			var railDown = oThis.KiwiRail().hostMouseDown(oThis, railPos.X, railPos.Y);
			if (railDown)
			{
				if ("rename" === railDown.action)
					oThis.KiwiStartRename(railDown.header);
				else if ("added" === railDown.action)
				{
					// lay the new bar out first, then edit its name:
					// nobody wants a section called "Neuer Abschnitt 3"
					oThis.KiwiRefresh();
					oThis.KiwiStartRename(oThis.KiwiRail()
						? oThis.KiwiRail().hostHeaderFor(oThis, railDown.section)
						: null);
				}
				else if ("track" !== railDown.action)
					oThis.KiwiRefresh();
				checkSelectionEnd();
				return false;
			}
		}

		var pos = oThis.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);""",
)

# --- 4. mouse move: header drag, then hover ---------------------------

hook(
    "onMouseMove header track",
    """		if (oThis.MouseDownTrack.IsStarted())
		{
			// this is a track for moving slides""",
    """		if (oThis.HeaderTrack && oThis.KiwiRail())
		{
			var trackPos = oThis.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
			var railTrack = oThis.KiwiRail().hostMouseMove(oThis, trackPos.X, trackPos.Y);
			oThis.m_oWordControl.m_oThumbnails.HtmlElement.style.cursor = railTrack.cursor || "move";
			return;
		}

		if (oThis.MouseDownTrack.IsStarted())
		{
			// this is a track for moving slides""",
)

hook(
    "onMouseMove hover",
    """		var cursor_moved = "default";

		if (pos.Page != -1)
		{
			oThis.m_arrPages[pos.Page].IsFocused = true;
			oThis.OnUpdateOverlay();

			cursor_moved = "pointer";
		} else if (_is_old_focused)
		{
			oThis.OnUpdateOverlay();
		}
""",
    """		var cursor_moved = "default";
		var railHover = oThis.KiwiRail()
			? oThis.KiwiRail().hostMouseMove(oThis, pos.X, pos.Y)
			: null;
		if (railHover)
		{
			if (railHover.cursor)
				cursor_moved = railHover.cursor;
			oThis.m_oWordControl.m_oThumbnails.HtmlElement.title = railHover.tooltip || "";
		}

		if (pos.Page != -1)
		{
			oThis.m_arrPages[pos.Page].IsFocused = true;
			cursor_moved = railHover && railHover.cursor ? cursor_moved : "pointer";
		}
		if (pos.Page != -1 || _is_old_focused || (railHover && railHover.repaint))
			oThis.OnUpdateOverlay();
""",
)

# --- 5. mouse up: finish a header drag --------------------------------

hook(
    "onMouseUp header track",
    """		oThis.CheckNeedAnimateScrolls(-1);

		if (!oThis.MouseDownTrack.IsStarted())
			return;""",
    """		oThis.CheckNeedAnimateScrolls(-1);

		if (oThis.HeaderTrack && oThis.KiwiRail())
		{
			var railUp = oThis.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
			if (oThis.KiwiRail().hostMouseUp(oThis, railUp.X, railUp.Y))
				oThis.KiwiRefresh();
			return;
		}

		if (!oThis.MouseDownTrack.IsStarted())
			return;""",
)

hook(
    "onMouseLeave clear hover",
    """		for (var i = 0; i < pages_count; i++)
		{
			oThis.m_arrPages[i].IsFocused = false;
		}
		oThis.OnUpdateOverlay();
	};""",
    """		for (var i = 0; i < pages_count; i++)
		{
			oThis.m_arrPages[i].IsFocused = false;
		}
		if (oThis.KiwiRail())
			oThis.KiwiRail().hostClearHover(oThis);
		oThis.OnUpdateOverlay();
	};""",
)

# --- 6. the kiwi helpers on the manager -------------------------------

hook(
    "manager helpers",
    """	this.onCheckUpdate = function()
	{""",
    """	// --- kiwi section rail: every hook above calls only these ---------
	this.KiwiRail = function () {
		var rail = window.AscCommonSlide && window.AscCommonSlide.PresentationSectionRail;
		return (rail && this.sectionPlan) ? rail : null;
	};
	this.KiwiRefresh = function () {
		this.CheckSizes();
		this.ClearCacheAttack();
		this.OnPaint();
	};
	this.KiwiStartRename = function (header) {
		var rail = window.AscCommonSlide && window.AscCommonSlide.PresentationSectionRail;
		var oThis2 = this;
		if (!rail || !header)
			return;
		rail.startRename(
			this,
			this.m_oWordControl.m_oThumbnails.HtmlElement,
			header,
			this.m_oWordControl.m_oLogicDocument,
			AscCommon.AscBrowser.retinaPixelRatio,
			function () { oThis2.KiwiRefresh(); }
		);
	};

	this.onCheckUpdate = function()
	{""",
)

# --- 7. paint: skip hidden thumbs, draw the bars ----------------------

hook(
    "OnPaint skip hidden",
    """			const oSlide = arrSlides[slideIndex];
			const slideType = oSlide.deleteLock.Lock.Get_Type();""",
    """			if (page && page.sectionHidden)
				continue;

			const oSlide = arrSlides[slideIndex];
			const slideType = oSlide.deleteLock.Lock.Get_Type();""",
)

hook(
    "OnUpdateOverlay draw",
    """		this.drawThumbnailsBorders(context, canvasWidth, canvasHeight);

		if (this.MouseDownTrack.IsDragged()) {""",
    """		this.drawThumbnailsBorders(context, canvasWidth, canvasHeight);
		if (this.KiwiRail())
			this.KiwiRail().hostDraw(this, context, GlobalSkin);

		if (this.MouseDownTrack.IsDragged()) {""",
)

# --- 8. the insertion line ---------------------------------------------
#
# Upstream takes the line's cross-axis extent from m_arrPages[0]. A
# collapsed first section would otherwise hand it a zero-height page, so
# take the first page that is actually on screen.

hook(
    "insertion line horizontal extent",
    """			let topY, bottomY;
			if (this.m_arrPages.length > 0) {
				topY = this.m_arrPages[0].top + 4;
				bottomY = this.m_arrPages[0].bottom - 4;
			} else {""",
    """			let topY, bottomY;
			// kiwi: a collapsed slide keeps its column but has no height,
			// so page 0 is not necessarily a usable reference any more
			const oRef = this.m_arrPages.find(page => !page.sectionHidden);
			if (oRef) {
				topY = oRef.top + 4;
				bottomY = oRef.bottom - 4;
			} else {""",
)

hook(
    "insertion line vertical",
    """			const y = oPage
				? (oPage.bottom + 1.5 * this.const_border_w) >> 0
				: this.const_offset_y / 2 >> 0;

			let _left_pos = 0;
			let _right_pos = canvasWidth;
			if (this.m_arrPages.length > 0) {
				_left_pos = this.m_arrPages[0].left + 4;
				_right_pos = this.m_arrPages[0].right - 4;
			}""",
    """			// kiwi: the rail resolved this drop in ConvertCoords2 and
			// left the answer on the manager; the index guard keeps a
			// stale gap out.
			const railGap = this.sectionGap;
			const snapY = (railGap && railGap.insertAt === nPosition) ? railGap.along : null;
			const y = snapY != null
				? snapY >> 0
				: oPage
					? (oPage.bottom + 1.5 * this.const_border_w) >> 0
					: this.const_offset_y / 2 >> 0;

			let _left_pos = 0;
			let _right_pos = canvasWidth;
			const oRef = this.m_arrPages.find(page => !page.sectionHidden);
			if (oRef) {
				_left_pos = oRef.left + 4;
				_right_pos = oRef.right - 4;
			}""",
)

# --- 9. ConvertCoords2: the rail owns the drop index ------------------

hook(
    "ConvertCoords2",
    """		let minDistance = Infinity;
		let minPositionPage = 0;

		for (let i = 0; i < this.m_arrPages.length; i++) {
			const page = this.m_arrPages[i];

			let distanceToStart, distanceToEnd;""",
    """		// kiwi: with section bars on screen the rail owns this answer --
		// it is the same layout that drew the bars, so the line the user
		// sees and the index the model gets cannot disagree.
		const railGap = this.KiwiRail()
			? this.KiwiRail().hostDropIndex(this, convertedX, convertedY, !isHorizontalThumbnails)
			: null;
		if (railGap)
			return railGap.insertAt;
		this.sectionGap = null;

		let minDistance = Infinity;
		let minPositionPage = 0;

		for (let i = 0; i < this.m_arrPages.length; i++) {
			const page = this.m_arrPages[i];
			if (page.sectionHidden)
				continue;

			let distanceToStart, distanceToEnd;""",
)

# --- 10. CalculatePlaces: ask for the plan, apply the slots -----------

hook(
    "CalculatePlaces layout",
    """		const totalSlidesCount = this.GetSlidesCount();
		for (let slideIndex = 0; slideIndex < totalSlidesCount; slideIndex++) {""",
    """		const totalSlidesCount = this.GetSlidesCount();
		const railApi = window.AscCommonSlide && window.AscCommonSlide.PresentationSectionRail;
		const sectionPlan = railApi ? railApi.hostLayout(this, {
			slideCount: totalSlidesCount,
			vertical: isVerticalThumbnails,
			startOffset: startOffset,
			currentScroll: currentScrollPx,
			thumbW: thSlideWidthPx,
			thumbH: thSlideHeightPx,
			ratio: AscCommon.AscBrowser.retinaPixelRatio,
			slideStep: (isVerticalThumbnails ? thSlideHeightPx : thSlideWidthPx) + 3 * this.const_border_w,
			railW: isVerticalThumbnails ? canvasWidth : canvasHeight
		}) : null;
		for (let slideIndex = 0; slideIndex < totalSlidesCount; slideIndex++) {""",
)

hook(
    "CalculatePlaces slots",
    """			const slideData = this.m_oWordControl.m_oLogicDocument.GetSlide(slideIndex);
			const slideRect = this.m_arrPages[slideIndex];
			slideRect.pageIndex = slideIndex;
""",
    """			const slideData = this.m_oWordControl.m_oLogicDocument.GetSlide(slideIndex);
			const slideRect = this.m_arrPages[slideIndex];
			slideRect.pageIndex = slideIndex;
			// kiwi: a bar takes its own space before the slide it heads.
			// A hidden slide keeps its column and loses only its height,
			// so everything that reads m_arrPages stays sane.
			slideRect.sectionHidden = false;
			if (sectionPlan && sectionPlan.slots[slideIndex]) {
				startOffset += sectionPlan.slots[slideIndex].advance;
				if (sectionPlan.slots[slideIndex].hidden) {
					slideRect.sectionHidden = true;
					slideRect.left = isRightToLeft ? this.const_offset_r : this.const_offset_x;
					slideRect.right = slideRect.left + thSlideWidthPx;
					slideRect.top = startOffset - currentScrollPx;
					slideRect.bottom = slideRect.top;
					continue;
				}
			}
""",
)

# --- 11. the scrollbar has to know about the bars ---------------------

hook(
    "CheckSizes extra",
    """			cumulativeThumbnailLength = isHorizontalOrientation
				? thumbnailWidth * slidesCount
				: thumbnailHeight * slidesCount;""",
    """			cumulativeThumbnailLength = isHorizontalOrientation
				? thumbnailWidth * slidesCount
				: thumbnailHeight * slidesCount;
			const railSize = window.AscCommonSlide && window.AscCommonSlide.PresentationSectionRail;
			if (railSize && !this.IsMasterMode() && oPresentation && oPresentation.Sections
				&& oPresentation.Sections.length && !oPresentation.IsVisioEditor()) {
				cumulativeThumbnailLength += railSize.extraLength(
					oPresentation,
					slidesCount,
					(isHorizontalOrientation ? thumbnailWidth : thumbnailHeight) + 3 * this.const_border_w,
					AscCommon.AscBrowser.retinaPixelRatio
				);
			}""",
)


def main(path):
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    # Most hooks contain their own anchor, so a second run would nest
    # them instead of failing. Today only the order saves us; say so.
    if "kiwi" in text:
        sys.stderr.write("%s already carries kiwi hooks; start from a "
                         "pristine checkout\n" % path)
        return 1
    # AGPL v3 section 5(a) and ONLYOFFICE's additional term 2: a modified
    # file must say so. Prepended, not anchored, so it costs nothing on a
    # rebase.
    NOTICE = ("/*\n"
              " * MODIFIED by KI macht Schule gGmbH, 2026: section rail in the\n"
              " * thumbnail strip. See KIWI-CHANGES.md for the full delta;\n"
              " * the hooks below are applied by apply-kiwi-rail.py.\n"
              " */\n")
    head = text.index(" */\n") + len(" */\n")
    text = text[:head] + NOTICE + text[head:]

    for name, old, new in HOOKS:
        found = text.count(old)
        if found != 1:
            sys.stderr.write(
                "anchor %r found %d times (expected 1)\n" % (name, found))
            return 1
        text = text.replace(old, new, 1)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text)
    print("applied %d hooks to %s" % (len(HOOKS), path))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
