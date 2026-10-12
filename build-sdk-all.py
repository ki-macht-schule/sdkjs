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
"""Build slide/sdk-all.js from the ``common`` group in configs/slide.json.

Upstream splits the editor into two scripts that share one realm:

- ``sdk-all-min.js`` — the ``min`` group (api, apiDefines, …). The
  Document Server image already ships this file. Do not replace it.
- ``sdk-all.js`` — the ``common`` group (Presentation, HtmlPage, the
  kiwi rail), wrapped in one IIFE so its ``const`` / ``let`` names
  do not collide with ``min``.

``build/build.py`` does exactly that split. Concatenating ``min`` into
``sdk-all.js`` as well double-declares ``const c_oAscPresentationViewMode``
(and the rest of apiDefines). DoctRenderer and the browser then fail
to parse the second file: ``WordControl`` stays null and the UI dies
on ``m_oDrawingDocument``.

Usage:
    python3 build-sdk-all.py [<sdkjs-root>] -o <out.js>
"""
import argparse
import json
import os
import sys

BANNER = (
    "/*\n"
    " * sdk-all.js -- built by build-sdk-all.py from configs/slide.json.\n"
    " * Kiwi fork of ONLYOFFICE/sdkjs. See KIWI-CHANGES.md for the delta.\n"
    " * common group only; sdk-all-min.js stays the image's min bundle.\n"
    " */\n"
)
IIFE_OPEN = "(function(window, undefined) {\n"
IIFE_CLOSE = "})(window);\n"


def load_slide_sdk(root):
    with open(os.path.join(root, "configs", "slide.json"), encoding="utf-8") as fh:
        return json.load(fh)["sdk"]


def sources(root):
    """The common-group file list, in load order.

    ``min`` is already inside the image's ``sdk-all-min.js``. A name
    that appears in both groups is refused rather than silently
    dropped — that would hide a config error.
    """
    sdk = load_slide_sdk(root)
    min_names = list(sdk.get("min", []))
    common = []
    for name in sdk.get("common", []):
        if name not in common:
            common.append(name)
    leaked = [n for n in common if n in min_names]
    if leaked:
        sys.stderr.write(
            "min group leaked into common (would redeclare in one realm):\n  %s\n"
            % "\n  ".join(leaked)
        )
        raise SystemExit(1)
    return common


def build(root, out_path):
    names = sources(root)
    missing = [n for n in names if not os.path.isfile(os.path.join(root, n))]
    if missing:
        sys.stderr.write("missing sources:\n  %s\n" % "\n  ".join(missing))
        return 1
    parts = [BANNER, IIFE_OPEN]
    for name in names:
        with open(os.path.join(root, name), encoding="utf-8") as fh:
            body = fh.read()
        # A file that ends inside a line comment would swallow the next
        # one; the newline is not cosmetic.
        parts.append("\n;// --- %s ---\n%s\n" % (name, body))
    parts.append(IIFE_CLOSE)
    text = "".join(parts)
    if "const c_oAscPresentationViewMode" in text:
        sys.stderr.write(
            "sdk-all.js contains const c_oAscPresentationViewMode — "
            "the min group leaked; refuse to write\n"
        )
        return 1
    if "kiwiProtocol" not in text:
        sys.stderr.write("sdk-all.js has no kiwiProtocol — rail did not take\n")
        return 1
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(text)
    size = os.path.getsize(out_path)
    print("wrote %s (%d files, %.1f MB)" % (out_path, len(names), size / 1048576.0))
    return 0


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("root", nargs="?", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("-o", "--out", required=True)
    args = ap.parse_args(argv)
    return build(args.root, args.out)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
