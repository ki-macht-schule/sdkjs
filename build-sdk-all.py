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
"""Concatenate slide/sdk-all.js exactly as configs/slide.json lists it.

This is the whole build. The Document Server serves ``sdk-all.js`` as a
plain concatenation of the files in ``configs/slide.json`` -- "min" then
"common", in order -- so no toolchain is involved and the output is
deterministic: the same commit always produces the same bytes.

That matters beyond convenience. The image must be reproducible from
published source (AGPL section 13), and a ``docker cp`` of a locally
built blob is not. The overlay Dockerfile calls this script.

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
    " */\n"
)


def sources(root):
    """The file list, in load order, as the editor expects it."""
    with open(os.path.join(root, "configs", "slide.json"), encoding="utf-8") as fh:
        sdk = json.load(fh)["sdk"]
    out = []
    for group in ("min", "common"):
        for name in sdk.get(group, []):
            if name not in out:
                out.append(name)
    return out


def build(root, out_path):
    names = sources(root)
    missing = [n for n in names if not os.path.isfile(os.path.join(root, n))]
    if missing:
        sys.stderr.write("missing sources:\n  %s\n" % "\n  ".join(missing))
        return 1
    parts = [BANNER]
    for name in names:
        with open(os.path.join(root, name), encoding="utf-8") as fh:
            body = fh.read()
        # A file that ends inside a line comment would swallow the next
        # one; the newline is not cosmetic.
        parts.append("\n;// --- %s ---\n%s\n" % (name, body))
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write("".join(parts))
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
