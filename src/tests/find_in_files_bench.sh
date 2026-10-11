#!/usr/bin/env bash
# This file is part of the MultiReplace plugin for Notepad++.
# Copyright (C) 2026 Thomas Knoefel
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

# Builds the Find in Files test bench (find_in_files_bench.cpp) against the pinned Notepad++
# sources, creates its corpora and runs the measurement.
#
# Needs bash, git, g++ with C++17 and access to github.com. Runs on Linux or WSL.
# Usage: find_in_files_bench.sh [work dir] [rounds] [tools]
#   work dir  sources, objects, corpora and results; default ${TMPDIR:-/tmp}/find_in_files_bench
#   rounds    measured rounds per cell; default 15
#   tools     comma-separated; default MR,NppNew,NppOld; MR's later steps: MR+utf8, MR+prefilter,
#             MR+dock, MR+all
# Results: <work dir>/results/summary.md and <work dir>/results/raw.csv
set -euo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
work="${1:-${TMPDIR:-/tmp}/find_in_files_bench}"
rounds="${2:-15}"
tools="${3:-MR,NppNew,NppOld}"

# Notepad++ master containing 20d8f02 "Improve Find in Files performance"; also corpus content
npp_commit=697f46b6b6d7245408185370b85306a1f03da179
# MultiReplace sources used as corpus content
mr_commit=a43e1628b559c988496143f0847c6298ae409093
# Fingerprints of the corpora of the published baseline
expected_source=c28452ea1406d76b
expected_mixed=ff4facbed884c827

mkdir -p "$work"
cd "$work"

# Sparse, blob-less checkout of the folders needed at a fixed commit
fetch() {
    local url=$1 dir=$2 commit=$3
    shift 3
    if [ ! -d "$dir/.git" ]; then
        git clone --quiet --filter=blob:none --no-checkout "$url" "$dir"
    fi
    git -C "$dir" sparse-checkout set --cone "$@"
    git -C "$dir" checkout --quiet "$commit"
}
fetch https://github.com/notepad-plus-plus/notepad-plus-plus.git npp "$npp_commit" \
    PowerEditor/src scintilla/include scintilla/src scintilla/win32 boostregex
fetch https://github.com/daddel80/notepadpp-multireplace.git mr "$mr_commit" src

npp="$work/npp"
uchardet="$npp/PowerEditor/src/uchardet"
flags=(-std=c++17 -O2 -DNDEBUG)
mkdir -p obj

# Scintilla document layer with the Notepad++ Boost regex bridge (SCI_OWNREGEX)
for f in CellBuffer Document PerLine RunStyles Decoration CaseFolder CaseConvert CharClassify \
         CharacterCategoryMap CharacterType UniConversion DBCS ChangeHistory UndoHistory RESearch; do
    [ -f "obj/$f.o" ] || g++ "${flags[@]}" -DSCI_OWNREGEX -I"$npp/scintilla/include" -I"$npp/scintilla/src" \
        -c "$npp/scintilla/src/$f.cxx" -o "obj/$f.o"
done
for f in BoostRegExSearch UTF8DocumentIterator; do
    [ -f "obj/$f.o" ] || g++ "${flags[@]}" -DSCI_OWNREGEX -DBOOST_REGEX_STANDALONE -I"$npp/scintilla/include" \
        -I"$npp/scintilla/src" -I"$npp/boostregex" -c "$npp/boostregex/$f.cxx" -o "obj/$f.o"
done
# uchardet as bundled with Notepad++
for src in "$uchardet"/*.cpp; do
    o="obj/uchardet_$(basename "$src" .cpp).o"
    [ -f "$o" ] || g++ "${flags[@]}" -w -include cstdint -I"$uchardet" -c "$src" -o "$o"
done
g++ "${flags[@]}" -Wall -Wextra -pthread -I"$here/.." -I"$npp/scintilla/include" -I"$npp/scintilla/src" -I"$uchardet" \
    "$here/find_in_files_bench.cpp" obj/*.o -o find_in_files_bench

# Corpus "source": Notepad++ and MultiReplace sources four times, icons and test programs included
if [ ! -d corpus_source ]; then
    for c in c1 c2 c3 c4; do
        mkdir -p "corpus_source/$c"
        cp -R "$npp/PowerEditor/src" "corpus_source/$c/npp_src"
        cp -R "$npp/scintilla" "corpus_source/$c/scintilla"
        cp -R "$npp/boostregex" "corpus_source/$c/boostregex"
        cp -R "$work/mr/src" "corpus_source/$c/mr_src"
    done
fi
# Corpus "mixed": generated text in several encodings plus a few binaries
[ -d corpus_mixed ] || ./find_in_files_bench --generate-mixed corpus_mixed --gen-seed 1 --gen-files 1500

./find_in_files_bench --source corpus_source --source-text corpus_source --mixed corpus_mixed \
    --rounds "$rounds" --tools "$tools" --out results

for pair in "source:$expected_source" "mixed:$expected_mixed"; do
    name=${pair%%:*}
    fp=${pair#*:}
    if ! grep -q "Corpus \`$name\`: .*fingerprint $fp" results/summary.md; then
        echo "Note: corpus '$name' differs from the baseline corpus (expected fingerprint $fp)." \
             "Compare results only within this run." >&2
    fi
done
echo "Results: $work/results/summary.md"
