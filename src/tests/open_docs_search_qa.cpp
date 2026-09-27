// QA harness for "Search open documents instead of files on disk" (stage 1):
//
//   1. Source switch in Find in Files: files open in N++ are searched via
//      their live Scintilla document; everything else takes the disk path.
//   2. AttachedDoc / probe AddRef / DocPtrReleaser reference protocol and the
//      hidden buffer's own document (create()) - modeled against Scintilla's
//      documented semantics (CREATEDOCUMENT refs the new document once;
//      SETDOCPOINTER releases current, sets new, refs new; 0 creates a fresh one).
//   3. Visibility gate: hidden views are not probed (NPPM_ACTIVATEDOC would
//      be refused and doc state read from the wrong document).
//   4. Size-limit policy for attached documents (noteSkip -> TooLarge) and
//      its effect on the searched-files denominator.
//   5. Replace uses only dirty entries (skip), Find uses every docPtr.
//   6. Column mode in the scan loops: every file loads its delimiters with
//      loadDelimitersForScan (parse, full scan of the bound document). The
//      editor path handleDelimiterPositions is not used there: a pending
//      document switch made it rebind the scan to the active editor view.
//   7. A run that cycles tabs holds a DocCycleGuard: the switches are MR's own,
//      so the FlowTabs cleanup machine stays out of the run (no posted hop, the
//      editor ends on the user's tab) and the panel's column state survives.
//      A later real user switch still cleans the padding. Every activation
//      drops the view's flow tab stops and column styles, so the guard has to
//      rebuild them for the document it returns to.
//
// Mirrors the panel/guard logic verbatim where mirrorable; Scintilla calls
// are replaced by a reference-counting mock that asserts protocol violations.
//
// Build: g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined
//        -o open_docs_search_qa open_docs_search_qa.cpp
#include <cstdio>
#include <cwctype>
#include <map>
#include <set>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

static int failures = 0;
static void CHECK(const std::string& d, bool ok) {
    std::printf("%-4s %s\n", ok ? "PASS" : "FAIL", d.c_str());
    if (!ok) ++failures;
}

// ------------------------------------------------ Scintilla document mock
// Implements the documented semantics of SCI_GETDOCPOINTER / SCI_SETDOCPOINTER /
// SCI_ADDREFDOCUMENT / SCI_RELEASEDOCUMENT / SCI_CREATEDOCUMENT for one view.
// Documents are shared by all views (one refcount each). Asserts misuse.
struct DocSpace {
    std::map<int, int> ref;          // docId -> refcount
    std::set<int> destroyed;
    int nextId = 100;
    bool protocolError = false;
};

struct MockSci {
    std::map<int, int>& ref;
    std::set<int>& destroyed;
    int& nextId;
    bool& protocolError;
    int current = 0;

    explicit MockSci(DocSpace& s) : ref(s.ref), destroyed(s.destroyed), nextId(s.nextId), protocolError(s.protocolError) {}
    int freshDoc() { int id = nextId++; ref[id] = 0; return id; }
    void init() { current = freshDoc(); addref(current); }   // view holds its doc
    int createDocument() { int id = freshDoc(); addref(id); return id; }   // SCI_CREATEDOCUMENT: held once by the caller

    void addref(int d) {
        if (!ref.count(d) || destroyed.count(d)) { protocolError = true; return; }
        ++ref[d];
    }
    void release(int d) {
        if (!ref.count(d) || destroyed.count(d) || ref[d] <= 0) { protocolError = true; return; }
        if (--ref[d] == 0) destroyed.insert(d);
    }
    // SCI_SETDOCPOINTER: view releases current, then holds the new doc (0 = fresh).
    void setDoc(int d) {
        release(current);
        current = (d == 0) ? freshDoc() : d;
        addref(current);
    }
    int getDoc() const { return current; }
    bool alive(int d) const { return ref.count(d) && !destroyed.count(d) && ref.at(d) > 0; }
};

// -------------------------- MIRROR: HiddenSciGuard::create(), own document (in mock terms)
static int createOwnDocM(MockSci& hidden) {
    const int doc = hidden.createDocument();   // SCI_CREATEDOCUMENT
    if (!doc) return 0;
    hidden.setDoc(doc);                        // SCI_SETDOCPOINTER: the view holds it as well
    hidden.release(doc);                       // SCI_RELEASEDOCUMENT: the view alone
    return doc;
}

// -------------------------- MIRROR: AttachedDoc (HiddenSciGuard.h, in mock terms)
struct AttachedDocM {
    MockSci* g = nullptr;
    int ownDoc = 0;
    void attach(MockSci& hidden, int foreignDoc) {
        g = &hidden;
        ownDoc = g->getDoc();       // SCI_GETDOCPOINTER
        g->addref(ownDoc);          // keep ours alive
        g->setDoc(foreignDoc);      // releases ours, refs foreign
    }
    ~AttachedDocM() {
        if (!g) return;
        g->setDoc(ownDoc);          // releases foreign, refs ours
        g->release(ownDoc);         // drop keep-alive ref
    }
};

// -------------------------- MIRROR: source switch + probe/release policy
struct OpenScanDocM { int view; long index; bool dirty; int docPtr; };

static std::wstring openDocPathKey(std::wstring s) {
    for (auto& c : s) c = towlower(c);
    return s;
}

static const OpenScanDocM* pickSource(
    const std::unordered_map<std::wstring, OpenScanDocM>& openDocs, const std::wstring& path)
{
    if (openDocs.empty()) return nullptr;
    const auto od = openDocs.find(openDocPathKey(path));
    return (od != openDocs.end() && od->second.docPtr) ? &od->second : nullptr;
}

// Probe-side AddRef policy: only when grabDocPtrs, only non-null, one per path
// (clone in both views deduped via probeKeys).
static std::unordered_map<std::wstring, OpenScanDocM> probeM(
    MockSci& viewSci,
    const std::vector<std::pair<std::wstring, int>>& openDocsInViews,  // (key, docId), may repeat
    const std::unordered_set<std::wstring>& scanSet,
    const std::unordered_set<std::wstring>& dirtySet,
    bool grabDocPtrs)
{
    std::unordered_map<std::wstring, OpenScanDocM> result;
    std::unordered_set<std::wstring> probeKeys;
    for (const auto& [key, docId] : openDocsInViews) {
        if (key.empty() || scanSet.count(key) == 0 || !probeKeys.insert(key).second) continue;
        OpenScanDocM e{ 0, 0, dirtySet.count(key) != 0, 0 };
        if (grabDocPtrs) {
            e.docPtr = docId;
            if (e.docPtr) viewSci.addref(e.docPtr);   // keep alive; caller releases
        }
        result.emplace(key, e);
    }
    return result;
}

struct DocPtrReleaserM {
    MockSci& g;
    const std::unordered_map<std::wstring, OpenScanDocM>& m;
    ~DocPtrReleaserM() {
        for (const auto& kv : m)
            if (kv.second.docPtr) g.release(kv.second.docPtr);
    }
};

// -------------------------- MIRROR: guard skip counters + denominator
struct SkipCountersM {
    size_t binary = 0, large = 0, unreadable = 0, undecodable = 0;
    void noteTooLarge() { ++large; }
    size_t total() const { return binary + large + unreadable + undecodable; }
};
static size_t searchedFiles(size_t reached, const SkipCountersM& c) {
    return (reached > c.total()) ? (reached - c.total()) : 0;
}

// -------------------------- MIRROR: visibility gate (collectOpenScanDocs)
static bool viewIsProbed(bool handleValid, bool visible) {
    return handleValid && visible;   // if (!viewSci || !IsWindowVisible(viewSci)) continue;
}

int main() {
    std::printf("=== S1 source switch ===\n\n");
    {
        std::unordered_map<std::wstring, OpenScanDocM> docs;
        docs.emplace(L"c:\\p\\a.txt", OpenScanDocM{ 0, 0, true, 42 });
        docs.emplace(L"c:\\p\\b.txt", OpenScanDocM{ 0, 0, false, 0 });   // no doc ptr grabbed
        CHECK("S1 open file with docPtr -> live document, case-insensitive",
              pickSource(docs, L"C:\\P\\A.TXT") != nullptr);
        CHECK("S1 zero docPtr -> disk path", pickSource(docs, L"c:\\p\\b.txt") == nullptr);
        CHECK("S1 file not open -> disk path", pickSource(docs, L"c:\\p\\c.txt") == nullptr);
        std::unordered_map<std::wstring, OpenScanDocM> empty;
        CHECK("S1 option off / no open docs -> disk path (zero-cost)",
              pickSource(empty, L"c:\\p\\a.txt") == nullptr);
    }

    std::printf("\n=== S2 reference protocol (attach/detach/release) ===\n\n");
    {
        // Normal scan: probe -> attach -> search -> detach -> releaser.
        DocSpace docSpace;
        MockSci userView(docSpace);   userView.init();
        MockSci hidden(docSpace);     hidden.init();
        const int F = userView.getDoc();                 // user's live doc, refcount 1
        const int first = hidden.getDoc();               // the document the view was created with
        const int H = createOwnDocM(hidden);             // hidden view's own doc
        CHECK("S2 create(): the view's first document is gone, its own one held by the view alone",
              !hidden.alive(first) && hidden.getDoc() == H && hidden.ref[H] == 1);

        auto docs = probeM(userView, { { L"k", F } }, { L"k" }, {}, /*grab*/true);
        CHECK("S2 probe AddRef'd the foreign doc once", userView.ref[F] == 2);
        {
            DocPtrReleaserM rel{ userView, docs };
            {
                AttachedDocM att;
                att.attach(hidden, F);
                CHECK("S2 while attached: hidden view holds foreign, own doc kept alive",
                      hidden.getDoc() == F && hidden.alive(H));
            }
            CHECK("S2 after detach: hidden back on its own doc, foreign released by hidden",
                  hidden.getDoc() == H && userView.ref[F] == 2);
        }
        CHECK("S2 after releaser: only the user's view still holds the doc",
              userView.ref[F] == 1 && userView.alive(F));
        CHECK("S2 no protocol violations", !userView.protocolError && !hidden.protocolError);
    }
    {
        // Tab closed mid-scan while attached: doc must survive until detach+release.
        DocSpace docSpace;
        MockSci userView(docSpace);   userView.init();
        MockSci hidden(docSpace);     hidden.init();
        const int H = createOwnDocM(hidden);
        const int F = userView.getDoc();
        auto docs = probeM(userView, { { L"k", F } }, { L"k" }, {}, true);
        {
            DocPtrReleaserM rel{ userView, docs };
            AttachedDocM att;
            att.attach(hidden, F);
            userView.setDoc(0);   // user closes the tab -> view drops its ref
            CHECK("S2 tab close mid-attach: doc kept alive by probe ref + hidden view",
                  userView.alive(F));
        }
        CHECK("S2 after detach + releaser: doc destroyed exactly once, no leak",
              !userView.alive(F) && !userView.protocolError && !hidden.protocolError);
        CHECK("S2 and the hidden view is back on its own document", hidden.getDoc() == H && hidden.ref[H] == 1);
    }
    {
        // Clone open in both views: deduped -> exactly one AddRef, one release.
        DocSpace docSpace;
        MockSci userView(docSpace);   userView.init();
        const int F = userView.getDoc();
        auto docs = probeM(userView, { { L"k", F }, { L"k", F } }, { L"k" }, {}, true);
        CHECK("S2 clone in both views: one entry, one AddRef",
              docs.size() == 1 && userView.ref[F] == 2);
        { DocPtrReleaserM rel{ userView, docs }; }
        CHECK("S2 clone release balanced", userView.ref[F] == 1 && !userView.protocolError);
    }

    std::printf("\n=== S3 visibility gate ===\n\n");
    {
        CHECK("S3 visible view is probed", viewIsProbed(true, true));
        CHECK("S3 hidden view is skipped", !viewIsProbed(true, false));
        CHECK("S3 missing handle is skipped", !viewIsProbed(false, true));
    }

    std::printf("\n=== S4 size limit on attached documents ===\n\n");
    {
        SkipCountersM c;
        const size_t maxBytes = 100;
        const size_t docLen = 250;
        const bool skip = (maxBytes > 0 && docLen > maxBytes);
        if (skip) c.noteTooLarge();
        CHECK("S4 over-limit attached doc -> TooLarge skip", skip && c.large == 1);
        CHECK("S4 denominator subtracts the buffer skip", searchedFiles(10, c) == 9);
        SkipCountersM c2;
        const size_t noLimit = 0;   // getEffectiveMaxFileSize() == 0 when disabled
        CHECK("S4 limit disabled -> no skip", !(noLimit > 0 && docLen > noLimit) && c2.total() == 0);
    }

    std::printf("\n=== S5 dirty policy: Replace skips, Find attaches ===\n\n");
    {
        DocSpace docSpace;
        MockSci userView(docSpace); userView.init();
        const int F1 = userView.getDoc();
        userView.setDoc(0); const int F2 = userView.getDoc(); (void)F1;

        // Replace: grabDocPtrs=false -> dirty drives the skip set, no refs taken.
        auto rep = probeM(userView, { { L"a", F2 }, { L"b", F2 } },
                          { L"a", L"b" }, { L"a" }, /*grab*/false);
        std::unordered_set<std::wstring> dirtyOpenPaths;
        for (const auto& [key, d] : rep) if (d.dirty) dirtyOpenPaths.insert(key);
        CHECK("S5 Replace: only the dirty file is skipped",
              dirtyOpenPaths.size() == 1 && dirtyOpenPaths.count(L"a") == 1);
        CHECK("S5 Replace: no doc refs taken", userView.ref[F2] == 1);

        // Find: every open file gets a docPtr, dirty or not (N++ parity).
        auto fnd = probeM(userView, { { L"a", F2 }, { L"b", F2 } },
                          { L"a", L"b" }, { L"a" }, /*grab*/true);
        CHECK("S5 Find: clean AND dirty open files carry a docPtr",
              fnd.at(L"a").docPtr != 0 && fnd.at(L"b").docPtr != 0);
        { DocPtrReleaserM rel{ userView, fnd }; }
        CHECK("S5 Find: refs balanced after release", userView.ref[F2] == 1 && !userView.protocolError);
    }

    std::printf("\n=== S6 column mode: delimiters per file, on the scanned document ===\n\n");
    {
        // Where the panel's Scintilla calls go: the scanned document (0) or the active editor (1)
        struct PanelM { int bound = 0; bool documentSwitched = false; int scans = 0; };
        // MIRROR (before): handleDelimiterPositions(LoadAll) in the scan loop. A pending
        // document switch ran handleClearDelimiterState -> pointerToScintilla.
        auto editorPath = [](PanelM& p) {
            if (p.documentSwitched) { p.bound = 1; p.documentSwitched = false; }
            ++p.scans;
        };
        // MIRROR (now): loadDelimitersForScan - parse, scan the bound document, nothing else
        auto scanPath = [](PanelM& p) { ++p.scans; };

        PanelM before; before.documentSwitched = true;   // a tab switch since the last column operation
        editorPath(before);
        CHECK("S6 before: the first file after a tab switch was searched in the active editor (the bug)",
              before.bound == 1);
        PanelM now; now.documentSwitched = true;
        for (int file = 0; file < 3; ++file) scanPath(now);
        CHECK("S6 now: every file scanned on its own document, the switch left to the editor",
              now.bound == 0 && now.scans == 3 && now.documentSwitched);
    }

    std::printf("\n=== S7 a run cycles tabs: the FlowTabs cleanup machine stays out ===\n\n");
    {
        struct Panel {
            int  prevBuf = 0;                  // g_prevBufId
            std::unordered_set<int> padBufs;   // g_padBufs: buffers carrying flow padding
            bool cleanInProgress = false;      // g_cleanInProgress
            int  pendingCleanId = 0;           // g_pendingCleanId
            int  returnBufId = 0;              // g_returnBufId
            bool docSwitched = false;          // documentSwitched
            bool cycleActive = false;          // DocCycleGuard::active()
            std::unordered_set<int> marks;     // buffers with painted column highlight
            bool sorted = true;                // the user's sort state
            bool viewStops = true;             // per-line flow tab stops (view state)
            bool viewStyles = true;            // column styles, redefined per language
            int  delimMapOf = 0;               // buffer the delimiter map describes
            std::vector<int> scannedWith;      // map used per searched document
        };
        struct Editor {
            std::vector<int> tabs;
            int active = 0;
            std::vector<int> posted;           // queued NPPM_ACTIVATEDOC (PostMessage)
            int buf() const { return tabs[static_cast<size_t>(active)]; }
            int indexOf(int bufId) const {
                for (size_t i = 0; i < tabs.size(); ++i)
                    if (tabs[i] == bufId) return static_cast<int>(i);
                return -1;
            }
        };
        // MIRROR: onDocumentSwitched(). The guard's early return keeps the
        // per-view undo mask (not modeled) and skips everything below it.
        auto onDocumentSwitched = [](Editor& ed, Panel& p) {
            if (p.cycleActive) return;
            const int cur = ed.buf();
            if (p.prevBuf == 0) p.prevBuf = cur;
            if (p.cleanInProgress && p.pendingCleanId == cur) {          // PHASE A
                p.padBufs.erase(cur);
                p.cleanInProgress = false;
                p.pendingCleanId = 0;
                ed.posted.push_back(p.returnBufId);
                return;
            }
            if (!p.cleanInProgress && p.prevBuf != cur                   // PHASE B
                && p.padBufs.count(p.prevBuf) != 0) {
                p.returnBufId = cur;
                p.pendingCleanId = p.prevBuf;
                p.cleanInProgress = true;
                ed.posted.push_back(p.prevBuf);
                return;
            }
            p.docSwitched = true;                                        // arrival
            p.marks.erase(cur);
            p.sorted = false;
            p.prevBuf = cur;
        };
        auto activate = [&](Editor& ed, Panel& p, int index) {
            ed.active = index;
            p.viewStops = false;               // SCI_SETDOCPOINTER + defineDocType
            p.viewStyles = false;
            onDocumentSwitched(ed, p);
        };
        auto drain = [&](Editor& ed, Panel& p) {                 // message loop
            while (!ed.posted.empty()) {
                const int target = ed.posted.front();
                ed.posted.erase(ed.posted.begin());
                const int idx = ed.indexOf(target);
                if (idx >= 0) activate(ed, p, idx);
            }
        };
        // MIRROR (before): handleDelimiterPositions(LoadAll) rescans only when the
        // switch flagged a new document (clear -> empty map) or the map is empty.
        auto loadEditorPath = [](Panel& p, int cur) {
            if (p.docSwitched) { p.delimMapOf = 0; p.docSwitched = false; }
            if (p.delimMapOf == 0) p.delimMapOf = cur;
            p.scannedWith.push_back(p.delimMapOf);
        };
        // MIRROR (now): loadDelimitersForScan always scans the bound document
        auto loadScanPath = [](Panel& p, int cur) {
            p.delimMapOf = cur;
            p.scannedWith.push_back(cur);
        };
        // MIRROR: restoreColumnVisuals() in the guard's restore
        auto restoreVisuals = [](Panel& p) {
            if (!p.padBufs.empty()) p.viewStops = true;
            if (!p.marks.empty()) p.viewStyles = true;
        };
        auto run = [&](Editor& ed, Panel& p, bool gated, bool rebuild = true) {
            const int savedIdx = ed.active;
            p.cycleActive = gated;                               // DocCycleGuard
            for (size_t i = 0; i < ed.tabs.size(); ++i) {
                activate(ed, p, static_cast<int>(i));
                if (gated) loadScanPath(p, ed.buf()); else loadEditorPath(p, ed.buf());
            }
            activate(ed, p, savedIdx);                           // cycle.restore()
            if (gated && rebuild) restoreVisuals(p);
            p.cycleActive = false;
            drain(ed, p);
        };

        // The user sits on the last tab: flow padding on, column highlight painted, sorted
        const std::vector<int> tabs{ 11, 22, 33 };
        Editor edB{ tabs, 2, {} };
        Panel  pB;
        pB.prevBuf = 33; pB.padBufs = { 33 }; pB.marks = { 33 }; pB.delimMapOf = 33;
        run(edB, pB, /*gated=*/false);
        CHECK("S7 before: the run ends on the first tab, not the user's",
              edB.buf() == 11);
        CHECK("S7 before: two documents were searched with a foreign delimiter map",
              pB.scannedWith == std::vector<int>({ 33, 22, 22 }));
        CHECK("S7 before: highlight, sort state and padding of the user's document are gone",
              pB.marks.empty() && !pB.sorted && pB.padBufs.empty());

        Editor edN{ tabs, 2, {} };
        Panel  pN;
        pN.prevBuf = 33; pN.padBufs = { 33 }; pN.marks = { 33 }; pN.delimMapOf = 33;
        run(edN, pN, /*gated=*/true);
        CHECK("S7 now: nothing was posted, the run ends on the user's tab",
              edN.posted.empty() && edN.buf() == 33);
        CHECK("S7 now: every document was searched with its own delimiters",
              pN.scannedWith == tabs);
        CHECK("S7 now: highlight, sort state and padding of the user's document survive",
              pN.marks.count(33) == 1 && pN.sorted && pN.padBufs.count(33) == 1);
        CHECK("S7 now: the run left no document switch behind",
              pN.prevBuf == 33 && !pN.docSwitched && !pN.cleanInProgress);

        // Without the rebuild the flags would lie: padding in the text, no stops
        Editor edL{ tabs, 2, {} };
        Panel  pL;
        pL.prevBuf = 33; pL.padBufs = { 33 }; pL.marks = { 33 }; pL.delimMapOf = 33;
        run(edL, pL, /*gated=*/true, /*rebuild=*/false);
        CHECK("S7 trap: padding kept but stops and styles dropped by the cycle",
              pL.padBufs.count(33) == 1 && !pL.viewStops && !pL.viewStyles);
        CHECK("S7 now: stops and styles are rebuilt for the document returned to",
              pN.viewStops && pN.viewStyles);

        // The machine itself still works: a real switch cleans the padding
        activate(edN, pN, 0);
        drain(edN, pN);
        CHECK("S7 now: a later user switch still cleans the padding",
              pN.padBufs.empty() && edN.buf() == 11 && pN.docSwitched);
    }

    std::printf("\n%s (%d failure%s)\n",
                failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED",
                failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
