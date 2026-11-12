// Ownership regressions for the MSOFFICE plugin's 4D string/variable/object
// handling. Unity build of the real plugin source (so its file-static command
// functions and helpers are reachable) plus the real SDK, C_TEXT/C_BLOB and
// support/4DPlugin-JSON.cpp, running against tests/ownership/fake_engine
// instead of 4D. Spreadsheets are synthetic workbooks created with the bundled
// OpenXLSX in a private temporary directory; no user documents are touched.
// Usage: harness [expected-failure ...]  (names of tests that must fail)
#include "4DPlugin-MSOFFICE.cpp"
#include "fake_engine.h"
#include <cstdio>
#include <functional>
#include <set>
#include <string>
#include <unistd.h>

using namespace fake4d;

static std::vector<std::string> gFailures;
#define CHECK(cond) do { if (!(cond)) gFailures.push_back(#cond); } while (0)

static std::string gDir;

static std::string tmp_path(const char *name) { return gDir + "/" + name; }

// 4D hands the plugin HFS paths on macOS (C_TEXT::copyPath converts them back)
static CUTF16String hfs(const std::string &posix) {
    CFURLRef url = CFURLCreateFromFileSystemRepresentation(kCFAllocatorDefault, (const UInt8 *)posix.c_str(),
                                                           (CFIndex)posix.size(), false);
    CFStringRef s = CFURLCopyFileSystemPath(url, kCFURLHFSPathStyle);
    CFIndex len = CFStringGetLength(s);
    CUTF16String u(len, 0);
    CFStringGetCharacters(s, CFRangeMake(0, len), (UniChar *)&u[0]);
    CFRelease(s);
    CFRelease(url);
    return u;
}

static void make_workbook(const std::string &path) {
    OpenXLSX::XLDocument doc;
    doc.create(path);
    auto wb = doc.workbook();
    wb.addWorksheet("Second");
    auto ws = wb.worksheet("Sheet1");
    ws.cell("A1").value() = "hello";
    ws.cell("B1").value() = 42;
    doc.save();
    doc.close();
}

// Plugin call with a PackagePtr of text/object/blob/longint slots.
struct Call {
    std::vector<CUTF16String> texts;
    std::vector<PA_Unistring> unistrings;
    std::vector<PA_ObjectRef> objects;
    std::vector<PA_Handle> blobs;
    std::vector<PA_long32> longs;
    std::vector<void *> slots;
    PA_ObjectRef result = nullptr;
    Call() { texts.reserve(8); unistrings.reserve(8); objects.reserve(8); blobs.reserve(8); longs.reserve(8); }
    Call &text(const CUTF16String &t) {
        texts.push_back(t);
        PA_Unistring u = {(PA_long32)texts.back().size(), (PA_Unichar *)texts.back().c_str(), 0, 0};
        unistrings.push_back(u);
        slots.push_back(&unistrings.back());
        return *this;
    }
    Call &path(const std::string &p) { return text(hfs(p)); }
    Call &object(Obj *o) { objects.push_back((PA_ObjectRef)o); slots.push_back(&objects.back()); return *this; }
    Call &blob(const void *bytes, size_t len) {
        PA_Handle h = PA_NewHandle((PA_long32)len);
        memcpy(PA_LockHandle(h), bytes, len);
        PA_UnlockHandle(h);
        blobs.push_back(h);
        slots.push_back(&blobs.back());
        return *this;
    }
    Call &longint(PA_long32 v) { longs.push_back(v); slots.push_back(&longs.back()); return *this; }
    Obj *run(PA_long32 selector) {
        PluginBlock pb = {slots.data(), &result, nullptr};
        PluginMain(selector, &pb);
        for (PA_Handle h : blobs) PA_DisposeHandle(h);
        blobs.clear();
        return (Obj *)result;
    }
};

static bool status_success(Obj *status) {
    return ob_get_b((PA_ObjectRef)status, L"success");
}

struct Test { const char *name; std::function<void()> body; };

static const Test kTests[] = {
    // --- LIVE: commands through PluginMain ---
    {"live_read_from_spreadsheet_all_sheets", [] { // L2 sheet titles + L4 ob_set_s
        std::string xlsx = tmp_path("read.xlsx");
        make_workbook(xlsx);
        Call c;
        c.path(xlsx).object(nullptr);
        Obj *status = c.run(2);
        CHECK(status != nullptr);
        CHECK(status_success(status));
        CHECK(collection_length(status, "sheets") == 2); // engine-side read: ob_get_c is tested separately
        release(status);
    }},
    {"live_read_from_spreadsheet_with_references", [] { // L3 parseReferencesObject
        std::string xlsx = tmp_path("read_refs.xlsx");
        make_workbook(xlsx);
        Obj *refs = new_object(), *cells = new_object(true);
        push_string(cells, "A1");
        put_ref(refs, "Sheet1", cells);
        release(cells);
        Call c;
        c.path(xlsx).object(refs);
        Obj *status = c.run(2);
        CHECK(status != nullptr);
        CHECK(status_success(status));
        release(status);
        release(refs);
    }},
    {"live_write_to_spreadsheet", [] { // L1 copy_document_native + L3 + L4
        std::string in = tmp_path("write_in.xlsx"), out = tmp_path("out/write_out.xlsx");
        make_workbook(in);
        Obj *refs = new_object(), *sheet = new_object();
        put_string(sheet, "A1", "changed");
        put_real(sheet, "B2", 7);
        put_ref(refs, "Sheet1", sheet);
        release(sheet);
        Call c;
        c.path(in).path(out).object(refs).text(CUTF16String());
        Obj *status = c.run(1);
        CHECK(status != nullptr);
        CHECK(status_success(status));
        CHECK(copy_documents() == 1);
        release(status);
        release(refs);
    }},
    {"live_verify_office_document", [] { // L5 ob_set_a
        static const char zip[] = "PK\x03\x04 not really a zip";
        Call c;
        c.blob(zip, sizeof(zip));
        Obj *status = c.run(3);
        CHECK(status != nullptr);
        std::string format;
        CHECK(status && get_string(status, "format", &format) && !format.empty());
        release(status);
    }},
    // --- LIVE: the same defects isolated at helper level ---
    {"live_copy_document_native", [] { // L1
        std::string in = tmp_path("copy_in.bin"), out = tmp_path("copy_out.bin");
        FILE *f = fopen(in.c_str(), "wb");
        fputs("x", f);
        fclose(f);
        C_TEXT src, dst;
        CUTF16String a = hfs(in), b = hfs(out);
        src.setUTF16String(a.c_str(), (uint32_t)a.size());
        dst.setUTF16String(b.c_str(), (uint32_t)b.size());
        copy_document_native(src, dst);
        CHECK(copy_documents() == 1);
        CHECK(access(out.c_str(), F_OK) == 0);
    }},
    {"live_parse_references_object", [] { // L3
        Obj *refs = new_object();
        put_string(refs, "k", "v");
        Json::Value value;
        CHECK(parseReferencesObject((PA_ObjectRef)refs, &value));
        CHECK(value["k"].asString() == "v");
        release(refs);
    }},
    {"live_ob_set_s_wide_key", [] { // L4
        Obj *o = new_object();
        ob_set_s((PA_ObjectRef)o, L"k", "v\xC3\xA9");
        std::string s;
        CHECK(get_string(o, "k", &s) && s.size() == 2);
        release(o);
    }},
    {"live_ob_set_a_wide", [] { // L5
        Obj *o = new_object();
        ob_set_a((PA_ObjectRef)o, L"format", L"Zip");
        std::string s;
        CHECK(get_string(o, "format", &s) && s == "Zip");
        release(o);
    }},
    // --- UNREACHABLE: built helpers with no caller ---
    {"unreachable_ob_set_s_narrow_key", [] {
        Obj *o = new_object();
        ob_set_s((PA_ObjectRef)o, "k", "v");
        std::string s;
        CHECK(get_string(o, "k", &s) && s == "v");
        release(o);
    }},
    {"unreachable_ob_set_a_utf16", [] {
        Obj *o = new_object();
        CUTF16String v;
        v.push_back('v');
        ob_set_a((PA_ObjectRef)o, L"k", &v);
        std::string s;
        CHECK(get_string(o, "k", &s) && s == "v");
        release(o);
    }},
    {"unreachable_ob_set_p", [] {
        Obj *o = new_object();
        PA_Picture p = PA_CreatePicture(nullptr, 0);
        ob_set_p((PA_ObjectRef)o, L"p", p); // consumes p: its PA_ClearVariable(&v) disposes the value
        CHECK(has_picture(o, "p"));
        release(o);
    }},
    {"unreachable_ob_get_a_released", [] {
        Obj *o = new_object();
        put_string(o, "k", "hello");
        CUTF8String s;
        CHECK(ob_get_a((PA_ObjectRef)o, L"k", &s));
        CHECK(s == CUTF8String((const uint8_t *)"hello"));
        release(o);
    }},
    {"unreachable_ob_get_b_n_nonmatching_released", [] {
        Obj *o = new_object();
        put_string(o, "b", "yes");
        put_string(o, "n", "1");
        CHECK(!ob_get_b((PA_ObjectRef)o, L"b"));
        CHECK(ob_get_n((PA_ObjectRef)o, L"n") == 0.0);
        release(o);
    }},
    {"unreachable_ob_get_c_borrowed", [] {
        Obj *o = new_object(), *c = new_object(true);
        put_ref(o, "c", c);
        PA_CollectionRef got = ob_get_c((PA_ObjectRef)o, L"c");
        CHECK(got == (PA_CollectionRef)c);
        CHECK(refcount(c) == 2); // test + parent: the parent keeps the borrowed ref alive
        release(c);
        release(o);
        CHECK(refcount(c) == 0);
    }},
    // --- controls: must pass before and after the fix ---
    {"control_ob_get_error_not_cleared", [] {
        Obj *o = new_object();
        put_string(o, "a", "x");
        put_bool(o, "b", true);
        put_real(o, "n", 1);
        CUTF8String s;
        fail_next_gets(3);
        ob_get_a((PA_ObjectRef)o, L"a", &s);
        CHECK(!ob_get_b((PA_ObjectRef)o, L"b"));
        CHECK(ob_get_n((PA_ObjectRef)o, L"n") == 0.0);
        release(o);
    }},
    {"control_consuming_setters", [] {
        Obj *o = new_object(), *child = new_object(), *named = new_object(), *coll = new_object(true);
        ob_set_n((PA_ObjectRef)o, L"n", 1);
        ob_set_b((PA_ObjectRef)o, L"t", true);
        ob_set_i((PA_ObjectRef)o, L"i", 3);
        ob_set_o((PA_ObjectRef)o, L"o", (PA_ObjectRef)child);   // consumes the caller's reference
        ob_set_o((PA_ObjectRef)o, "x", (PA_ObjectRef)named);    // consumes the caller's reference
        ob_set_c((PA_ObjectRef)o, L"c", (PA_CollectionRef)coll); // consumes the caller's reference
        CHECK(refcount(child) == 1 && refcount(named) == 1 && refcount(coll) == 1);
        CHECK(ob_is_defined((PA_ObjectRef)o, L"n"));
        release(o);
        CHECK(refcount(child) == 0 && refcount(named) == 0 && refcount(coll) == 0);
    }},
    {"control_ob_stringify", [] {
        Obj *o = new_object();
        put_string(o, "k", "v");
        CUTF8String s;
        ob_stringify((PA_ObjectRef)o, &s);
        CHECK(s == CUTF8String((const uint8_t *)"{\"k\":\"v\"}"));
        release(o);
    }},
    {"control_write_same_path_rejected", [] {
        std::string in = tmp_path("same.xlsx");
        make_workbook(in);
        Obj *refs = new_object();
        Call c;
        c.path(in).path(in).object(refs).text(CUTF16String());
        Obj *status = c.run(1);
        CHECK(status && !status_success(status));
        CHECK(copy_documents() == 0);
        release(status);
        release(refs);
    }},
};

int main(int argc, char **argv) {
    install();
    char tmpl[] = "/tmp/msoffice-ownership-XXXXXX";
    if (!mkdtemp(tmpl)) { perror("mkdtemp"); return 2; }
    gDir = tmpl;
    std::set<std::string> expected(argv + 1, argv + argc);
    int unexpected = 0, passed = 0, xfailed = 0;
    for (const Test &t : kTests) {
        reset();
        gFailures.clear();
        size_t strings = live_strings(), handles = live_handles(), objects = live_objects(),
               pictures = live_pictures();
        t.body();
        for (auto &v : violations()) gFailures.push_back("engine: " + v);
        if (live_strings() != strings)
            gFailures.push_back("leaked strings: " + std::to_string((long)(live_strings() - strings)));
        if (live_handles() != handles) gFailures.push_back("leaked handles");
        if (live_objects() != objects)
            gFailures.push_back("leaked objects: " + std::to_string((long)(live_objects() - objects)));
        if (live_pictures() != pictures)
            gFailures.push_back("leaked pictures: " + std::to_string((long)(live_pictures() - pictures)));
        bool failed = !gFailures.empty();
        bool want = expected.erase(t.name) > 0;
        const char *verdict = failed ? (want ? "XFAIL" : "FAIL") : (want ? "XPASS" : "PASS");
        printf("%-5s %s\n", verdict, t.name);
        for (auto &f : gFailures) printf("      %s\n", f.c_str());
        if (failed != want) unexpected++;
        else if (failed) xfailed++;
        else passed++;
    }
    std::error_code ec;
    std::filesystem::remove_all(gDir, ec);
    for (auto &name : expected) { printf("UNKNOWN expected failure %s\n", name.c_str()); unexpected++; }
    printf("passed=%d xfailed=%d unexpected=%d\n", passed, xfailed, unexpected);
    printf("%s\n", unexpected ? "RESULT: UNEXPECTED" : "RESULT: OK");
    return unexpected ? 1 : 0;
}
