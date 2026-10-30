/* Shared helpers for the C examples: the build_*_bytes helpers every language
 * has, plus `check` and `has`. Every library call is checked; a failure prints
 * what failed and exits non-zero. */
#ifndef OO_EXAMPLE_COMMON_H
#define OO_EXAMPLE_COMMON_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>
#include "office_oxide.h"

static inline void check(int ok, const char *msg) {
    if (!ok) {
        fprintf(stderr, "FAIL: %s\n", msg);
        exit(1);
    }
}

/* A library status and error code must both be OFFICE_OK. */
static inline void check_ok(int rc, int err, const char *op) {
    if (rc != OFFICE_OK || err != OFFICE_OK) {
        fprintf(stderr, "FAIL: %s (status %d, error %d)\n", op, rc, err);
        exit(1);
    }
}

/* Assert that `hay` contains `needle`. */
static inline void has(const char *hay, const char *needle) {
    if (!strstr(hay, needle)) {
        fprintf(stderr, "FAIL: \"%s\" is missing from the text\n", needle);
        exit(1);
    }
}

/* The error codes as short text: the C ABI reports a code, not a message. */
static inline const char *error_text(int err) {
    switch (err) {
    case OFFICE_ERR_INVALID_ARG: return "invalid argument";
    case OFFICE_ERR_IO: return "I/O error";
    case OFFICE_ERR_PARSE: return "parse error";
    case OFFICE_ERR_EXTRACTION: return "extraction error";
    case OFFICE_ERR_UNSUPPORTED: return "unsupported";
    case OFFICE_ERR_INCOMPLETE: return "incomplete document";
    default: return "internal error";
    }
}

/* Index of the first occurrence of `needle` in `hay`, or -1. */
static inline long find_bytes(const uint8_t *hay, size_t len, const char *needle) {
    size_t n = strlen(needle);
    for (size_t i = 0; i + n <= len; i++)
        if (memcmp(hay + i, needle, n) == 0) return (long)i;
    return -1;
}

static inline char *take_str(char *s, int err, const char *op) {
    check_ok(s ? OFFICE_OK : OFFICE_ERR_INTERNAL, err, op);
    return s; /* free with office_oxide_free_string */
}

static inline OfficeDocumentHandle *open_bytes(const uint8_t *data, size_t len, const char *format) {
    int err = 0;
    OfficeDocumentHandle *doc = office_document_open_from_bytes(data, len, format, &err);
    check_ok(doc ? OFFICE_OK : OFFICE_ERR_PARSE, err, "open from bytes");
    return doc;
}

/* Plain text of a document held in memory (malloc'd; free with free()). */
static inline char *text_of_bytes(const uint8_t *data, size_t len, const char *format) {
    int err = 0;
    OfficeDocumentHandle *doc = open_bytes(data, len, format);
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    char *copy = strdup(text);
    office_oxide_free_string(text);
    office_document_free(doc);
    check(copy != NULL, "out of memory");
    return copy;
}

/* The Markdown source shared by every language's examples, with the title
 * replaceable (09 numbers its documents). */
static inline void report_markdown(char *out, size_t cap, const char *title) {
    snprintf(out, cap,
             "# %s\n\nRevenue grew in every region.\n\n"
             "| Region | Q1  | Q2  |\n|--------|-----|-----|\n"
             "| North  | 120 | 135 |\n| South  | 95  | 110 |\n",
             title);
}

/* The C ABI creates a document from Markdown by writing to a path. This uses a
 * private temp file for that hop and removes it before returning. Returns
 * malloc'd bytes (free with free()). */
static inline uint8_t *markdown_to_bytes(const char *markdown, const char *format, size_t *len) {
    char path[] = "/tmp/oo_example_XXXXXX";
    int fd = mkstemp(path);
    check(fd >= 0, "temp file");
    int err = 0;
    int rc = office_create_from_markdown(markdown, format, path, &err);
    check_ok(rc, err, "create from markdown");
    FILE *f = fdopen(fd, "rb");
    check(f != NULL, "fdopen");
    fseek(f, 0, SEEK_END);
    long n = ftell(f);
    check(n > 0, "the created file is empty");
    rewind(f);
    uint8_t *buf = malloc((size_t)n);
    check(buf && fread(buf, 1, (size_t)n, f) == (size_t)n, "read the created file back");
    fclose(f);
    unlink(path);
    *len = (size_t)n;
    return buf;
}

static inline uint8_t *build_docx_bytes(const char *title, size_t *len) {
    char md[512];
    report_markdown(md, sizeof md, title);
    return markdown_to_bytes(md, "docx", len);
}

/* Sheet "Sales": headers, two data rows, a SUM row, a boolean, a merged note,
 * a wide column A, and a bold header row. Free with office_oxide_free_bytes. */
static inline uint8_t *build_xlsx_bytes(size_t *len) {
    int err = 0;
    OfficeXlsxWriterHandle *w = office_xlsx_writer_new();
    check(w != NULL, "xlsx writer");
    uint32_t s = office_xlsx_writer_add_sheet(w, "Sales");
    check(s != UINT32_MAX, "add sheet");
    const char *head[] = {"Region", "Q1", "Q2", "Growth"};
    for (uint32_t c = 0; c < 4; c++) /* bold header row */
        check_ok(office_xlsx_sheet_set_cell_styled(w, s, 0, c, OFFICE_CELL_STRING, head[c], 0, true, NULL), 0, "header cell");
    check_ok(office_xlsx_sheet_set_cell(w, s, 1, 0, OFFICE_CELL_STRING, "North", 0), 0, "cell A2");
    check_ok(office_xlsx_sheet_set_cell(w, s, 1, 1, OFFICE_CELL_NUMBER, NULL, 120), 0, "cell B2");
    check_ok(office_xlsx_sheet_set_cell(w, s, 1, 2, OFFICE_CELL_NUMBER, NULL, 135), 0, "cell C2");
    check_ok(office_xlsx_sheet_set_cell(w, s, 1, 3, OFFICE_CELL_BOOLEAN, NULL, 1), 0, "cell D2");
    check_ok(office_xlsx_sheet_set_cell(w, s, 2, 0, OFFICE_CELL_STRING, "South", 0), 0, "cell A3");
    check_ok(office_xlsx_sheet_set_cell(w, s, 2, 1, OFFICE_CELL_NUMBER, NULL, 95), 0, "cell B3");
    check_ok(office_xlsx_sheet_set_cell(w, s, 2, 2, OFFICE_CELL_NUMBER, NULL, 110), 0, "cell C3");
    check_ok(office_xlsx_sheet_set_cell(w, s, 3, 0, OFFICE_CELL_STRING, "Total", 0), 0, "cell A4");
    check_ok(office_xlsx_sheet_set_cell(w, s, 3, 1, OFFICE_CELL_FORMULA, "SUM(B2:B3)", 0), 0, "cell B4");
    check_ok(office_xlsx_sheet_set_cell(w, s, 3, 2, OFFICE_CELL_FORMULA, "SUM(C2:C3)", 0), 0, "cell C4");
    check_ok(office_xlsx_sheet_set_cell(w, s, 5, 0, OFFICE_CELL_STRING, "Notes", 0), 0, "cell A6");
    office_xlsx_sheet_merge_cells(w, s, 5, 0, 1, 3); /* A6:C6 */
    office_xlsx_sheet_set_column_width(w, s, 0, 18.0);
    uint8_t *bytes = office_xlsx_writer_to_bytes(w, len, &err);
    check_ok(bytes ? OFFICE_OK : OFFICE_ERR_INTERNAL, err, "xlsx to bytes");
    office_xlsx_writer_free(w);
    return bytes;
}

/* The 1x1 PNG from the catalogue (its base64 form begins iVBORw0KGgo). */
static const uint8_t PNG_1X1[] = {
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0xDA, 0x63, 0x64, 0x60, 0xF8, 0x5F,
    0x0F, 0x00, 0x02, 0x87, 0x01, 0x80, 0xEB, 0x47, 0xBA, 0x92, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
    0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82};

/* Two slides: "Welcome" and "Results" (with the PNG). Free with
 * office_oxide_free_bytes. */
static inline uint8_t *build_pptx_bytes(size_t *len) {
    int err = 0;
    OfficePptxWriterHandle *w = office_pptx_writer_new();
    check(w != NULL, "pptx writer");
    office_pptx_writer_set_presentation_size(w, 12192000, 6858000); /* 16:9, in EMU (914400 per inch) */
    uint32_t one = office_pptx_writer_add_slide(w);
    check(one != UINT32_MAX, "add slide 1");
    check_ok(office_pptx_slide_set_title(w, one, "Welcome"), 0, "title 1");
    check_ok(office_pptx_slide_add_text(w, one, "Quarterly Report"), 0, "text 1");
    uint32_t two = office_pptx_writer_add_slide(w);
    check(two != UINT32_MAX, "add slide 2");
    check_ok(office_pptx_slide_set_title(w, two, "Results"), 0, "title 2");
    check_ok(office_pptx_slide_add_text(w, two, "Revenue grew in every region."), 0, "text 2");
    check_ok(office_pptx_slide_add_image(w, two, PNG_1X1, sizeof PNG_1X1, "png", 914400, 2743200, 1828800, 1828800), 0, "image");
    uint8_t *bytes = office_pptx_writer_to_bytes(w, len, &err);
    check_ok(bytes ? OFFICE_OK : OFFICE_ERR_INTERNAL, err, "pptx to bytes");
    office_pptx_writer_free(w);
    return bytes;
}

/* The IR JSON holds one object per sheet or slide inside "sections":[...].
 * The C ABI has no JSON parser, so count them: objects at depth zero of that
 * array, skipping over strings. */
static inline int count_sections(const char *ir) {
    const char *p = strstr(ir, "\"sections\":[");
    check(p != NULL, "the IR JSON has no sections");
    p += strlen("\"sections\":[");
    int depth = 0, count = 0, in_string = 0;
    for (; *p; p++) {
        if (in_string) {
            if (*p == '\\') p++;
            else if (*p == '"') in_string = 0;
        } else if (*p == '"') {
            in_string = 1;
        } else if (*p == '{' || *p == '[') {
            if (depth++ == 0 && *p == '{') count++;
        } else if (*p == '}' || *p == ']') {
            if (depth == 0) break;
            depth--;
        }
    }
    return count;
}

#endif
