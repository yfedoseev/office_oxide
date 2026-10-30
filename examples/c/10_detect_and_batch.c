/* 10_detect_and_batch: detect the format of bytes, then process a mixed batch.
 *
 * Run: make -C examples/c build/10_detect_and_batch && examples/c/build/10_detect_and_batch */
#include "common.h"

/* The C ABI detects a format from a path only. A ZIP stores each part name in
 * the clear, and the main part's name tells the three formats apart. */
static const char *detect_format(const uint8_t *data, size_t len) {
    static const struct { const char *part, *format; } kinds[] = {
        {"word/document.xml", "docx"}, {"xl/workbook.xml", "xlsx"}, {"ppt/presentation.xml", "pptx"}};
    if (len < 2 || data[0] != 'P' || data[1] != 'K') return NULL;
    for (size_t k = 0; k < 3; k++)
        if (find_bytes(data, len, kinds[k].part) >= 0) return kinds[k].format;
    return NULL;
}

int main(void) {
    printf("step 1: build a DOCX, an XLSX and a PPTX in memory\n");
    const char *names[] = {"report.docx", "report.xlsx", "report.pptx"};
    size_t len[3];
    uint8_t *batch[3] = {build_docx_bytes("Quarterly Report", &len[0]), build_xlsx_bytes(&len[1]), build_pptx_bytes(&len[2])};

    printf("step 2: detect each format and open it from bytes\n");
    for (int i = 0; i < 3; i++) {
        const char *format = detect_format(batch[i], len[i]);
        check(format != NULL && strcmp(format, names[i] + strlen("report.")) == 0, "the detected format differs from the one built");
        char *text = text_of_bytes(batch[i], len[i], format);
        check(strlen(text) > 0, "a document has no text");
        printf("  %s: format %s, chars %zu\n", names[i], format, strlen(text));
        free(text);
    }

    printf("step 3: check that bytes that are not a document are not detected\n");
    check(detect_format((const uint8_t *)"not an office file", 18) == NULL, "bytes that are not a document were detected as one");

    free(batch[0]);
    office_oxide_free_bytes(batch[1], len[1]);
    office_oxide_free_bytes(batch[2], len[2]);
    printf("OK 10_detect_and_batch\n");
    return 0;
}
