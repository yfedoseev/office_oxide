/* 10_detect_and_batch: detect the format of bytes, then process a mixed batch.
 * Run: make -C examples/c run   (or build/10_detect_and_batch after make -C examples/c)
 * The C ABI detects a format from a path only, so the bytes are sniffed here. */
#include "common.h"

/* All three formats are ZIP archives; the part names inside tell them apart. */
static const char *detect(const uint8_t *d, size_t n) {
    static const struct { const char *marker, *format; } kinds[] = {
        {"word/", "docx"}, {"xl/", "xlsx"}, {"ppt/", "pptx"}};
    for (size_t k = 0; k < 3; k++) {
        size_t m = strlen(kinds[k].marker);
        for (size_t i = 0; i + m <= n; i++)
            if (memcmp(d + i, kinds[k].marker, m) == 0) return kinds[k].format;
    }
    return NULL;
}

int main(void) {
    printf("step 1: build a DOCX, an XLSX and a PPTX in memory\n");
    size_t len[3];
    uint8_t *batch[3] = {build_docx_bytes("Quarterly Report", &len[0]), build_xlsx_bytes(&len[1]),
                         build_pptx_bytes(&len[2])};
    const char *expect[3] = {"docx", "xlsx", "pptx"};

    printf("step 2: detect each format and process the batch\n");
    for (int i = 0; i < 3; i++) {
        const char *format = detect(batch[i], len[i]);
        check(format && strcmp(format, expect[i]) == 0, "the detected format matches the one built");
        char *text = text_of_bytes(batch[i], len[i], format);
        printf("  report.%s: format %s, chars %zu\n", expect[i], format, strlen(text));
        check(strlen(text) > 0, "every document has text");
        free(text);
    }

    printf("step 3: check that unknown bytes are not detected\n");
    check(detect((const uint8_t *)"not an office file", 18) == NULL, "garbage has no format");

    free(batch[0]);
    office_oxide_free_bytes(batch[1], len[1]);
    office_oxide_free_bytes(batch[2], len[2]);
    printf("OK 10_detect_and_batch\n");
    return 0;
}
