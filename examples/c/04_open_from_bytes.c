/* 04_open_from_bytes: build a DOCX, an XLSX and a PPTX in memory and open each from its bytes.
 *
 * Run: make -C examples/c build/04_open_from_bytes && examples/c/build/04_open_from_bytes */
#include "common.h"

int main(void) {
    printf("step 1: build a DOCX, an XLSX and a PPTX in memory\n");
    const char *formats[] = {"docx", "xlsx", "pptx"};
    size_t len[3];
    uint8_t *built[3] = {build_docx_bytes("Quarterly Report", &len[0]), build_xlsx_bytes(&len[1]), build_pptx_bytes(&len[2])};

    printf("step 2: open each from bytes\n");
    char *texts[3];
    for (int i = 0; i < 3; i++) {
        OfficeDocumentHandle *doc = open_bytes(built[i], len[i], formats[i]);
        check(strcmp(office_document_format(doc), formats[i]) == 0, "the opened format differs from the requested one");
        office_document_free(doc);
        texts[i] = text_of_bytes(built[i], len[i], formats[i]);
        printf("  format: %s\n  chars: %zu\n", formats[i], strlen(texts[i]));
    }

    printf("step 3: check the text\n");
    has(texts[0], "Revenue grew in every region.");
    has(texts[1], "North");
    has(texts[1], "Total");
    has(texts[2], "Welcome");
    has(texts[2], "Results");

    printf("step 4: check that a wrong format is an error\n");
    int err = 0;
    OfficeDocumentHandle *wrong = office_document_open_from_bytes(built[1], len[1], "pptx", &err);
    check(wrong == NULL && err != OFFICE_OK, "an XLSX opened as a PPTX was accepted");
    printf("  error: %s\n", error_text(err));

    for (int i = 0; i < 3; i++) free(texts[i]);
    free(built[0]);
    office_oxide_free_bytes(built[1], len[1]);
    office_oxide_free_bytes(built[2], len[2]);
    printf("OK 04_open_from_bytes\n");
    return 0;
}
