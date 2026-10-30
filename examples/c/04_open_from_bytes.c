/* 04_open_from_bytes: build a DOCX, an XLSX and a PPTX in memory and open each from its bytes.
 * Run: make -C examples/c run   (or build/04_open_from_bytes after make -C examples/c) */
#include "common.h"

int main(void) {
    size_t dlen, xlen, plen;

    printf("step 1: build a DOCX, an XLSX and a PPTX in memory\n");
    uint8_t *docx = build_docx_bytes("Quarterly Report", &dlen);
    uint8_t *xlsx = build_xlsx_bytes(&xlen);
    uint8_t *pptx = build_pptx_bytes(&plen);

    printf("step 2: open each from bytes\n");
    char *dt = text_of_bytes(docx, dlen, "docx");
    char *xt = text_of_bytes(xlsx, xlen, "xlsx");
    char *pt = text_of_bytes(pptx, plen, "pptx");
    printf("  format: docx\n  chars: %zu\n", strlen(dt));
    printf("  format: xlsx\n  chars: %zu\n", strlen(xt));
    printf("  format: pptx\n  chars: %zu\n", strlen(pt));

    printf("step 3: check the text\n");
    check_has(dt, "Revenue grew in every region.", "the DOCX has its paragraph");
    check_has(xt, "North", "the XLSX has its first data row");
    check_has(pt, "Welcome", "the PPTX has its first title");

    printf("step 4: check that a wrong format is an error\n");
    int err = 0;
    OfficeDocumentHandle *bad = office_document_open_from_bytes(xlsx, xlen, "pptx", &err);
    check(bad == NULL && err != OFFICE_OK, "an XLSX opened as a PPTX is refused");
    printf("  error: %d\n", err);

    free(dt);
    free(xt);
    free(pt);
    free(docx);
    office_oxide_free_bytes(xlsx, xlen);
    office_oxide_free_bytes(pptx, plen);
    printf("OK 04_open_from_bytes\n");
    return 0;
}
