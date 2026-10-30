/* 01_extract: read a document as plain text, Markdown and IR JSON.
 * Run: make -C examples/c run   (or build/01_extract after make -C examples/c) */
#include "common.h"

int main(void) {
    printf("step 1: build a DOCX in memory\n");
    size_t len = 0;
    uint8_t *bytes = build_docx_bytes("Quarterly Report", &len);

    printf("step 2: open it from bytes\n");
    OfficeDocumentHandle *doc = open_bytes(bytes, len, "docx");
    free(bytes);
    printf("  format: %s\n", office_document_format(doc));
    check(strcmp(office_document_format(doc), "docx") == 0, "format is docx");

    printf("step 3: extract plain text, Markdown and IR JSON\n");
    int err = 0;
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    char *md = take_str(office_document_to_markdown(doc, &err), err, "markdown");
    char *ir = take_str(office_document_to_ir_json(doc, &err), err, "ir json");
    printf("  chars: %zu\n", strlen(text));

    printf("step 4: check the output\n");
    check_has(text, "Revenue grew in every region.", "plain text has the paragraph");
    check_has(md, "Quarterly Report", "markdown has the title");
    check(ir[0] == '{', "IR JSON is an object");

    office_oxide_free_string(text);
    office_oxide_free_string(md);
    office_oxide_free_string(ir);
    office_document_free(doc);
    printf("OK 01_extract\n");
    return 0;
}
