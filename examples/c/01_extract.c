/* 01_extract: read a document as plain text, Markdown and IR JSON.
 *
 * Run: make -C examples/c build/01_extract && examples/c/build/01_extract */
#include "common.h"

int main(void) {
    printf("step 1: build a DOCX in memory\n");
    size_t len = 0;
    uint8_t *data = build_docx_bytes("Quarterly Report", &len);

    printf("step 2: open it from bytes\n");
    OfficeDocumentHandle *doc = open_bytes(data, len, "docx");
    free(data);
    const char *format = office_document_format(doc);
    printf("  format: %s\n", format);
    check(strcmp(format, "docx") == 0, "the opened format differs from the requested one");

    printf("step 3: extract plain text, Markdown and IR JSON\n");
    int err = 0;
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    char *markdown = take_str(office_document_to_markdown(doc, &err), err, "markdown");
    char *ir_json = take_str(office_document_to_ir_json(doc, &err), err, "ir json");
    printf("  chars: %zu\n", strlen(text));

    printf("step 4: check the output\n");
    has(text, "Revenue grew in every region.");
    has(markdown, "Quarterly Report");
    check(ir_json[0] == '{', "the IR JSON is not an object");

    office_oxide_free_string(text);
    office_oxide_free_string(markdown);
    office_oxide_free_string(ir_json);
    office_document_free(doc);
    printf("OK 01_extract\n");
    return 0;
}
