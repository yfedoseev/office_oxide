/* 07_all_outputs: render one document as plain text, Markdown, HTML and IR JSON.
 *
 * Run: make -C examples/c build/07_all_outputs && examples/c/build/07_all_outputs */
#include "common.h"

int main(void) {
    printf("step 1: build a DOCX in memory\n");
    size_t len = 0;
    uint8_t *data = build_docx_bytes("Quarterly Report", &len);

    printf("step 2: open it from bytes\n");
    OfficeDocumentHandle *doc = open_bytes(data, len, "docx");
    free(data);
    printf("  format: %s\n", office_document_format(doc));

    printf("step 3: render every output\n");
    int err = 0;
    char *plain = take_str(office_document_plain_text(doc, &err), err, "plain text"); /* words only: no markup, no structure */
    char *markdown = take_str(office_document_to_markdown(doc, &err), err, "markdown"); /* headings and pipe tables survive as syntax */
    char *html = take_str(office_document_to_html(doc, &err), err, "html"); /* a fragment (no <html> wrapper), tables as <table> */
    char *ir_json = take_str(office_document_to_ir_json(doc, &err), err, "ir json"); /* the format-neutral tree as JSON, for programs */
    const char *names[] = {"plain", "markdown", "html", "ir_json"};
    const char *values[] = {plain, markdown, html, ir_json};
    for (int i = 0; i < 4; i++) printf("  chars: %s %zu\n", names[i], strlen(values[i]));

    printf("step 4: check each output\n");
    for (int i = 0; i < 4; i++) {
        char message[64];
        snprintf(message, sizeof message, "the %s output is empty", names[i]);
        check(strlen(values[i]) > 0, message);
    }
    has(plain, "Quarterly Report");
    check(strstr(markdown, "# Quarterly Report") != NULL, "the Markdown is missing the heading");
    check(strstr(html, "<table") != NULL, "the HTML has no table");
    /* The C ABI has no JSON parser, so check the shape: one object. */
    size_t n = strlen(ir_json);
    while (n && (ir_json[n - 1] == '\n' || ir_json[n - 1] == ' ')) n--;
    check(ir_json[0] == '{' && ir_json[n - 1] == '}', "the IR JSON does not parse to an object");

    office_oxide_free_string(plain);
    office_oxide_free_string(markdown);
    office_oxide_free_string(html);
    office_oxide_free_string(ir_json);
    office_document_free(doc);
    printf("OK 07_all_outputs\n");
    return 0;
}
