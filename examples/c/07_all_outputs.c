/* 07_all_outputs: render one document as plain text, Markdown, HTML and IR JSON.
 * Run: make -C examples/c run   (or build/07_all_outputs after make -C examples/c) */
#include "common.h"

int main(void) {
    printf("step 1: build a DOCX in memory\n");
    size_t len = 0;
    uint8_t *bytes = build_docx_bytes("Quarterly Report", &len);

    printf("step 2: open it from bytes\n");
    OfficeDocumentHandle *doc = open_bytes(bytes, len, "docx");
    free(bytes);

    printf("step 3: render every output\n");
    int err = 0;
    /* Plain text: words only, no markup. */
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    /* Markdown: headings and tables survive as syntax. */
    char *md = take_str(office_document_to_markdown(doc, &err), err, "markdown");
    /* HTML: a fragment (no <html> wrapper) to embed in a page. */
    char *html = take_str(office_document_to_html(doc, &err), err, "html");
    /* IR JSON: the format-neutral tree, for code that wants structure. */
    char *ir = take_str(office_document_to_ir_json(doc, &err), err, "ir json");
    printf("  chars: text %zu, markdown %zu, html %zu, ir %zu\n",
           strlen(text), strlen(md), strlen(html), strlen(ir));

    printf("step 4: check each output\n");
    check(strlen(text) && strlen(md) && strlen(html) && strlen(ir), "every output is non-empty");
    check_has(text, "Quarterly Report", "plain text has the title");
    check_has(md, "# Quarterly Report", "markdown has the heading");
    check_has(html, "<", "html has markup");
    check_has(html, "Quarterly Report", "html has the title");
    /* The C ABI has no JSON parser, so check the shape: one object. */
    size_t n = strlen(ir);
    while (n && (ir[n - 1] == '\n' || ir[n - 1] == ' ')) n--;
    check(ir[0] == '{' && ir[n - 1] == '}', "IR JSON is an object");
    check_has(ir, "Quarterly Report", "IR JSON has the title");

    office_oxide_free_string(text);
    office_oxide_free_string(md);
    office_oxide_free_string(html);
    office_oxide_free_string(ir);
    office_document_free(doc);
    printf("OK 07_all_outputs\n");
    return 0;
}
