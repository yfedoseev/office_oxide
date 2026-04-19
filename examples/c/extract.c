/* Extract plain text and Markdown using the raw office_oxide C FFI.
 *
 * Build (from the monorepo root):
 *   cargo build --release --lib
 *   gcc -I include/office_oxide_c examples/c/extract.c \
 *       -L target/release -loffice_oxide \
 *       -Wl,-rpath,$(pwd)/target/release \
 *       -o /tmp/extract
 *
 * Run:
 *   /tmp/extract report.docx
 */

#include <stdio.h>
#include <stdlib.h>
#include "office_oxide.h"

int main(int argc, char** argv) {
    if (argc != 2) {
        fprintf(stderr, "usage: %s <file>\n", argv[0]);
        return 1;
    }
    printf("office_oxide version: %s\n", office_oxide_version());

    int err = 0;
    OfficeDocumentHandle* doc = office_document_open(argv[1], &err);
    if (!doc) {
        fprintf(stderr, "open failed, code=%d\n", err);
        return 1;
    }
    printf("format: %s\n", office_document_format(doc));

    char* text = office_document_plain_text(doc, &err);
    if (text) {
        printf("--- plain text ---\n%s\n", text);
        office_oxide_free_string(text);
    }
    char* md = office_document_to_markdown(doc, &err);
    if (md) {
        printf("--- markdown ---\n%s\n", md);
        office_oxide_free_string(md);
    }
    office_document_free(doc);
    return 0;
}
