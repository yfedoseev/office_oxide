/* 02_create_from_markdown: turn one Markdown source into a DOCX, an XLSX and a PPTX.
 *
 * Run: make -C examples/c build/02_create_from_markdown && examples/c/build/02_create_from_markdown */
#include "common.h"

int main(void) {
    printf("step 1: write the Markdown source\n");
    char markdown[512];
    report_markdown(markdown, sizeof markdown, "Quarterly Report");

    printf("step 2: create a DOCX, an XLSX and a PPTX from it\n");
    const char *formats[] = {"docx", "xlsx", "pptx"};
    uint8_t *created[3];
    size_t lens[3];
    for (int i = 0; i < 3; i++) created[i] = markdown_to_bytes(markdown, formats[i], &lens[i]);

    printf("step 3: open each and check the text\n");
    for (int i = 0; i < 3; i++) {
        char *text = text_of_bytes(created[i], lens[i], formats[i]);
        printf("  format: %s\n  chars: %zu\n", formats[i], strlen(text));
        has(text, "North");
        free(text);
        free(created[i]);
    }

    printf("OK 02_create_from_markdown\n");
    return 0;
}
