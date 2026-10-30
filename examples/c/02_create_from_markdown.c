/* 02_create_from_markdown: turn one Markdown source into a DOCX, an XLSX and a PPTX.
 * Run: make -C examples/c run   (or build/02_create_from_markdown after make -C examples/c) */
#include "common.h"

int main(void) {
    printf("step 1: write the Markdown source\n");
    char md[512];
    report_markdown(md, sizeof md, "Quarterly Report");

    printf("step 2: create a DOCX, an XLSX and a PPTX from it\n");
    const char *formats[] = {"docx", "xlsx", "pptx"};
    uint8_t *bufs[3];
    long lens[3];
    for (int i = 0; i < 3; i++) {
        char path[] = "/tmp/oo_example_XXXXXX";
        int fd = mkstemp(path);
        check(fd >= 0, "temp file");
        close(fd);
        int err = 0;
        int rc = office_create_from_markdown(md, formats[i], path, &err);
        check_ok(rc, err, "create from markdown");

        FILE *f = fopen(path, "rb");
        check(f != NULL, "open the created file");
        fseek(f, 0, SEEK_END);
        long n = ftell(f);
        rewind(f);
        uint8_t *buf = malloc((size_t)n);
        check(buf && fread(buf, 1, (size_t)n, f) == (size_t)n, "read the created file");
        fclose(f);
        unlink(path);
        bufs[i] = buf;
        lens[i] = n;
    }

    printf("step 3: open each and check the text\n");
    for (int i = 0; i < 3; i++) {
        char *text = text_of_bytes(bufs[i], (size_t)lens[i], formats[i]);
        printf("  format: %s\n  chars: %zu\n", formats[i], strlen(text));
        check_has(text, "North", "the table survived");
        free(text);
        free(bufs[i]);
    }
    printf("OK 02_create_from_markdown\n");
    return 0;
}
