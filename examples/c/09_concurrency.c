/* 09_concurrency: open and read 32 documents from 8 workers at once, with logging on.
 *
 * Run: make -C examples/c build/09_concurrency && examples/c/build/09_concurrency */
#include <pthread.h>
#include "common.h"

#define WORKERS 8
#define DOCUMENTS 32

static uint8_t *docs[DOCUMENTS];
static size_t doc_len[DOCUMENTS];
static int read_ok[DOCUMENTS];   /* 1 once the document was read */
static int text_ok[DOCUMENTS];   /* 1 when it read back its own title */

/* Worker w handles documents w, w+8, w+16, ... Each document has its own
 * handle, which the ABI allows to be used from any one thread at a time. */
static void *worker(void *arg) {
    int w = (int)(intptr_t)arg;
    for (int i = w; i < DOCUMENTS; i += WORKERS) {
        int err = 0;
        OfficeDocumentHandle *doc = office_document_open_from_bytes(docs[i], doc_len[i], "docx", &err);
        if (!doc) continue;
        char *text = office_document_plain_text(doc, &err);
        char want[64];
        snprintf(want, sizeof want, "Quarterly Report %d", i);
        if (text) {
            read_ok[i] = 1;
            text_ok[i] = strstr(text, want) != NULL;
            office_oxide_free_string(text);
        }
        office_document_free(doc);
    }
    return NULL;
}

/* There is no host logger to switch on in the C ABI, and `make run` wraps the
 * run in `timeout`, which turns a deadlock into a failure. */
int main(void) {
    printf("step 1: build 32 DOCX files in memory\n");
    for (int i = 0; i < DOCUMENTS; i++) {
        char title[64];
        snprintf(title, sizeof title, "Quarterly Report %d", i);
        docs[i] = build_docx_bytes(title, &doc_len[i]);
    }
    printf("  documents: %d\n", DOCUMENTS);

    printf("step 2: skipped, the C ABI has no logging bridge\n");

    printf("step 3: parse them on 8 workers\n");
    pthread_t threads[WORKERS];
    for (intptr_t w = 0; w < WORKERS; w++)
        check(pthread_create(&threads[w], NULL, worker, (void *)w) == 0, "start a worker");
    for (int w = 0; w < WORKERS; w++)
        check(pthread_join(threads[w], NULL) == 0, "join a worker");
    printf("  workers: %d\n", WORKERS);

    printf("step 4: check every result\n");
    for (int i = 0; i < DOCUMENTS; i++) check(read_ok[i], "not every document was read");
    for (int i = 0; i < DOCUMENTS; i++) {
        char message[64];
        snprintf(message, sizeof message, "document %d came back with the wrong text", i);
        check(text_ok[i], message);
    }

    for (int i = 0; i < DOCUMENTS; i++) free(docs[i]);
    printf("OK 09_concurrency\n");
    return 0;
}
