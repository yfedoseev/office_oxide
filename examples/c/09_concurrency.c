/* 09_concurrency: open and read many documents from several threads at once.
 * Run: make -C examples/c run   (or build/09_concurrency after make -C examples/c)
 * The C ABI has no logging bridge, so this proves only that parallel parsing is safe. */
#include <pthread.h>
#include "common.h"

#define WORKERS 8
#define DOCUMENTS 32

static uint8_t *docs[DOCUMENTS];
static size_t doc_len[DOCUMENTS];
static int failed[WORKERS];

/* Worker w handles documents w, w+8, w+16, ... Each document has its own
 * handle, which the ABI allows to be used from any one thread at a time. */
static void *worker(void *arg) {
    int w = (int)(intptr_t)arg;
    for (int i = w; i < DOCUMENTS; i += WORKERS) {
        int err = 0;
        OfficeDocumentHandle *doc = office_document_open_from_bytes(docs[i], doc_len[i], "docx", &err);
        if (!doc) { failed[w] = 1; continue; }
        char *text = office_document_plain_text(doc, &err);
        char want[64];
        snprintf(want, sizeof want, "Quarterly Report %d", i);
        if (!text || !strstr(text, want)) failed[w] = 1;
        if (text) office_oxide_free_string(text);
        office_document_free(doc);
    }
    return NULL;
}

int main(void) {
    printf("step 1: build 32 DOCX files in memory\n");
    for (int i = 0; i < DOCUMENTS; i++) {
        char title[64];
        snprintf(title, sizeof title, "Quarterly Report %d", i);
        docs[i] = build_docx_bytes(title, &doc_len[i]);
    }

    printf("step 2: open and read them from 8 workers\n");
    pthread_t threads[WORKERS];
    for (intptr_t w = 0; w < WORKERS; w++)
        check(pthread_create(&threads[w], NULL, worker, (void *)w) == 0, "start a worker");
    for (int w = 0; w < WORKERS; w++)
        check(pthread_join(threads[w], NULL) == 0, "join a worker");
    printf("  workers: %d\n  documents: %d\n", WORKERS, DOCUMENTS);

    printf("step 3: check every result\n");
    for (int w = 0; w < WORKERS; w++)
        check(!failed[w], "every worker read its own documents back");

    for (int i = 0; i < DOCUMENTS; i++) free(docs[i]);
    printf("OK 09_concurrency\n");
    return 0;
}
