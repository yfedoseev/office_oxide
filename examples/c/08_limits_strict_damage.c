/* 08_limits_strict_damage: strict options refuse a truncated archive; a lenient open recovers or fails cleanly.
 * Run: make -C examples/c run   (or build/08_limits_strict_damage after make -C examples/c) */
#include "common.h"

int main(void) {
    printf("step 1: build a DOCX in memory\n");
    size_t len = 0;
    uint8_t *good = build_docx_bytes("Quarterly Report", &len);

    printf("step 2: keep only the first half of its bytes\n");
    size_t cut = len / 2;

    printf("step 3: open the damaged bytes with strict options\n");
    OfficeParseOptionsHandle *strict = office_options_strict();
    OfficeParseOptionsHandle *lenient = office_options_new();
    check(strict != NULL && lenient != NULL, "options handles");
    check(office_options_set(strict, "text_chars", 1000000) == OFFICE_OK, "a known option is accepted");
    check(office_options_set(strict, "no_such_option", 1) == OFFICE_ERR_INVALID_ARG, "an unknown option is refused");
    int err = 0;
    OfficeDocumentHandle *doc = office_document_open_from_bytes_with(good, cut, "docx", strict, &err);
    check(doc == NULL && err != OFFICE_OK, "strict mode refuses a truncated archive");
    printf("  error: %d\n", err);

    printf("step 4: open the damaged bytes leniently\n");
    err = 0;
    doc = office_document_open_from_bytes_with(good, cut, "docx", lenient, &err);
    if (doc) {
        char *damage = take_str(office_document_damage_json(doc, &err), err, "damage json");
        printf("  damage: %s\n", damage);
        check(strcmp(damage, "[]") != 0, "a half archive is not reported as sound");
        office_oxide_free_string(damage);
        office_document_free(doc);
    } else {
        check(err != OFFICE_OK, "a failed open carries an error code");
        printf("  error: %d\n", err);
    }

    printf("step 5: open the intact bytes with strict options\n");
    doc = office_document_open_from_bytes_with(good, len, "docx", strict, &err);
    check_ok(doc ? OFFICE_OK : OFFICE_ERR_PARSE, err, "strict open of an intact file");
    char *damage = take_str(office_document_damage_json(doc, &err), err, "damage json");
    printf("  damage: %s\n", damage);
    check(strcmp(damage, "[]") == 0, "an intact file reports no damage");
    office_oxide_free_string(damage);
    office_document_free(doc);

    office_options_free(strict);
    office_options_free(lenient);
    free(good);
    printf("OK 08_limits_strict_damage\n");
    return 0;
}
