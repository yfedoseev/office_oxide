/* 08_limits_strict_damage: strict mode refuses a damaged archive; a lenient open reports the damage.
 *
 * Run: make -C examples/c build/08_limits_strict_damage && examples/c/build/08_limits_strict_damage */
#include "common.h"

#define PART "word/_rels/document.xml.rels"

/* Overwrites the start of one part's compressed data. A first deflate byte of
 * 0xFF declares a reserved block type, so every inflater rejects the part, and
 * the central directory stays intact. The part is found by name, not by
 * offset, so the recipe does not depend on how big the other parts are.
 * Returns a malloc'd copy. */
static uint8_t *damage_part(const uint8_t *data, size_t len) {
    long at = find_bytes(data, len, PART); /* the first hit is the part's local header */
    check(at > 0, "the part to damage is missing from the archive");
    /* The extra-field length sits in the two bytes just before the name. */
    size_t start = (size_t)at + strlen(PART) + (data[at - 2] | (data[at - 1] << 8));
    check(start + 16 <= len, "the part to damage is missing from the archive");
    uint8_t *damaged = malloc(len);
    check(damaged != NULL, "out of memory");
    memcpy(damaged, data, len);
    memset(damaged + start, 0xFF, 16);
    return damaged;
}

/* The kinds in a damage report, joined by ", ". The C ABI has no JSON parser,
 * so scan for each "kind":"..." member. */
static void damage_kinds(const char *json, char *out, size_t cap) {
    out[0] = 0;
    for (const char *p = json; (p = strstr(p, "\"kind\":\"")) != NULL;) {
        p += strlen("\"kind\":\"");
        const char *end = strchr(p, '"');
        check(end != NULL, "the damage report is malformed");
        if (out[0]) strncat(out, ", ", cap - strlen(out) - 1);
        strncat(out, p, (size_t)(end - p));
        p = end;
    }
}

/* Opens bytes as a DOCX under `options`: the error code, or OFFICE_OK when it opened. */
static int open_error(const uint8_t *data, size_t len, const OfficeParseOptionsHandle *options) {
    int err = 0;
    OfficeDocumentHandle *doc = office_document_open_from_bytes_with(data, len, "docx", options, &err);
    if (doc) office_document_free(doc);
    return doc ? OFFICE_OK : err;
}

int main(void) {
    printf("step 1: build a DOCX in memory\n");
    size_t len = 0;
    uint8_t *good = build_docx_bytes("Quarterly Report", &len);

    printf("step 2: damage one part: overwrite 16 bytes of " PART "\n");
    uint8_t *damaged = damage_part(good, len);

    OfficeParseOptionsHandle *strict = office_options_strict();
    OfficeParseOptionsHandle *lenient = office_options_new();
    check(strict != NULL && lenient != NULL, "options handles");

    printf("step 3: open it in strict mode\n");
    int err = open_error(damaged, len, strict);
    check(err != OFFICE_OK, "strict mode accepted a damaged part");
    printf("  error: %s\n", error_text(err));

    printf("step 4: open it leniently\n");
    OfficeDocumentHandle *doc = office_document_open_from_bytes_with(damaged, len, "docx", lenient, &err);
    check_ok(doc ? OFFICE_OK : OFFICE_ERR_PARSE, err, "a lenient open reads past a damaged part");
    char *json = take_str(office_document_damage_json(doc, &err), err, "damage json");
    char kinds[256];
    damage_kinds(json, kinds, sizeof kinds);
    printf("  damage: %s\n", kinds);
    check(strstr(kinds, "unreadable_part") != NULL, "the damage report does not name an unreadable part");
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    has(text, "Quarterly Report");
    office_oxide_free_string(text);
    office_oxide_free_string(json);
    office_document_free(doc);

    printf("step 5: truncate the archive to its first half and open it both ways\n");
    err = open_error(good, len / 2, strict);
    check(err != OFFICE_OK, "strict mode accepted a truncated archive");
    printf("  error: %s\n", error_text(err));
    err = open_error(good, len / 2, lenient);
    check(err != OFFICE_OK, "a lenient open accepted a truncated archive");
    printf("  error: %s\n", error_text(err));

    printf("step 6: a limit applies to one document only\n");
    OfficeParseOptionsHandle *limited = office_options_new();
    check(limited != NULL && office_options_set(limited, "part_bytes", 100) == OFFICE_OK, "part_bytes is a known option");
    err = open_error(good, len, limited);
    check(err != OFFICE_OK, "the part_bytes limit was not applied");
    printf("  error: %s\n", error_text(err));
    text = text_of_bytes(good, len, "docx"); /* the default limits still accept it */
    has(text, "Quarterly Report");
    free(text);

    office_options_free(strict);
    office_options_free(lenient);
    office_options_free(limited);
    free(damaged);
    free(good);
    printf("OK 08_limits_strict_damage\n");
    return 0;
}
