/* 03_edit: replace text in a DOCX and set cells in an XLSX, in memory.
 *
 * Run: make -C examples/c build/03_edit && examples/c/build/03_edit */
#include "common.h"

int main(void) {
    int err = 0;

    printf("step 1: build a DOCX in memory\n");
    size_t len = 0;
    uint8_t *docx = build_docx_bytes("Quarterly Report", &len);

    printf("step 2: replace text\n");
    OfficeEditableHandle *ed = office_editable_open_from_bytes(docx, len, "docx", &err);
    check_ok(ed ? OFFICE_OK : OFFICE_ERR_PARSE, err, "open editable docx");
    free(docx);
    int64_t replaced = office_editable_replace_text(ed, "Quarterly Report", "Annual Report", &err);
    printf("  replacements: %ld\n", (long)replaced);
    check(replaced == 1 && err == OFFICE_OK, "replace_text did not make exactly one replacement");

    printf("step 3: save to bytes and check the text\n");
    size_t out_len = 0;
    uint8_t *edited = office_editable_save_to_bytes(ed, &out_len, &err);
    check_ok(edited ? OFFICE_OK : OFFICE_ERR_INTERNAL, err, "save docx to bytes");
    office_editable_free(ed);
    char *text = text_of_bytes(edited, out_len, "docx");
    has(text, "Annual Report");
    check(strstr(text, "Quarterly Report") == NULL, "the old title is still present");
    free(text);
    office_oxide_free_bytes(edited, out_len);

    printf("step 4: build an XLSX in memory\n");
    uint8_t *xlsx = build_xlsx_bytes(&len);

    printf("step 5: set cells\n");
    ed = office_editable_open_from_bytes(xlsx, len, "xlsx", &err);
    check_ok(ed ? OFFICE_OK : OFFICE_ERR_PARSE, err, "open editable xlsx");
    office_oxide_free_bytes(xlsx, len);
    check_ok(office_editable_set_cell(ed, 0, "B2", OFFICE_CELL_NUMBER, NULL, 999, &err), err, "set B2");
    check_ok(office_editable_set_cell(ed, 0, "A2", OFFICE_CELL_STRING, "Central", 0, &err), err, "set A2");

    printf("step 6: save to bytes and check the cells\n");
    edited = office_editable_save_to_bytes(ed, &out_len, &err);
    check_ok(edited ? OFFICE_OK : OFFICE_ERR_INTERNAL, err, "save xlsx to bytes");
    office_editable_free(ed);
    text = text_of_bytes(edited, out_len, "xlsx");
    has(text, "Central");
    has(text, "999");
    check(strstr(text, "North") == NULL, "the old A2 value is still present");
    free(text);
    office_oxide_free_bytes(edited, out_len);

    printf("OK 03_edit\n");
    return 0;
}
