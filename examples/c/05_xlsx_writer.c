/* 05_xlsx_writer: write a workbook with typed cells, a formula, a style, a merge and a column width, then read it back.
 *
 * Run: make -C examples/c build/05_xlsx_writer && examples/c/build/05_xlsx_writer */
#include "common.h"

int main(void) {
    printf("step 1: build an XLSX in memory\n");
    size_t len = 0;
    uint8_t *data = build_xlsx_bytes(&len); /* see common.h for the cell-by-cell calls */

    printf("step 2: open it from bytes\n");
    int err = 0;
    OfficeDocumentHandle *doc = open_bytes(data, len, "xlsx");
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    char *ir_json = take_str(office_document_to_ir_json(doc, &err), err, "ir json");
    int sheets = count_sections(ir_json);
    printf("  format: %s\n  chars: %zu\n  sheets: %d\n", office_document_format(doc), strlen(text), sheets);

    printf("step 3: check the cells\n");
    check(sheets == 1, "expected one sheet");
    const char *expected[] = {"Sales", "Region", "Q1", "Q2", "Growth", "North", "120", "135", "South", "95", "110", "Total", "Notes"};
    for (size_t i = 0; i < sizeof expected / sizeof *expected; i++) has(text, expected[i]);
    check(strstr(text, "TRUE") != NULL, "the boolean cell is missing");
    check(strstr(text, "SUM(B2:B3)") && strstr(text, "SUM(C2:C3)"), "a formula is missing");

    office_oxide_free_string(text);
    office_oxide_free_string(ir_json);
    office_document_free(doc);
    office_oxide_free_bytes(data, len);
    printf("OK 05_xlsx_writer\n");
    return 0;
}
