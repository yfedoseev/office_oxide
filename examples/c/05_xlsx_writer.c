/* 05_xlsx_writer: write a workbook with typed cells, formulas, styling, a merge and a column width.
 * Run: make -C examples/c run   (or build/05_xlsx_writer after make -C examples/c) */
#include "common.h"

int main(void) {
    printf("step 1: build an XLSX in memory\n");
    size_t len = 0;
    uint8_t *bytes = build_xlsx_bytes(&len); /* see common.h for the cell-by-cell calls */

    printf("step 2: check that a write to a missing sheet is refused\n");
    OfficeXlsxWriterHandle *w = office_xlsx_writer_new();
    check(w != NULL, "xlsx writer");
    check(office_xlsx_sheet_set_cell(w, 9, 0, 0, OFFICE_CELL_STRING, "x", 0) == OFFICE_ERR_UNSUPPORTED,
          "a missing sheet is an error, not a silent drop");
    office_xlsx_writer_free(w);

    printf("step 3: open it from bytes\n");
    int err = 0;
    OfficeDocumentHandle *doc = open_bytes(bytes, len, "xlsx");
    char *text = take_str(office_document_plain_text(doc, &err), err, "plain text");
    printf("  format: %s\n  chars: %zu\n  sheets: 1\n", office_document_format(doc), strlen(text));

    printf("step 4: check the cells\n");
    check_has(text, "Sales", "the sheet name is present");
    check_has(text, "North", "text cells are present");
    check_has(text, "135", "number cells are present");
    check_has(text, "TRUE", "the boolean cell is present");
    check_has(text, "=SUM(B2:B3)", "the formula survived");
    check_has(text, "Notes", "the merged cell text is present");

    office_oxide_free_string(text);
    office_document_free(doc);
    office_oxide_free_bytes(bytes, len);
    printf("OK 05_xlsx_writer\n");
    return 0;
}
