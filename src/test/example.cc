#include <xlnt/xlnt.hpp>

int main() {
	xlnt::workbook wb;
	xlnt::worksheet ws = wb.active_sheet();
	ws.cell("A1").value(5);
	ws.cell("B2").value("string data");
	ws.cell("C3").formula("string data");
	ws.merge_cells("C3:C4");
	ws.freeze_panes("B2");
	wb.save("example.xlsx");
	return 0;
}
//compile with -std=c++14 -Ixlnt/include -Ixlnt