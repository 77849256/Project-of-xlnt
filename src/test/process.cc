#include <iostream>
#include <xlnt/xlnt.hpp>

int main() {
	xlnt::workbook wb;
	wb.load("E:/develop/FilterDataProject/src/test.xlsx");
	auto ws = wb.active_sheet();
	std::clog << "Processing spread sheet" << std::endl;
	// 逐行逐个遍历单元格内容，增强for循环，
	// 用变量row去遍历取出每一行，
	// 用变量cell去遍历每行的每个单元格
	for (auto row:ws.rows(false)) {
		for (auto cell:row) {
			std::clog << cell.to_string() << std::endl;//将看到的内容转换成字符串（无论什么格式）
		}
	}
	std::clog << "Processing complete" << std::endl;
	return 0;
}