#include <iostream>
#include <iomanip>
#include <string>
#include <xlnt/xlnt.hpp>
#include "print_cell_as_date.h"
#include "isDate.h"
using namespace std;
// 核心函数：将单元格转换为日期格式并打印
// 参数：cell - 待处理的Excel单元格
void print_cell_as_date(xlnt::cell cell) {
	if (!isDate(cell)) { // 先判断是否为日期，避免无效转换
		cerr << "单元格不是日期类型，无法转换" << endl;
		return;
	}

	xlnt::date date_val(1900, 1, 1); // 存储转换后的日期对象

	try {
		// 情况1：单元格是原生日期类型
		if (cell.data_type() == xlnt::cell_type::date) {
			date_val = cell.value<xlnt::date>(); // 直接获取日期对象
		}
		// 情况2：单元格是数字类型但实际是日期
		else if (cell.data_type() == xlnt::cell_type::number) {
			double date_num = cell.value<double>(); // 先获取数字值
			// 从数字转换为日期（使用Excel的windows_1900日历规则）
			date_val = xlnt::date::from_number(date_num, xlnt::calendar::windows_1900);
		}

		// 格式化日期为 YYYY-MM-DD 并打印
		int year = date_val.year;
		int month = date_val.month;
		int day = date_val.day;
		cout << "格式化日期："
			<< year << "/"
			<< setw(2) << setfill('0') << month << "/"  // 确保月份为2位（如8→08）
			<< setw(2) << setfill('0') << day << endl;   // 确保日期为2位（如5→05）
	}
	catch (const std::exception& e) {
		cerr << "转换日期失败：" << e.what() << endl;
	}
}