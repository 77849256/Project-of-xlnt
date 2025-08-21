#include <iostream>
#include <xlnt/xlnt.hpp>
#include <string>
#include <algorithm>

#include "isDate.h"

//辅助函数：判断是否为日期格式
bool isDate(xlnt::cell& cell) {
	//情况1：直接是日期类型
	if (cell.data_type() == xlnt::cell_type::date) {
		return true;
	}
	//情况2：数字类型但格式为日期
	if (cell.data_type() == xlnt::cell_type::number) {	//非日期格式
		std::string format = cell.number_format().format_string();
		std::transform(format.begin(), format.end(), format.begin(), ::tolower);

		//扩展关键词：增加yy（两位年份）、m（单字母月）、d（单字母日）
		if (format.find("yyyy") != std::string::npos ||
			format.find("yy") != std::string::npos ||
			format.find("mm") != std::string::npos ||
			format.find("m") != std::string::npos ||
			format.find("dd") != std::string::npos ||
			format.find("d") != std::string::npos ||
			format.find("年") != std::string::npos ||
			format.find("月") != std::string::npos ||
			format.find("日") != std::string::npos) {
			return true;
		}
	}
	//cout << "单元格" << column_to_letter(cell) << cell.row() << "数据不是日期格式。" << endl;
	return false;	//确保此处必然返回
}