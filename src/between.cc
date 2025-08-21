#include <xlnt/xlnt.hpp>
#include "between.h"


//辅助函数：判断是否在指定日期范围内
bool between(xlnt::cell cell, xlnt::date start_date, xlnt::date end_date) {
	if (cell.data_type() == xlnt::cell_type::date) {//若是日期格式就直接赋值
		xlnt::date date_val = cell.value<xlnt::date>();
		double target = date_val.to_number(xlnt::calendar::windows_1900);
		double start = start_date.to_number(xlnt::calendar::windows_1900);
		double end = end_date.to_number(xlnt::calendar::windows_1900);
		return (target >= start && target <= end);
	}
	else if (cell.data_type() == xlnt::cell_type::number) {	//若是数字格式，则要获取数字构造为日期格式，再赋值
		double date_num = cell.value<double>();
		xlnt::date date_val = xlnt::date::from_number(date_num, xlnt::calendar::windows_1900);
		double target = date_val.to_number(xlnt::calendar::windows_1900);
		double start = start_date.to_number(xlnt::calendar::windows_1900);
		double end = end_date.to_number(xlnt::calendar::windows_1900);
		return (target >= start && target <= end);

	}
	else {
		return false;
	}
}