#include <str_to_date.h>
#include <string>
#include <sstream>
#include <xlnt/xlnt.hpp>
#include "replace_all.h"
#include "str_to_date.h"

//辅助函数：解析字符串 例："2025/08/15"或"2025-08-05"，并将其转换为xlnt::date格式并输出
xlnt::date str_to_date(std::string& date_str) {
	//把date_str字符串的/和-替换为空格
	replace_all(date_str, '/', ' ');
	replace_all(date_str, '-', ' ');
	std::istringstream iss(date_str);
	int year, month, day;
	if (!(iss >> year >> month >> day)) {
		throw std::invalid_argument("日期格式错误！请使用YYYY-MM-DD或YYYY/MM/DD格式。");
	}
	//尝试构造xlnt::date（内部会验证日期有效性，如月份范围、天数是否合法等）
	try {
		return xlnt::date(year, month, day);
	}
	catch (const std::exception& e) {
		throw std::invalid_argument("无效日期" + std::string(e.what()));
	}
	return xlnt::date(1970, 1, 1);
}