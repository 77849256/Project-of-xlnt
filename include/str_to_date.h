#ifndef STR_TO_DATE_H
#define STR_TO_DATE_H
#include <string>
#include <xlnt/xlnt.hpp>
//辅助函数：解析字符串 例："2025/08/15"或"2025-08-05"，并将其转换为xlnt::date格式并输出
xlnt::date str_to_date(std::string& date_str);
#endif