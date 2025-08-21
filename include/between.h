#ifndef BETWEEN_H
#define BETWEEN_H
#include <xlnt/xlnt.hpp>

//辅助函数：判断是否在指定日期范围内
bool between(xlnt::cell cell, xlnt::date start_date, xlnt::date end_date);
#endif