#ifndef COLUMN_TO_LETTER_H
#define COLUMN_TO_LETTER_H

#include <string>
#include <xlnt/xlnt.hpp>

//辅助函数：列号转字母
std::string column_to_letter(xlnt::cell cell);
#endif

