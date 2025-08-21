#include <iostream>
#include <xlnt/xlnt.hpp>
#include <string>
#include "column_to_letter.h"

//辅助函数：列号转字母
std::string column_to_letter(xlnt::cell cell) {
	int column = cell.column_index();	//获取列号
	std::string letter;
	while (column > 0) {
		column--;	//调整为0基索引
		char c = 'A' + (column % 26);	//计算当前字符
		letter.insert(letter.begin(), c);//计算当前位字母
		column /= 26;	//处理更高位（如AA的第二位）
	}
	return letter;
}