#include <string>
#include "replace_all.h"


//辅助函数：字符替换
void replace_all(std::string& str, char target, char blank) 	//使用引用 令用户可以改变参数str的实际值
{
	for (char& c : str) {
		if (c == target) {	//该字符是需要替换的
			c = blank;
		}
	}
}
