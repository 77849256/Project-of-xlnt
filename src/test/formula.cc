#include <xlnt/xlnt.hpp>
#include <iostream>
//#include<string>
using namespace std;

int main() {
	//创建一个工作簿
	xlnt::workbook wb;
	//获取第一个工作表
	auto ws = wb.active_sheet();
	ws.title("Formulas");

	//设置单元格的值
	
	ws.cell("B1").value("1");
	ws.cell("B2").value("2");
	ws.cell("B3").value("3");
	ws.cell("B4").value("4");
	ws.cell("B5").value("5");
	ws.cell("B6").value("6");

	ws.cell("A1").value("=SUM(B1:B5)");
	ws.cell("A2").formula("=SUM(B1:B5)");

	// 获取A1单元格的值
	string a2;
	cout << "请输入单元格的坐标，将会判断他的类型是否为字符串或公式，并输出字符串或公式，请输入：" ;
	cin >> a2;
	cout << endl;
	auto cell_a2 = ws.cell(a2);
	if (cell_a2.has_value()) {  // 检查是否有值
		//判断类型（1.5.0用cell_type())
		if (cell_a2.data_type() == xlnt::cell_type::shared_string) {
			//直接用cell.as_string()获取值（无需cell_value)
			string str_val = cell_a2.to_string();
			cout << "A2的值（字符串）：" << str_val << endl << endl;
		}
		else{
			cout << "不是字符串" << endl << endl;
		}
	}
	else if (cell_a2.has_formula()) {
		cout << "A2包含公式：" << ws.cell(a2).formula() << endl;
	}
	else {
		cout << "A2没有值和公式" << endl << endl;
	}
	//保存工作簿
	wb.save("formulas.xlsx");
	return 0;
}