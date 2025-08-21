#include <xlnt/xlnt.hpp>
#include <iostream>
#include <string>
#include <cmath> // 用于std::floor

int main() {
    try {
        std::string filename = "副本附件二 新建需求清单及自查自纠结果（黑龙江）.xlsx";
        xlnt::workbook wb;
        wb.load(filename);

        std::cout << "成功加载文件: " << filename << std::endl;

        xlnt::worksheet ws = wb.active_sheet();
        std::cout << "当前工作表: " << ws.title() << std::endl << std::endl;

        for (const auto& row : ws.rows()) {
            for (const auto& cell : row) {
                auto cell_type = cell.data_type();

                if (cell_type == xlnt::cell_type::number) {
                    double val = cell.value<double>();
                    if (val == std::floor(val)) {
                        std::cout << static_cast<int>(val) << "\t";
                    }
                    else {
                        std::cout << val << "\t";
                    }
                }
                else if (cell_type == xlnt::cell_type::inline_string ||
                    cell_type == xlnt::cell_type::shared_string) {
                    std::cout << cell.value<std::string>() << "\t";
                }
                else if (cell_type == xlnt::cell_type::boolean) {
                    std::cout << (cell.value<bool>() ? "true" : "false") << "\t";
                }
                else if (cell_type == xlnt::cell_type::date) {
                    std::cout << cell.value<xlnt::datetime>().to_string() << "\t";
                }
                else if (cell_type == xlnt::cell_type::error) {
                    std::cout << "[error]\t";
                }
                else if (cell_type == xlnt::cell_type::empty) {
                    std::cout << "[empty]\t";
                }
                else {
                    // 处理未明确的类型
                    std::cout << "[unknown_type]\t";
                }
            }
            std::cout << std::endl;
        }
    }
    catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return 1;
    }
    return 0;
}
