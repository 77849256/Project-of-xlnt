#include <xlnt/xlnt.hpp>
#include <iostream>

int main() {
    try {
        // 1. 创建工作簿
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        ws.title("产品清单");

        // 2. 写入基础数据
        ws.cell("A1").value("ID");
        ws.cell("B1").value("产品名称");
        ws.cell("C1").value("库存");
        ws.cell("D1").value("单价");

        // 3. 样式设置（适配 1.6.1）
        auto& style_manager = wb.style(); // 获取样式管理器
        xlnt::style header_style = style_manager.create_style(); // 创建新样式（代替删除的默认构造）
        header_style.font().bold(true); // 加粗
        style_manager.add_style(header_style); // 注册样式到工作簿
        // 应用样式到标题行
        ws.cell("A1").style(header_style);
        ws.cell("B1").style(header_style);
        ws.cell("C1").style(header_style);
        ws.cell("D1").style(header_style);

        // 4. 列宽设置（适配 1.6.1，用 set_column_width）
        ws.set_column_width(xlnt::column_t("A"), 8.0);   // ID 列
        ws.set_column_width(xlnt::column_t("B"), 18.0);  // 名称列
        ws.set_column_width(xlnt::column_t("C"), 10.0);  // 库存列
        ws.set_column_width(xlnt::column_t("D"), 10.0);  // 单价列

        // 5. 写入数据行（示例）
        ws.cell("A2").value(1001);
        ws.cell("B2").value("笔记本");
        ws.cell("C2").value(50);
        ws.cell("D2").value(4999.99);

        // 6. 保存文件
        wb.save("product.xlsx");
        std::cout << "Excel 创建成功！" << std::endl;

    }
    catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return 1;
    }
    return 0;
}