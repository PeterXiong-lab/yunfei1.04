import pandas as pd
import re

file_path = '321.xlsx'


def extract_to_notepad_format():

    print("--- 正在进行智能物流费用提取 ---")

    try:

        # ==================================================
        # 读取Excel
        # ==================================================

        df = pd.read_excel(file_path, header=None)

        output_blocks = []

        # ==================================================
        # 状态变量
        # ==================================================

        state = "SEARCH_HEADER"

        block_order = None
        block_city = None
        block_express_name = "韵达"

        # 更宽松的PO匹配
        order_pattern = re.compile(r'(PO-[A-Za-z0-9\-]+)')

        provinces = [

            '湖南', '湖北', '广东', '山东', '河南', '广西',
            '重庆', '四川', '江苏', '浙江', '安徽', '福建',
            '江西', '北京', '天津', '上海', '河北', '山西',
            '辽宁', '吉林', '黑龙江', '陕西', '甘肃', '青海',
            '贵州', '云南', '海南', '内蒙古', '西藏',
            '宁夏', '新疆'
        ]

        # ==================================================
        # 开始扫描
        # ==================================================

        for index, row in df.iterrows():

            # 当前行有效内容
            row_cells = [

                str(cell).strip()

                for cell in row

                if pd.notnull(cell)
                and str(cell).strip() != ''
            ]

            row_str = " ".join(row_cells)

            # 调试输出（可注释）
            # print(f"第{index+1}行：{row_str}")

            # ==================================================
            # 1. 检测仓开始（只认PO）
            # ==================================================

            order_match = order_pattern.search(row_str)

            if order_match:

                block_order = order_match.group(1)

                # 提取仓名前缀
                prefix_part = row_str.split(block_order)[0].strip()

                # 去掉省份
                for p in provinces:

                    if prefix_part.startswith(p):

                        prefix_part = prefix_part[len(p):].strip()

                        break

                # 清洗特殊字符
                prefix_part = re.sub(r'[总件：:\s]', '', prefix_part)

                # 取前两个字作为城市
                block_city = (
                    prefix_part[:2]
                    if len(prefix_part) >= 2
                    else prefix_part
                )

                # 防止空城市
                if not block_city:
                    block_city = "未知"

                # 重置
                block_express_name = "韵达"

                state = "SEARCH_COLUMNS"

                # print(f"发现仓：{block_city} {block_order}")

                continue

            # ==================================================
            # 2. 查找快递名称
            # ==================================================

            if state in ["SEARCH_COLUMNS", "SEARCH_REAL_PAY"]:

                for cell in row_cells:

                    if "费用" in cell:

                        # 排除物流费用
                        if "物流" in cell:
                            continue

                        # 排除壹米
                        if "壹米" in cell:
                            continue

                        express_name = (
                            cell
                            .replace("费用", "")
                            .replace("快递", "")
                            .strip()
                        )

                        if express_name:
                            block_express_name = express_name

                # 进入下一状态
                state = "SEARCH_REAL_PAY"

            # ==================================================
            # 3. 查找实付
            # ==================================================

            if state == "SEARCH_REAL_PAY":

                if "实付" in row_str:

                    valid_numbers = []

                    for cell in row:

                        if pd.notnull(cell):

                            try:

                                val = float(str(cell).strip())

                                valid_numbers.append(val)

                            except:
                                pass

                    # 防止缺数据
                    kuaidi_val = (
                        valid_numbers[0]
                        if len(valid_numbers) >= 1
                        else 0
                    )

                    wuliu_val = (
                        valid_numbers[1]
                        if len(valid_numbers) >= 2
                        else 0
                    )

                    # 数字格式化
                    kd_str = (
                        f"{kuaidi_val:.2f}"
                        .rstrip('0')
                        .rstrip('.')
                    )

                    wl_str = (
                        f"{wuliu_val:.5f}"
                        .rstrip('0')
                        .rstrip('.')
                    )

                    # 生成文本
                    block_text = (

                        f"{block_order}\n"
                        f"辛苦下单邮费链接备注："
                        f"天猫美团{block_city}仓{block_express_name}\n"
                        f"快递{kd_str} 物流{wl_str}"
                    )

                    output_blocks.append(block_text)

                    print(f"【成功提取】{block_order}")

                    # 重置状态
                    state = "SEARCH_HEADER"

                    block_order = None
                    block_city = None

        # ==================================================
        # 输出TXT
        # ==================================================

        if output_blocks:

            final_text = "\n\n".join(output_blocks)

            with open(
                "物流费用提取结果.txt",
                "w",
                encoding="utf-8"
            ) as f:

                f.write(final_text)

            print("\n🎉【全部完成】")
            print(f"成功提取：{len(output_blocks)} 个仓")
            print("输出文件：物流费用提取结果.txt")

        else:

            print("\n❌ 未提取到任何数据")
            print("请检查Excel格式")

    except Exception as e:

        print(f"\n运行出错：{e}")


if __name__ == '__main__':

    extract_to_notepad_format()
