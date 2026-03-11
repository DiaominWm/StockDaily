import requests
import time
import json
import sys
import os
from datetime import datetime, timedelta
from typing import Dict, List, Optional

# ===================== 全局配置(可直接修改。无需改代码逻辑) =====================
# 【核心可配置项】
# 1. 播报间隔(秒)：直接修改数值即可。比如改30=30秒。改120=2分钟
REPORT_INTERVAL = 15  # 原INTERVAL改为全局变量。命名更清晰
# 2. 网络请求重试次数(失败后自动重试)
REQUEST_RETRY_TIMES = 3
# 3. 重试间隔(秒)：每次重试前等待时间
RETRY_INTERVAL = 2

# 调整列宽适配内容。让表头和数据视觉对齐,不要乱动
COL_WIDTHS = {
    "code": 14,        # 代码列：18字符
    "name": 12,        # 名称列：12字符(适配股票名长度)
    "current": 12,     # 当前价列：12字符
    "open": 12,        # 开盘价列：12字符
    "high": 12,        # 最高价列：12字符
    "low": 12,         # 最低价列：12字符
    "change": 12       # 涨跌幅列：12字符
}
COL_HEADER_WIDTHS = {
    "code": 12,        # 代码列：18字符
    "name": 14,        # 名称列：18字符(适配股票名长度)
    "current": 9,     # 当前价列：12字符
    "open": 9,        # 开盘价列：12字符
    "high": 9,        # 最高价列：12字符
    "low": 9,         # 最低价列：12字符
    "change": 9       # 涨跌幅列：12字符
}
# 【股票配置】
# 需监控的10只A股代码(格式：股票代码+市场。sh=沪市/科创板。sz=深市/创业板)
STOCK_CODES = [
    "sh600699",  # 均胜电子
    "sh600783",  # 鲁信创投
    "sz002156",  # 通富微电
    "sz002957",  # 科瑞技术
    "sz300058",  # 蓝色光标
    "sz300115",  # 长盈精密
]

# 【其他配置】
# 持仓信息(首次运行会引导手动输入。也可直接在这里配置)
HOLDINGS = {}
# 股票数据接口
STOCK_API = "https://hq.sinajs.cn/list={}"
# 是否清屏(True=清屏。False=保留历史数据)
CLEAR_SCREEN = True

# ===================== 环境检查与报错机制 =====================
def check_environment() -> None:
    """检查运行环境。包含依赖和网络。异常则终止程序"""
    print("🔍 正在检查运行环境...")

    # 1. 检查Python版本
    if sys.version_info < (3, 7):
        raise RuntimeError("❌ Python版本需≥3.7。请升级后重试！")

    # 2. 检查依赖包
    required_packages = ["requests"]
    missing_packages = []
    for pkg in required_packages:
        try:
            __import__(pkg)
        except ImportError:
            missing_packages.append(pkg)

    if missing_packages:
        raise ImportError(
            f"❌ 缺少依赖包：{', '.join(missing_packages)}\n"
            f"请执行:pip install {' '.join(missing_packages)}"
        )

    # 3. 检查网络连接
    retry_count = 0
    while retry_count < REQUEST_RETRY_TIMES:
        try:
            response = requests.get("https://www.sina.com.cn", timeout=5)
            response.raise_for_status()
            break
        except requests.exceptions.RequestException:
            retry_count += 1
            print(f"⚠️  网络连接检查失败。第{retry_count}次重试...")
            time.sleep(RETRY_INTERVAL)
    else:
        raise ConnectionError("❌ 网络连接失败。请检查网络后重试！")

    print("✅ 环境检查通过！")

# ===================== 股票数据获取(带自动重试) =====================
def get_stock_data(stock_codes: List[str]) -> Dict[str, Dict]:
    """
    获取股票实时数据(适配新浪接口最新格式+自动重试)
    :param stock_codes: 股票代码列表
    :return: 股票数据字典,key=股票代码。value=详细数据
    """
    # 循环重试逻辑
    for retry in range(REQUEST_RETRY_TIMES):
        try:
            # 拼接股票代码请求接口
            codes_str = ",".join(stock_codes)
            # 添加请求头。模拟浏览器访问
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Referer": "https://finance.sina.com.cn/",
                "Accept-Language": "zh-CN,zh;q=0.9"
            }
            response = requests.get(STOCK_API.format(codes_str), headers=headers, timeout=10)
            response.raise_for_status()
            response.encoding = "gbk"
            data_lines = response.text.strip().split("\n")

            stock_data = {}
            for line in data_lines:
                if not line or "hq_str_" not in line:
                    continue

                # 解析接口返回数据
                try:
                    # 提取代码和数据部分
                    code_part, data_part = line.split("=", 1)
                    code = code_part.replace("var hq_str_", "").strip()
                    # 清理数据部分的引号和分号
                    values = data_part.strip().strip('";').split(",")

                    # 校验数据长度(至少需要33个字段。确保关键字段存在)
                    if len(values) < 33:
                        print(f"⚠️  股票{code}返回字段不足。跳过")
                        continue

                    # 适配最新字段索引(修正核心：正确提取关键字段)
                    name = values[0] if values[0] else "未知股票"
                    open_price = float(values[1]) if values[1] else 0.0  # 开盘价
                    pre_close = float(values[2]) if values[2] else 0.0  # 昨收盘价(涨跌幅基准)
                    current_price = float(values[3]) if values[3] else 0.0  # 当前价
                    high_price = float(values[4]) if values[4] else 0.0  # 最高价
                    low_price = float(values[5]) if values[5] else 0.0  # 最低价

                    # 计算当日涨跌幅(核心修正：标准A股涨跌幅公式)
                    if pre_close != 0:
                        change_price = current_price - pre_close  # 涨跌额(当前价-昨收价)
                        change_percent = (change_price / pre_close) * 100  # 涨跌幅(%)
                    else:
                        change_price = 0.0
                        change_percent = 0.0

                    stock_data[code] = {
                        "name": name,
                        "current_price": current_price,
                        "open_price": open_price,
                        "high_price": high_price,
                        "low_price": low_price,
                        "change_percent": round(change_percent, 2),  # 保留2位小数
                        "change_price": round(change_price, 2),
                        "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                except (ValueError, IndexError) as e:
                    print(f"⚠️  股票{code}数据解析失败。错误：{str(e)}")
                    continue

            if stock_data:
                return stock_data
            elif retry < REQUEST_RETRY_TIMES - 1:
                print(f"⚠️  未获取到股票数据。第{retry+1}次重试...")
                time.sleep(RETRY_INTERVAL)

        except Exception as e:
            if retry < REQUEST_RETRY_TIMES - 1:
                print(f"⚠️  获取股票数据失败({str(e)})。第{retry+1}次重试...")
                time.sleep(RETRY_INTERVAL)
            else:
                raise Exception(f"❌ 多次重试后仍获取失败：{str(e)}")

    # 所有重试都失败
    return {}

# ===================== 持仓盈亏计算 =====================
def input_holdings() -> Dict[str, Dict]:
    """引导用户手动输入持仓信息"""
    print("\n📝 请输入持仓信息(输入完成后按回车确认。留空则跳过)")
    holdings = {}

    while True:
        # 显示可选的股票代码
        print(f"\n可选股票代码：{', '.join(STOCK_CODES)}")
        code = input("请输入持仓股票代码(如sh600000)：").strip()

        if not code:
            break

        if code not in STOCK_CODES:
            print(f"❌ 股票代码{code}不在监控列表中。请重新输入！")
            continue

        # 输入成本价
        while True:
            cost_input = input(f"请输入{code}的成本价：").strip()
            try:
                cost_price = float(cost_input)
                if cost_price <= 0:
                    print("❌ 成本价需大于0。请重新输入！")
                    continue
                break
            except ValueError:
                print("❌ 成本价必须是数字。请重新输入！")

        # 输入持仓数量
        while True:
            qty_input = input(f"请输入{code}的持仓数量：").strip()
            try:
                quantity = int(qty_input)
                if quantity <= 0:
                    print("❌ 持仓数量需大于0。请重新输入！")
                    continue
                break
            except ValueError:
                print("❌ 持仓数量必须是整数。请重新输入！")

        holdings[code] = {
            "cost_price": cost_price,
            "quantity": quantity
        }
        print(f"✅ 已添加持仓：{code} - 成本价{cost_price}。数量{quantity}")

    return holdings

def calculate_profit_loss(stock_data: Dict[str, Dict], holdings: Dict[str, Dict]) -> Dict[str, Dict]:
    """计算持仓盈亏"""
    profit_data = {}
    for code, hold in holdings.items():
        if code not in stock_data:
            print(f"⚠️  未获取到{code}的实时数据。跳过盈亏计算")
            continue

        current_price = stock_data[code]["current_price"]
        cost_price = hold["cost_price"]
        quantity = hold["quantity"]

        # 计算盈亏
        total_cost = cost_price * quantity
        if total_cost == 0:
            print(f"⚠️  {code}总成本为0。跳过盈亏计算")
            continue

        total_current = current_price * quantity
        profit_amount = total_current - total_cost
        profit_percent = (profit_amount / total_cost) * 100

        profit_data[code] = {
            "cost_price": cost_price,
            "current_price": current_price,
            "quantity": quantity,
            "total_cost": total_cost,
            "total_current": total_current,
            "profit_amount": round(profit_amount, 2),
            "profit_percent": round(profit_percent, 2)
        }

    return profit_data

# ===================== 终端播报(修改盈亏分析为键值对格式) =====================
def left_align_fixed(content: str, width: int) -> str:
    """
    强制左对齐：内容从列的第一个字符开始。不足补空格到固定宽度
    确保表头和数据在同一列的起始位置完全一致
    """
    content_str = str(content).strip()
    # 超长截断(避免撑列)。不足补空格(核心：左对齐同位置)
    if len(content_str) > width:
        return content_str[:width-1] + "…"
    # ljust：内容左对齐。右侧补空格。保证列宽固定
    return content_str.ljust(width)

def print_stock_report(stock_data: Dict[str, Dict], profit_data: Dict[str, Dict]) -> None:
    if CLEAR_SCREEN:
        os.system("clear" if os.name == "posix" else "cls")

    separator = "=" * 120
    update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(separator)
    print(f"📈 股票实时播报 | 更新时间：{update_time}".ljust(120))
    print(separator)

    # ================== 股票实时数据(去掉分隔线+精准对齐)==================
    print("\n📊 股票实时数据：")

    # 构建表头(竖线分隔。每列左对齐。宽度固定)
    header = (
        f"| {'代    码':<{COL_HEADER_WIDTHS['code'] - 2}} "  # -2是预留竖线和空格
        f"| {'名    称':<{COL_HEADER_WIDTHS['name'] - 2}} "
        f"| {'当前价':<{COL_HEADER_WIDTHS['current'] - 2}} "
        f"| {'开盘价':<{COL_HEADER_WIDTHS['open'] - 2}} "
        f"| {'最高价':<{COL_HEADER_WIDTHS['high'] - 2}} "
        f"| {'最低价':<{COL_HEADER_WIDTHS['low'] - 2}} "
        f"| {'涨跌幅(%)':<{COL_HEADER_WIDTHS['change'] - 2}} |"
    )
    # 只打印表头(删除分隔线)
    print(header)

    # 打印数据行(和表头列宽1:1匹配)
    for code, data in stock_data.items():
        # 提前计算各数值的字符串(避免引号冲突)
        current_price_str = f"{data['current_price']:.2f}"
        open_price_str = f"{data['open_price']:.2f}"
        high_price_str = f"{data['high_price']:.2f}"
        low_price_str = f"{data['low_price']:.2f}"
        change_percent_str = f"{data['change_percent']:+.2f}"

        row = (
            f"| {code:<{COL_WIDTHS['code'] - 2}} "  # 代码列：左对齐。占16字符
            f"| {data['name']:<{COL_WIDTHS['name'] - 2}} "  # 名称列：左对齐。占16字符
            f"| {current_price_str:<{COL_WIDTHS['current'] - 2}} "  # 当前价：左对齐
            f"| {open_price_str:<{COL_WIDTHS['open'] - 2}} "  # 开盘价：左对齐
            f"| {high_price_str:<{COL_WIDTHS['high'] - 2}} "  # 最高价：左对齐
            f"| {low_price_str:<{COL_WIDTHS['low'] - 2}} "  # 最低价：左对齐
            f"| {change_percent_str:<{COL_WIDTHS['change'] - 2}} |"  # 涨跌幅：左对齐
        )
        print(row)

    # ================== 持仓盈亏分析(改为键值对格式)==================
    if profit_data:
        print("\n💰 持仓盈亏分析：")
        print("-" * 120)  # 分隔线

        total_profit = 0.0
        # 逐行输出每个股票的盈亏信息(键值对格式)
        for code, profit in profit_data.items():
            # 构建键值对字符串。每个字段之间保留固定空格。排版整洁
            stock_name = stock_data.get(code, {}).get("name", "未知股票")
            profit_line = (
                f"名称：{stock_name:<6}  "
                f"成本价：{profit['cost_price']:<6.2f}  "
                f"当前价：{profit['current_price']:<6.2f}  "
                f"数量：{profit['quantity']:<4d}  "
                f"总成本：{profit['total_cost']:<10.2f}  "
                f"当前市值：{profit['total_current']:<10.2f}  "
                f"盈亏金额：{profit['profit_amount']:<+10.2f}  "
                f"盈亏百分比：{profit['profit_percent']:<+8.2f}%"
            )
            print(profit_line)
            total_profit += profit["profit_amount"]

        # 总盈亏汇总
        print("-" * 120)
        total_profit_line = f"📝 总盈亏：{total_profit:>+10.2f} 元"
        print(total_profit_line)

    # 底部信息
    next_time = (datetime.now() + timedelta(seconds=REPORT_INTERVAL)).strftime("%H:%M:%S")
    tip = f"⌛ 下次播报时间：{next_time} | 已收盘 | 间隔 {REPORT_INTERVAL}秒"
    print(f"\n{separator}")
    print(left_align_fixed(tip, 180))
    print(separator)

# ===================== 主程序 =====================
def main():
    try:
        # 1. 环境检查
        check_environment()

        # 2. 引导输入持仓信息
        if not HOLDINGS:
            HOLDINGS.update(input_holdings())
            if not HOLDINGS:
                print("\n⚠️  未配置持仓信息。仅显示股票实时数据！")

        # 3. 循环获取数据并播报
        print(f"\n🚀 股票实时播报程序已启动。每{REPORT_INTERVAL}秒播报一次(按Ctrl+C退出)...")
        while True:
            try:
                # 获取股票数据
                stock_data = get_stock_data(STOCK_CODES)
                if not stock_data:
                    print("❌ 未获取到任何股票数据！")
                    time.sleep(REPORT_INTERVAL)
                    continue

                # 计算盈亏
                profit_data = calculate_profit_loss(stock_data, HOLDINGS)

                # 打印播报
                print_stock_report(stock_data, profit_data)

                # 等待指定间隔
                time.sleep(REPORT_INTERVAL)

            except KeyboardInterrupt:
                print("\n\n🛑 程序已手动终止！")
                sys.exit(0)
            except Exception as e:
                print(f"\n❌ 单次播报异常：{str(e)},{REPORT_INTERVAL}秒后重试...")
                time.sleep(REPORT_INTERVAL)

    except Exception as e:
        print(f"\n💥 程序启动失败：{str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()