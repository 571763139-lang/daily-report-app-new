import os
import re
import io
import json
import math
import warnings
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import unicodedata
from zipfile import ZipFile, BadZipFile
import openpyxl

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 虚拟文件系统中的路径
TEMPLATE_FILE = "新模板.xlsx"
MIS_FILE = "mis_file.xlsx"
WEIGHING_FILE = "weighing_file.xlsx"

# 全局变量由 JavaScript 动态注入：
# baiyun_text, zengcheng_text, checks_json, fly_ash_amount

def normalize_weighing_header(value):
    """兼容导出表头的全角、空白、换行和明确的吨单位，不猜测业务列名。"""
    text = unicodedata.normalize('NFKC', str(value if value is not None else ''))
    text = re.sub(r'\s+', '', text.replace('\ufeff', '').replace('\u200b', ''))
    if re.fullmatch(r'垃圾重量(?:\((?:吨|[tT])\)|\[(?:吨|[tT])\]|/(?:吨|[tT])|吨)', text):
        return '垃圾重量'
    return text


def locate_weighing_table(wb):
    """按工作表顺序扫描前50行，不依赖Excel保存时选中的活动页。"""
    required = ['垃圾重量', '货物去向', '收集者']
    candidates, duplicates, best = [], [], (0, '', 0, [])
    for sheet in wb.worksheets:
        for row_no, row in enumerate(sheet.iter_rows(min_row=1, max_row=50, max_col=200,
                                                    values_only=True), 1):
            headers = [normalize_weighing_header(value) for value in row]
            score = sum(name in headers for name in required)
            if score > best[0]:
                best = (score, sheet.title, row_no, [text for text in headers if text])
            if score == len(required):
                if any(headers.count(name) != 1 for name in required):
                    duplicates.append((sheet.title, row_no))
                    break
                candidates.append((sheet, row_no, tuple(headers.index(name) for name in required)))
                break
    if candidates:
        if len(candidates) > 1:
            report_warnings.append(f"多张工作表含称重表头，本次只读取‘{candidates[0][0].title}’，未跨表合并，请核对")
        return candidates[0]
    if duplicates:
        title, row_no = duplicates[0]
        raise ValueError(f"称重表‘{title}’第{row_no}行存在重复的垃圾重量、货物去向或收集者列，请核对")
    detail = ''
    if best[0]:
        _, title, row_no, labels = best
        detail = f" 最接近的是‘{title}’第{row_no}行，检测到：{'、'.join(labels[:12])}。"
    raise ValueError("未找到称重表头：垃圾重量、货物去向、收集者。已检查各工作表前50行，支持标题行、空格、换行及‘垃圾重量（吨）’。" + detail)


def process_weighing_data(weighing_file):
    """只读逐行汇总，避免加载 pandas；坏文件不得伪装成零数据。"""
    wb = openpyxl.load_workbook(weighing_file, read_only=True, data_only=True)
    try:
        sheet, header_row, columns = locate_weighing_table(wb)
        # 个别导出软件的dimension错误地写成A1；不让元数据截掉实际行列。
        sheet.reset_dimensions()
        rows = sheet.iter_rows(min_row=header_row + 1, max_col=max(columns) + 1, values_only=True)
        weight_col, dest_col, collector_col = columns
        total = huanfu = power = huantou = Decimal('0')
        count = 0
        for row_no, row in enumerate(rows, header_row + 1):
            if row_no > 100000:
                raise ValueError("称重表超过10万行，请精简空白格式或分批重新导出")
            if all(v is None or str(v).strip() == '' for v in row):
                continue
            # 部分导出文件每页重复一次列标题，不应把列标题当成重量。
            if [normalize_weighing_header(row[c]) for c in columns] == ['垃圾重量', '货物去向', '收集者']:
                continue
            raw = row[weight_col]
            try:
                weight = Decimal(str(raw).strip().replace(',', '').replace('，', ''))
                if not weight.is_finite() or isinstance(raw, bool):
                    raise InvalidOperation
            except (InvalidOperation, ValueError):
                raise ValueError(f"称重表第{row_no}行‘垃圾重量’不是有效数字，请核对空值、公式缓存或合计行") from None
            dest = str(row[dest_col] or '').strip()
            collector = str(row[collector_col] or '').strip()
            if weight and (not dest or not collector):
                raise ValueError(f"称重表第{row_no}行的货物去向或收集者为空，不能可靠分类")
            if weight < 0:
                report_warnings.append(f"称重表第{row_no}行含负重量，已按原值汇总，请核对是否退运冲减")
            total += weight
            if '兴丰' in dest:
                huanfu += weight
            if '资源热力电厂' in dest or '生物质处理厂' in dest:
                power += weight
            if '环投环境' in collector:
                huantou += weight
            if weight and not any(k in dest for k in ['兴丰', '资源热力电厂', '生物质处理厂']):
                report_warnings.append(f"称重表第{row_no}行去向未归入环服/热力，已计入总量，请核对分类")
            count += 1
        if not count:
            raise ValueError("称重表没有数据行；零运输日也需提供经确认的零值记录")
        if huanfu + power != total:
            report_warnings.append("处置分类之和与运输总量不一致，请核对未分类或同时命中两类的去向")
        global weighing_source_info
        weighing_source_info = {'sheet': sheet.title, 'header_row': header_row, 'records': count}
        # 先求和再四舍五入，不逐行四舍五入。
        return tuple(float(v.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
                     for v in (total, huanfu, power, huantou, total - huantou))
    finally:
        wb.close()


def extract_float(pattern, text):
    """安全提取文字中的数字"""
    match = re.search(pattern, text)
    return float(match.group(1)) if match else 0.0


def extract_metric(labels, text, display_name):
    normalized = unicodedata.normalize('NFKC', text).replace('－', '-').replace('—', '-')
    names = '|'.join(re.escape(unicodedata.normalize('NFKC', label)) for label in sorted(labels, key=len, reverse=True))
    pattern = (rf'(?:{names})[)）]?(?:进厂量|出厂量|重量|净重|量)?'
               r'(?:[（(](?:吨|立方米|m³)[）)])?[ \t:：=]*([+-]?\d[\d,，]*(?:\.\d+)?)')
    match = re.search(pattern, normalized)
    if not match:
        report_warnings.append(f"{display_name}未识别到数字，暂填0，请与来源文字核对")
        return 0.0
    value = float(match.group(1).replace(',', '').replace('，', ''))
    if not math.isfinite(value) or value < 0:
        raise ValueError(f"{display_name}须为大于等于0的有效数字")
    return value


def get_row_by_keywords(sheet, keywords_list, default_row):
    """获取区域分界行号"""
    for r in range(1, sheet.max_row + 1):
        row_text = "".join(
            [str(sheet.cell(r, c).value) for c in range(1, sheet.max_column + 1) if sheet.cell(r, c).value])
        row_text = row_text.replace(' ', '').replace('\n', '')
        for kw in keywords_list:
            if kw in row_text:
                return r
    return default_row


def inject_manual_data(sheet, keywords_list, value, start_row=1, end_row=None):
    """【网格数据注入】寻找周围独立的 0 并替换"""
    max_r = end_row or sheet.max_row
    matches = [r for r in range(start_row, max_r)
               if any(kw in str(sheet.cell(r, c).value or '')
                      for c in (3, 4) for kw in keywords_list)]
    if len(matches) != 1:
        raise ValueError(f"模板指标缺失或重复：{keywords_list[0]}")
    sheet.cell(matches[0], 5).value = value
    return True


def copy_system_data_sequential(raw_sheet, tmpl_sheet):
    """【系统数据精准注入】完美解决数字存为文本及标号穿透问题"""
    # 1. 拷贝日期标题
    title_val = None
    for r in range(1, 4):
        for c in range(1, 5):
            val = str(raw_sheet.cell(r, c).value)
            if "生产运营情况" in val:
                title_val = val
                break
        if title_val:
            break

    if title_val:
        for r in range(1, 4):
            for c in range(1, 5):
                if "生产运营情况" in str(tmpl_sheet.cell(r, c).value):
                    tmpl_sheet.cell(r, c).value = title_val
                    break

    # 2. 从上往下顺藤摸瓜寻找对应项
    tmpl_search_start = 1
    matched_rows = set()
    for raw_r in range(1, raw_sheet.max_row + 1):
        numbers, texts = [], []

        for c in range(1, raw_sheet.max_column + 1):
            val = raw_sheet.cell(raw_r, c).value
            if val is None:
                continue

            if isinstance(val, (int, float)):
                numbers.append(val)
            elif isinstance(val, str):
                clean_str = val.strip()
                try:
                    num_val = float(clean_str.replace(',', ''))
                    numbers.append(num_val)
                except ValueError:
                    clean_str_no_space = clean_str.replace(' ', '').replace('\n', '')
                    if len(clean_str_no_space) > 1:
                        texts.append(clean_str_no_space)

        if numbers and texts:
            best_label = texts[-1]

            p_raw = re.match(r'^(\d+(?:\.\d+)*)[-－—_~、.]', best_label)
            raw_prefix = p_raw.group(1) if p_raw else None

            found_tmpl_r = -1
            for t_r in range(tmpl_search_start, tmpl_sheet.max_row + 1):
                row_match = False
                for c in range(1, tmpl_sheet.max_column + 1):
                    t_val_raw = tmpl_sheet.cell(t_r, c).value
                    if not isinstance(t_val_raw, str):
                        continue

                    t_val = t_val_raw.replace(' ', '').replace('\n', '')
                    if len(t_val) < 2:
                        continue

                    if raw_prefix:
                        p_tmpl = re.match(r'^(\d+(?:\.\d+)*)[-－—_~、.]', t_val)
                        if p_tmpl and p_tmpl.group(1) == raw_prefix:
                            row_match = True
                            break
                    else:
                        if best_label in t_val or t_val in best_label:
                            row_match = True
                            break

                if row_match:
                    found_tmpl_r = t_r
                    break

            if found_tmpl_r != -1:
                matched_rows.add(found_tmpl_r)
                num_idx = 0
                for c in range(1, tmpl_sheet.max_column + 1):
                    cell = tmpl_sheet.cell(found_tmpl_r, c)
                    if str(cell.value).strip() in ['0', '0.0', '0.00']:
                        cell.value = numbers[num_idx]
                        num_idx += 1
                        if num_idx >= len(numbers):
                            break
                tmpl_search_start = found_tmpl_r + 1

    if len(matched_rows) < 5:
        raise ValueError("MIS表未匹配到足够的生产指标，请确认选中的是生产运营导出表")
    end = get_row_by_keywords(tmpl_sheet, ['增城建废'], tmpl_sheet.max_row)
    for r in range(4, end):
        label = str(tmpl_sheet.cell(r, 4).value or tmpl_sheet.cell(r, 3).value or '')
        if r not in matched_rows and '飞灰外运资源化' not in label:
            report_warnings.append(f"MIS未匹配指标：{label}，请核对模板中的零值")


def fill_report_table_data(sheet, fly_ash, to_huanfu, to_power, by_huantou, by_social):
    """按指标标签填入新版模板，避免插行后依赖固定行号。"""
    values = {
        "1.5-飞灰外运资源化（吨）": fly_ash,
        "环境集团（吨）": by_huantou,
        "社会单位（吨）": by_social,
        "环服公司（吨）": to_huanfu,
        "各资源热力电厂（吨）": to_power,
    }
    for label, value in values.items():
        rows = [r for r in range(1, sheet.max_row + 1)
                if any(isinstance(sheet.cell(r, c).value, str)
                       and sheet.cell(r, c).value.replace(' ', '').replace('\n', '') == label
                       for c in (3, 4))]
        if len(rows) != 1:
            raise ValueError(f"新版模板指标缺失或重复：{label}，请刷新页面加载新版模板")
        sheet.cell(rows[0], 5).value = float(value)


def fix_bottom_paragraphs(raw_sheet, tmpl_sheet, checks):
    """更新机炉运行情况和检查内容；污水数据由表格单独填入。"""
    boiler_str, turbine_str = None, None

    # 2. 获取锅炉/汽机正常数
    for r in range(1, raw_sheet.max_row + 1):
        for c in range(1, raw_sheet.max_column + 1):
            val = str(raw_sheet.cell(r, c).value)
            if "锅炉总数" in val:
                boiler_str = raw_sheet.cell(r, c).value
            if "汽机总数" in val:
                turbine_str = raw_sheet.cell(r, c).value

    # 3. 构造检查部分的替换字符串
    if checks:
        check_lines = []
        for chk in checks[:6]:
            val_strip = str(chk).strip()
            if val_strip and val_strip != "无":
                check_lines.append(f"{len(check_lines) + 1}.{val_strip}")
        if check_lines:
            check_text = "检查：\n" + "\n".join(check_lines)
        else:
            check_text = "检查：\n  无。"
    else:
        check_text = "检查：\n  无。"

    # 4. 表尾单元格遍历更新
    for r in range(1, tmpl_sheet.max_row + 1):
        for c in range(1, tmpl_sheet.max_column + 1):
            cell = tmpl_sheet.cell(r, c)
            val = cell.value

            if isinstance(val, str):
                if "锅炉总数" in val and boiler_str:
                    cell.value = boiler_str
                    continue
                if "汽机总数" in val and turbine_str:
                    cell.value = turbine_str
                    continue

                if "检查：" in val:
                    cell.value = check_text


def main():
    global report_warnings, weighing_source_info
    report_warnings = []
    weighing_source_info = {}
    for path, label in [(MIS_FILE, 'MIS'), (WEIGHING_FILE, '称重')]:
        try:
            with ZipFile(path) as archive:
                if sum(item.file_size for item in archive.infolist()) > 100 * 1024 * 1024:
                    raise ValueError(f"{label}文件展开后超过100MB，请精简空白格式或重新导出")
        except BadZipFile:
            raise ValueError(f"{label}文件不是有效的.xlsx，请用Excel重新另存为xlsx，不要只改后缀") from None
    try:
        fly_ash = float(fly_ash_amount)
    except (NameError, TypeError, ValueError):
        raise ValueError("请填写有效的飞灰外运资源化量（吨）") from None
    if not math.isfinite(fly_ash) or fly_ash < 0:
        raise ValueError("飞灰外运资源化量必须为大于等于 0 的有效数字")

    try:
        # 反序列化 checks 数组
        checks = json.loads(checks_json)
    except Exception:
        checks = []

    # 1. 污水数据统计（新版：5项返回值）
    tot_sewage, to_huanfu, to_power, by_huantou, by_social = process_weighing_data(WEIGHING_FILE)

    # 2. 解析白云建废文字数据
    by_shizha = extract_metric(['石渣', '石渣）'], baiyun_text, '白云石渣')
    by_shuinitou = extract_metric(['水泥头', '水泥头）'], baiyun_text, '白云水泥头')
    by_yuanshengshi = extract_metric(['原生石', '原生石）'], baiyun_text, '白云原生石')
    by_jianfei_total = by_shizha + by_shuinitou + by_yuanshengshi

    by_hunningtu = extract_metric(['原生混凝土'], baiyun_text, '白云原生混凝土')
    by_zaisheng12 = extract_metric(['再生1-2石', '再生骨料1-2石'], baiyun_text, '白云再生1-2石')
    by_zaishengshuiwen = extract_metric(['再生水稳'], baiyun_text, '白云再生水稳')

    # 3. 解析增城建废文字数据
    z_zx = extract_metric(['装修垃圾'], zengcheng_text, '增城装修垃圾')
    z_sw = extract_metric(['再生水稳'], zengcheng_text, '增城再生水稳')
    z_sf = extract_metric(['再生石粉'], zengcheng_text, '增城再生石粉')
    z_zz = extract_metric(['砖渣'], zengcheng_text, '增城砖渣')
    z_hnt = extract_metric(['混凝土块', '混凝土'], zengcheng_text, '增城混凝土')
    if '出厂' in zengcheng_text:
        report_warnings.append("增城来源含‘出厂’，而模板砖渣/混凝土行标为‘进厂’，请人工确认口径，程序不做进出厂转换")

    # 4. 读取工作簿填充
    tmpl_wb = openpyxl.load_workbook(TEMPLATE_FILE)
    tmpl_sheet = tmpl_wb.active
    raw_wb = openpyxl.load_workbook(MIS_FILE, data_only=True)
    raw_sheet = raw_wb.active
    if raw_sheet.max_row > 2000 or raw_sheet.max_column > 100:
        raise ValueError("MIS表范围过大，请删除无用空白行列后重新导出（最多2000行、100列）")

    # 动态日期提取
    report_date_str = None

    for r in range(1, 4):
        for c in range(1, 5):
            val = str(raw_sheet.cell(r, c).value)
            if "生产运营情况" in val:
                match = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', val)
                if match:
                    report_date_str = f"{match.group(1)}{int(match.group(2)):02d}{int(match.group(3)):02d}"
                    datetime.strptime(report_date_str, '%Y%m%d')
                break

    if report_date_str is None:
        raise ValueError("MIS表前3行标题中缺少有效的 YYYY-MM-DD 生产日期，请核对导出文件")

    output_filename = f"安全生产每日简报 {report_date_str}.xlsx"

    # 执行注入
    copy_system_data_sequential(raw_sheet, tmpl_sheet)

    r_zc = get_row_by_keywords(tmpl_sheet, ["增城建废", "增城"], 50)
    r_by = get_row_by_keywords(tmpl_sheet, ["白云建废", "白云"], r_zc + 10)
    r_sewage = get_row_by_keywords(tmpl_sheet, ["生活垃圾污水"], r_by + 10)

    inject_manual_data(tmpl_sheet, ["装修垃圾"], z_zx, r_zc, r_by)
    inject_manual_data(tmpl_sheet, ["砖渣"], z_zz, r_zc, r_by)
    inject_manual_data(tmpl_sheet, ["混凝土"], z_hnt, r_zc, r_by)
    inject_manual_data(tmpl_sheet, ["再生水稳"], z_sw, r_zc, r_by)
    inject_manual_data(tmpl_sheet, ["再生石粉"], z_sf, r_zc, r_by)

    inject_manual_data(tmpl_sheet, ["建筑废弃物进厂", "建废"], by_jianfei_total, r_by, r_sewage)
    inject_manual_data(tmpl_sheet, ["混凝土出厂量", "原生混凝土", "混凝土出厂"], by_hunningtu, r_by, r_sewage)
    inject_manual_data(tmpl_sheet, ["再生骨料1-2石", "再生1-2石", "再生骨料"], by_zaisheng12, r_by, r_sewage)
    inject_manual_data(tmpl_sheet, ["再生水稳"], by_zaishengshuiwen, r_by, r_sewage)

    fill_report_table_data(tmpl_sheet, fly_ash, to_huanfu, to_power, by_huantou, by_social)
    fix_bottom_paragraphs(raw_sheet, tmpl_sheet, checks)

    # 保存
    tmpl_wb.save("output_file.xlsx")
    raw_wb.close()
    tmpl_wb.close()

    # 封包预览数据返回
    parsed_preview = {
        "fly_ash": fly_ash,
        "sewage": {
            "total": tot_sewage,
            "to_huanfu": to_huanfu,
            "to_power": to_power,
            "by_huantou": by_huantou,
            "by_social": by_social
        },
        "baiyun": {
            "hunningtu": by_hunningtu,
            "shizha": by_shizha,
            "shuinitou": by_shuinitou,
            "total": by_jianfei_total,
            "yuanshengshi": by_yuanshengshi,
            "zaisheng12": by_zaisheng12,
            "zaishengshuiwen": by_zaishengshuiwen
        },
        "zengcheng": {
            "装修垃圾": z_zx,
            "再生水稳": z_sw,
            "再生石粉": z_sf,
            "砖渣": z_zz,
            "混凝土块": z_hnt
        }
    }

    global output_json_str
    output_json_str = json.dumps({
        "success": True,
        "parsed_data": parsed_preview,
        "filename": output_filename,
        "weighing_source": weighing_source_info,
        "warnings": list(dict.fromkeys(report_warnings))[:30]
    }, ensure_ascii=False, allow_nan=False)

main()
