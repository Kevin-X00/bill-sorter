#!/usr/bin/env python3
"""
Bill Sorter MVP
Excel/CSV 导入 → 自动归类 → 月度报表
支持：支付宝CSV、微信CSV、任意标准格式
"""

import io
import re
import json
import hashlib
from datetime import datetime, timedelta
from collections import defaultdict

import pandas as pd
from flask import Flask, render_template, request, jsonify, session
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'bill-sorter-mvp-key-2026'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# ============================================================
# 归类别名库
# ============================================================
RULES = [
    # 餐饮
    (r'瑞幸|星巴克|Manner|库迪|咖啡|茶颜|喜茶|奈雪|蜜雪|一点点', '餐饮·饮品'),
    (r'美团外卖|饿了么|外卖|美团【|盒马|叮咚|朴朴', '餐饮·外卖'),
    (r'麦当劳|肯德基|汉堡王|必胜客|达美乐|老乡鸡|沙县|兰州|海底捞|火锅|烤肉|日料|川菜|湘菜', '餐饮·正餐'),
    (r'全家|罗森|7-11|711|便利店|超市|沃尔玛|山姆|Costco|永辉', '餐饮·零售'),

    # 交通
    (r'滴滴|高德打车|T3|曹操|首汽|花小猪', '交通·打车'),
    (r'地铁|公交|上海公共交通|北京公交', '交通·公交地铁'),
    (r'铁路|12306|高铁|火车票', '交通·城际'),
    (r'中国航信|国航|东航|南航|春秋|机票|航旅', '交通·机票'),
    (r'中石化|中石油|壳牌|加油|充电桩', '交通·加油充电'),

    # 购物
    (r'京东|淘宝|天猫|拼多多|闲鱼|转转|得物|唯品会|抖音商城', '购物·电商'),
    (r'名创优品|无印良品|宜家|NITORI|小米之家', '购物·日用'),
    (r'优衣库|ZARA|H&M|Nike|Adidas|安踏|李宁|服装|鞋', '购物·服饰'),

    # 居住
    (r'物业|水费|电费|燃气|国家电网|自来水|燃气公司', '居住·水电物业'),
    (r'房租|自如|贝壳|链家|我爱我家|蛋壳', '居住·租房'),

    # 通信
    (r'中国移动|中国联通|中国电信|话费|宽带', '通信·话费'),

    # 金融
    (r'房贷|还款|贷款|按揭|公积金|建设银行|工商银行|招商银行|还贷', '金融·房贷'),
    (r'信用卡|花呗|白条|最低还款', '金融·信用卡'),

    # 娱乐
    (r'腾讯视频|爱奇艺|优酷|B站|bilibili|网易云|QQ音乐|Spotify|Netflix|Disney', '娱乐·订阅'),
    (r'Steam|Epic|PSN|Switch|Xbox|游戏|原神|王者', '娱乐·游戏'),
    (r'电影院|电影票|猫眼|淘票票|演出|话剧|音乐节', '娱乐·文娱'),

    # 医疗
    (r'医院|诊所|挂号|药品|药房|医保|体检|牙科', '医疗健康'),

    # 教育
    (r'得到|极客时间|知识星球|小鹅通|课程|培训|驾校|考试|报名', '教育·学习'),

    # 转账
    (r'转账|红包|微信红包', '社交·转账'),
]


def classify_by_rules(merchant: str, category_hint: str = '') -> str:
    """先用规则匹配，规则匹配不到用已有类目提示"""
    if not merchant:
        return '未分类'
    merchant_lower = merchant.lower()
    for pattern, cat in RULES:
        if re.search(pattern, merchant_lower, re.I):
            return cat
    # 如果原始数据有类目信息，保留
    if category_hint and category_hint != '未分类':
        return category_hint
    return '未分类'


def parse_alipay_csv(file_content: str) -> list[dict]:
    """解析支付宝 CSV 账单"""
    records = []
    lines = file_content.strip().split('\n')
    # 支付宝 CSV 格式：
    # 交易时间, 交易分类, 交易对方, 商品说明, 收/支, 金额, 支付方式, 交易状态, 交易订单号
    for line in lines:
        if not line.strip():
            continue
        # 跳过表头或非数据行
        if '交易时间' in line and '交易分类' in line:
            continue
        parts = [p.strip().strip('"') for p in line.split(',')]
        if len(parts) < 6:
            continue
        date_str = parts[0]
        category = parts[1] if len(parts) > 1 else ''
        merchant = parts[2] if len(parts) > 2 else ''
        description = parts[3] if len(parts) > 3 else ''
        direction = parts[4] if len(parts) > 4 else ''
        amount_str = parts[5] if len(parts) > 5 else ''

        if not direction or '支出' not in direction:
            continue

        try:
            amount = float(amount_str.replace('¥', '').replace(',', '').strip())
        except (ValueError, AttributeError):
            continue

        try:
            date = datetime.strptime(date_str.strip(), '%Y-%m-%d %H:%M:%S')
        except ValueError:
            try:
                date = datetime.strptime(date_str.strip(), '%Y-%m-%d')
            except ValueError:
                date = datetime.now()

        records.append({
            'date': date,
            'merchant': merchant or description or '未知商户',
            'amount': amount,
            'category_raw': category,
            'source': '支付宝',
            'description': description,
        })
    return records


def parse_wechat_csv(file_content: str) -> list[dict]:
    """解析微信 CSV 账单"""
    records = []
    lines = file_content.strip().split('\n')
    # 微信 CSV 格式：
    # 交易时间, 交易类型, 交易对方, 商品, 收/支, 金额(元), 支付方式, 当前状态, 交易订单号, 商户单号, 备注
    for line in lines:
        if not line.strip():
            continue
        if '交易时间' in line:
            continue
        parts = [p.strip().strip('"') for p in line.split(',')]
        if len(parts) < 6:
            continue
        date_str = parts[0]
        txn_type = parts[1] if len(parts) > 1 else ''
        merchant = parts[2] if len(parts) > 2 else ''
        description = parts[3] if len(parts) > 3 else ''
        direction = parts[4] if len(parts) > 4 else ''
        amount_str = parts[5] if len(parts) > 5 else ''

        if not direction or '支出' not in direction:
            continue

        try:
            amount = float(amount_str.replace('¥', '').replace(',', '').strip())
        except (ValueError, AttributeError):
            continue

        try:
            date = datetime.strptime(date_str.strip(), '%Y-%m-%d %H:%M:%S')
        except ValueError:
            try:
                date = datetime.strptime(date_str.strip(), '%Y-%m-%d')
            except ValueError:
                date = datetime.now()

        records.append({
            'date': date,
            'merchant': merchant or description or '未知商户',
            'amount': amount,
            'category_raw': txn_type,
            'source': '微信',
            'description': description,
        })
    return records


def auto_detect_and_parse(file_content: str) -> list[dict]:
    """自动检测支付宝还是微信格式"""
    if '交易时间, 交易分类, 交易对方' in file_content:
        return parse_alipay_csv(file_content)
    elif '交易时间, 交易类型, 交易对方' in file_content:
        return parse_wechat_csv(file_content)
    else:
        # 试试通用格式
        return parse_alipay_csv(file_content)


# ============================================================
# 路由
# ============================================================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload():
    """上传 CSV/Excel 文件"""
    if 'file' not in request.files:
        return jsonify({'error': '请选择文件'}), 400

    f = request.files['file']
    if not f.filename:
        return jsonify({'error': '文件名为空'}), 400

    filename = secure_filename(f.filename)
    content = f.read()

    # 判断文件类型
    if filename.endswith('.csv'):
        try:
            text = content.decode('utf-8-sig')
        except UnicodeDecodeError:
            text = content.decode('gbk', errors='replace')

        records = auto_detect_and_parse(text)

    elif filename.endswith(('.xlsx', '.xls')):
        # Excel 通用解析
        df = pd.read_excel(io.BytesIO(content))
        records = []
        for _, row in df.iterrows():
            date = row.iloc[0] if len(row) > 0 else ''
            merchant = str(row.iloc[2]) if len(row) > 2 else ''
            amount_str = str(row.iloc[5]) if len(row) > 5 else ''
            direction = str(row.iloc[4]) if len(row) > 4 else ''
            try:
                amount = float(amount_str.replace('¥', ''))
            except (ValueError, AttributeError):
                continue
            if '支出' not in direction:
                continue
            records.append({
                'date': pd.to_datetime(date) if not isinstance(date, str) else datetime.now(),
                'merchant': merchant,
                'amount': amount,
                'category_raw': '',
                'source': 'Excel',
                'description': '',
            })
    else:
        return jsonify({'error': '仅支持 CSV 和 Excel 文件'}), 400

    if not records:
        return jsonify({'error': '未能解析出有效支出记录，请确认文件格式'}), 400

    # 归类
    for r in records:
        r['category'] = classify_by_rules(r['merchant'], r.get('category_raw', ''))

    # 按月份分组
    monthly = defaultdict(list)
    for r in records:
        month_key = r['date'].strftime('%Y-%m') if hasattr(r['date'], 'strftime') else '未知'
        monthly[month_key].append(r)

    # 生成报表
    report = {}
    for month, items in sorted(monthly.items()):
        total = sum(it['amount'] for it in items)
        by_category = defaultdict(float)
        by_source = defaultdict(float)
        for it in items:
            by_category[it['category']] += it['amount']
            by_source[it['source']] += it['amount']

        report[month] = {
            'total': round(total, 2),
            'count': len(items),
            'by_category': dict(sorted(by_category.items(), key=lambda x: -x[1])),
            'by_source': dict(by_source),
        }

    # 生成唯一 ID 存 session
    data_id = hashlib.md5(str(datetime.now().timestamp()).encode()).hexdigest()[:8]
    session['last_data_id'] = data_id

    return jsonify({
        'data_id': data_id,
        'total_records': len(records),
        'records': records,
        'report': report,
    })


@app.route('/api/classify_all', methods=['POST'])
def classify_all():
    """用 LLM 对未分类记录进行再分类（可选）"""
    data = request.json
    records = data.get('records', [])
    api_key = data.get('api_key', '')

    if not api_key:
        # 无 key 时返回未分类列表
        unclassified = [r for r in records if r.get('category') == '未分类']
        return jsonify({'unclassified_count': len(unclassified), 'unclassified': unclassified})

    # 有 key 时调用 LLM 分类
    import urllib.request
    unclassified = [r for r in records if r.get('category') == '未分类']
    if not unclassified:
        return jsonify({'status': 'ok', 'classified': 0, 'records': records})

    # 构建 prompt
    merchants = [{'merchant': r['merchant'], 'amount': r['amount']} for r in unclassified]
    prompt = f"""请为以下每笔交易分配一个消费类别。可用类别：
餐饮·饮品, 餐饮·外卖, 餐饮·正餐, 餐饮·零售,
交通·打车, 交通·公交地铁, 交通·城际, 交通·机票, 交通·加油充电,
购物·电商, 购物·日用, 购物·服饰,
居住·水电物业, 居住·租房,
通信·话费,
金融·房贷, 金融·信用卡,
娱乐·订阅, 娱乐·游戏, 娱乐·文娱,
医疗健康,
教育·学习,
社交·转账,
或者你判断更合适的类别。

只返回 JSON 数组，格式：[{{"merchant":"xxx","category":"xxx"}}]

交易列表：
{json.dumps(merchants, ensure_ascii=False)}"""

    req = urllib.request.Request(
        'https://api.openai.com/v1/chat/completions',
        data=json.dumps({
            'model': 'gpt-4o-mini',
            'messages': [{'role': 'user', 'content': prompt}],
            'temperature': 0.1,
        }).encode(),
        headers={
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {api_key}',
        }
    )
    try:
        resp = urllib.request.urlopen(req, timeout=30)
        result = json.loads(resp.read())
        text = result['choices'][0]['message']['content']
        # 提取 JSON
        import re as re2
        json_match = re2.search(r'\[.*\]', text, re2.DOTALL)
        if json_match:
            classifications = json.loads(json_match.group())
            merchant_map = {c['merchant']: c['category'] for c in classifications}
            for r in records:
                if r['merchant'] in merchant_map:
                    r['category'] = merchant_map[r['merchant']]
        return jsonify({'status': 'ok', 'classified': len(unclassified), 'records': records})
    except Exception as e:
        return jsonify({'status': 'error', 'error': str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5111, debug=True)
