# -*- coding: utf-8 -*-
"""
生成 pokedex_entries.json（图鉴说明数据，去重压缩版）。
供 GitHub Actions 自动运行：监听 pokedex xlsx 更新后重新生成。

只依赖 openpyxl（无 pypinyin/opencc/术语表依赖），因为：
  - 说明文本（desc/descTrad）直接来自 pokedex xlsx 本身
  - 宝可梦名字/拼音/世代在 index.html 内嵌的 POKEDEX_INDEX 里，不在这里
输出结构：
  { v:[{name,gen,simp,order}], descs:[{d,t,z}], pokes:{"id_form":[[descIdx,verIdx],...]} }
"""
import openpyxl, re, json, os, sys, glob

# 自动发现 pokedex xlsx：优先含"已更新"，否则任意 pokedex 开头的 xlsx
def find_dex_xlsx():
    env = os.environ.get('DEX_XLSX')
    if env and os.path.exists(env):
        return env
    # 当前目录下找 pokedex 开头的 xlsx
    cands = sorted(glob.glob('pokedex*.xlsx'))
    # 忽略 Excel 临时锁文件（~$ 开头）
    cands = [c for c in cands if not os.path.basename(c).startswith('~$')]
    if not cands:
        return None
    # 按文件名中的版本号排序，选最新版本（如 pokedex 1.1.xlsx 优先于 pokedex 1.0_已更新.xlsx）
    def ver_key(path):
        name = os.path.basename(path)
        m = re.search(r'(\d+)\.(\d+)', name)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        return (0, 0)
    cands.sort(key=ver_key, reverse=True)
    return cands[0]

DEX_PATH = find_dex_xlsx()
OUT_PATH = os.environ.get('OUT_JSON', 'pokedex_entries.json')

# 版本 -> 世代 映射（繁体版本名 -> 世代号）
VERSION_GEN = {
    '紅': 1, '綠': 1, '藍': 1, '皮卡丘': 1,
    '金': 2, '銀': 2, '水晶版': 2,
    '紅寶石': 3, '藍寶石': 3, '綠寶石': 3, '火紅': 3, '葉綠': 3,
    '鑽石': 4, '珍珠': 4, '白金': 4, '心金': 4, '魂銀': 4,
    '黑': 5, '白': 5, '黑２': 5, '白２': 5,
    'Ｘ': 6, 'Ｙ': 6, '歐米加紅寶石': 6, '阿爾法藍寶石': 6,
    '太陽': 7, '月亮': 7, '究極之日': 7, '究極之月': 7,
    "Let's Go！皮卡丘": 7, "Let's Go！伊布": 7,
    '劍': 8, '盾': 8, '晶燦鑽石': 8, '明亮珍珠': 8, '傳說 阿爾宙斯': 8,
    '朱': 9, '紫': 9, '傳說 Z-A': 9,
    # 旁支系列（非正代，用世代号 10 标记，前端归入「旁支系列」分组）
    'Pokopia': 10, 'Masters': 10, 'Go': 10,
}

# 版本显示顺序（游戏发布时间序）
VERSION_ORDER = [
    '紅', '綠', '藍', '皮卡丘',
    '金', '銀', '水晶版',
    '紅寶石', '藍寶石', '綠寶石', '火紅', '葉綠',
    '鑽石', '珍珠', '白金', '心金', '魂銀',
    '黑', '白', '黑２', '白２',
    'Ｘ', 'Ｙ', '歐米加紅寶石', '阿爾法藍寶石',
    '太陽', '月亮', '究極之日', '究極之月',
    "Let's Go！皮卡丘", "Let's Go！伊布",
    '劍', '盾', '晶燦鑽石', '明亮珍珠', '傳說 阿爾宙斯',
    '朱', '紫', '傳說 Z-A',
    'Pokopia', 'Masters', 'Go',
]
VERSION_ORDER_MAP = {v: i for i, v in enumerate(VERSION_ORDER)}

# 繁体 -> 简体（版本名），手动映射（避免依赖 opencc）
VER_SIMP = {
    '紅': '红', '綠': '绿', '藍': '蓝', '皮卡丘': '皮卡丘',
    '金': '金', '銀': '银', '水晶版': '水晶版',
    '紅寶石': '红宝石', '藍寶石': '蓝宝石', '綠寶石': '绿宝石',
    '火紅': '火红', '葉綠': '叶绿',
    '鑽石': '钻石', '珍珠': '珍珠', '白金': '白金',
    '心金': '心金', '魂銀': '魂银',
    '黑': '黑', '白': '白', '黑２': '黑２', '白２': '白２',
    'Ｘ': 'Ｘ', 'Ｙ': 'Ｙ',
    '歐米加紅寶石': '欧米加红宝石', '阿爾法藍寶石': '阿尔法蓝宝石',
    '太陽': '太阳', '月亮': '月亮',
    '究極之日': '究极之日', '究極之月': '究极之月',
    "Let's Go！皮卡丘": "Let's Go！皮卡丘", "Let's Go！伊布": "Let's Go！伊布",
    '劍': '剑', '盾': '盾',
    '晶燦鑽石': '晶灿钻石', '明亮珍珠': '明亮珍珠', '傳說 阿爾宙斯': '传说 阿尔宙斯',
    '朱': '朱', '紫': '紫', '傳說 Z-A': '传说 Z-A',
    'Pokopia': 'Pokopia', 'Masters': 'Masters', 'Go': 'Go',
}

# 暂译类型细分（与 index.html 里 JS 的 zantyiType 保持一致）
# 新版 xlsx 将「是否暂译」拆为「简中情况」+「繁中情况」两列，各自标注：
#   官译 / 暂译 / 卡牌官译 / 博物展官译
# 依据简繁两列的标注组合判断：
#   官译+官译           -> game（游戏官译）
#   暂译+暂译           -> fan（非官译）
#   卡牌官译/博物展官译（简繁都是其他官译） -> card
#   卡牌官译/博物展官译（仅简）+ 暂译         -> cardSimplOnly（仅简官）
#   暂译 + 卡牌官译/博物展官译（仅繁）         -> cardTradOnly（仅繁官）
OTHER_OFFICIAL = ('卡牌官译', '博物展官译')

def zantyi_type(jian, fan):
    jian = (jian or '').strip()
    fan = (fan or '').strip()
    j_game = jian == '官译'
    f_game = fan == '官译'
    j_fan = jian == '暂译'
    f_fan = fan == '暂译'
    j_card = jian in OTHER_OFFICIAL
    f_card = fan in OTHER_OFFICIAL
    if j_game and f_game:
        return 'game'
    if j_fan and f_fan:
        return 'fan'
    if j_card and f_card:
        return 'card'
    if j_card and f_fan:
        return 'cardSimplOnly'
    if j_fan and f_card:
        return 'cardTradOnly'
    return 'other'


def main():
    if not os.path.exists(DEX_PATH):
        print(f'错误：找不到 {DEX_PATH}')
        sys.exit(1)

    wb = openpyxl.load_workbook(DEX_PATH, read_only=True, data_only=True)
    ws = wb['pokedex_explained']
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)  # 跳过表头
    # 列: 0编号_形态_说明 1中文名_形态名 2简体说明 3繁体说明 4应用版本 5简中情况 6繁中情况
    # 兼容旧版：若表头第6列是「是否暂译」（单列合并），则按旧逻辑处理

    # 先收集所有版本，确定版本列表
    ver_names = set()
    desc_list = []
    desc_key_to_idx = {}
    poke_entries = {}

    for row in rows:
        if row is None or row[0] is None:
            continue
        code = str(row[0]).strip()
        m = re.match(r'^(\d{4})_(\d+)_(\d+)$', code)
        if not m:
            continue
        pid = int(m.group(1))
        formNo = m.group(2)
        desc_s = str(row[2]).strip() if len(row) > 2 and row[2] else ''
        desc_t = str(row[3]).strip() if len(row) > 3 and row[3] else ''
        ver = str(row[4]).strip() if len(row) > 4 and row[4] else ''
        jian = str(row[5]).strip() if len(row) > 5 and row[5] else ''
        fan = str(row[6]).strip() if len(row) > 6 and row[6] else ''

        key = f'{pid}_{formNo}'
        z = zantyi_type(jian, fan)
        dk = (desc_s, desc_t, z)
        if dk not in desc_key_to_idx:
            desc_key_to_idx[dk] = len(desc_list)
            desc_list.append({'d': desc_s, 't': desc_t, 'z': z})
        didx = desc_key_to_idx[dk]

        for v in ver.split(','):
            v = v.strip()
            if not v:
                continue
            ver_names.add(v)
            if key not in poke_entries:
                poke_entries[key] = []
            poke_entries[key].append([didx, v])  # 暂存版本名，后面转索引

    # 构建版本列表（按 VERSION_ORDER 排序）
    version_list = []
    for v in ver_names:
        version_list.append({
            'name': v,
            'gen': VERSION_GEN.get(v, 0),
            'simp': VER_SIMP.get(v, v),
            'order': VERSION_ORDER_MAP.get(v, 9999),
        })
    version_list.sort(key=lambda x: (x['order'], x['name']))

    # 版本名 -> 索引
    ver_to_idx = {v['name']: i for i, v in enumerate(version_list)}

    # 把 poke_entries 里的版本名转成索引
    pokes_out = {}
    for key, entries in poke_entries.items():
        pokes_out[key] = [[didx, ver_to_idx[v]] for didx, v in entries]

    out = {'v': version_list, 'descs': desc_list, 'pokes': pokes_out}

    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

    print(f'生成完成: {OUT_PATH}')
    print(f'版本数: {len(version_list)}')
    print(f'去重说明数: {len(desc_list)}')
    print(f'宝可梦单元: {len(pokes_out)}')
    print(f'总 entries 引用: {sum(len(v) for v in pokes_out.values())}')
    print(f'文件大小: {os.path.getsize(OUT_PATH)} bytes')


if __name__ == '__main__':
    main()
