---
name: daily-work-summary
description: >
  从金山文档 cic/活动记录/2026活动记录.xlsx 读取工作记录，按 PMS工作分类 sheet 的分类进行归类统计并输出。
  单天（如"昨天的工作"）→ 对话中直接输出 PMS 分类占比；
  多天（如"上周四到昨天的工作"）→ 生成站会格式周报，保存到 周报/站会/ 目录。
  触发词：帮我总结一下昨天的工作、总结今天/前天/某天/上周X到昨天/本周X到昨天的工
  作、daily summary、工作总结、站会周报
agent_created: true
version: 2.0
---

# 工作总结 Skill（单天 & 多天）

## 目标
读取金山文档中的活动记录，根据日期范围自动选择输出模式：
- **单天模式**：按 PMS 工作分类归类，输出分类占比（对话中直接展示）
- **多天模式**：生成站会格式周报，保存到 `周报/站会/` 目录

## 触发条件

### 单天模式触发
用户说类似以下任意一种：
- "帮我总结一下昨天的工作"
- "总结一下今天/前天/某天的工作"
- "昨天做了什么"
- "6月XX日的工作总结"
- "昨天的工作分类占比"

### 多天模式触发（新增）
用户说类似以下任意一种：
- "总结上周四到昨天的工作"
- "总结本周二到昨天的工作"
- "总结 X月X日到X月X日 的工作"
- "帮我生成站会周报"
- 任何包含**多天范围**的总结请求（含"到"、"~"、"这几天"等）

## 模式判断
- 用户请求中只涉及**一个日期** → **单天模式**
- 用户请求中涉及**日期范围**（如"上周X到昨天"、"X号到X号"） → **多天模式**
- 用户说"生成站会周报"但没指定日期 → 多天模式，默认覆盖**上周一到昨天**

## 执行步骤（两种模式共用前三步）

### 第一步：确定目标日期

单天：
- "昨天" → 当前日期 - 1 天
- "今天" → 当前日期
- 指定日期 → 直接使用

多天：
- "上周四到昨天" → 上周四（当前日期往回推到上周四）到昨天
- "本周二到昨天" → 本周二到昨天
- "X月X日到X月X日" → 直接解析两个日期
- 计算时需要排除周末（周六、周日），但读取时仍把周末包含在目标日期集合中以备有记录

日期格式：`2026-6-{日}` （Kdocs 存储格式，不补零）

### 第二步：读取金山文档数据

文件 ID：`hk1ycoMKQ9MwqbHPRm6UrxBiC2WaaraJG`

**⚠️ 重要：必须使用 `sheet get-range-data` 读取，禁止使用 `drive read-file`！**
- `drive read-file` 存在数据缓存/截断问题，用户新保存的数据经常读不到（2026-08-06 踩坑）
- `sheet get-range-data` 每次实时读取最新数据
- 需要先用 `sheet get-sheets-info` 获取每个 sheet 的 `sheetId`

**多天跨月处理**：如果日期范围跨了月份（如 6/28~7/3），需要分别读取两个月份的 sheet。

第一步：获取所有 sheet 的 ID：
```bash
kdocs-cli sheet get-sheets-info file_id="hk1ycoMKQ9MwqbHPRm6UrxBiC2WaaraJG" --silent > /tmp/kdocs_sheets.json
```
从返回的 `detail.sheetsInfo` 中按 `sheetName` 找到对应 sheet 的 `sheetId`（如 202608 是 35）。

第二步：读取 工作分类 sheet（获取分类描述）：
```bash
kdocs-cli sheet get-range-data file_id="hk1ycoMKQ9MwqbHPRm6UrxBiC2WaaraJG" sheetId=<工作分类的sheetId> range='{"rowFrom":0,"rowTo":15,"colFrom":0,"colTo":1}' --silent > /tmp/kdocs_pms.json
```

第三步：读取对应月份的 sheet（根据目标日期自动推算 sheet 名称和行号范围）：

**sheet 命名规则**：`YYYYMM`，例如：
- 2026年6月 → `202606`
- 2026年7月 → `202607`
- 2027年1月 → `202701`
- 以此类推：`f"{year}{month:02d}"`

**行号估算**：
- 每月约 5~8 行/天，每月 1 日从约 row 0 开始
- 目标日期的大致行号 = (日 - 1) * 7
- 单天：目标行 - 20 到目标行 + 40
- 多天：第一天行号 - 20 到最后一天行号 + 40

```bash
# sheetId 根据目标月份从 sheets_info 中获取
kdocs-cli sheet get-range-data file_id="hk1ycoMKQ9MwqbHPRm6UrxBiC2WaaraJG" sheetId=<月份sheetId> range='{"rowFrom":<start>,"rowTo":<end>,"colFrom":0,"colTo":5}' --silent > /tmp/kdocs_activity.json
```

**解析 JSON 结构注意**：`sheet get-range-data` 返回结构为 `data.detail.rangeData`（单元格数组），每个单元格含 `originRow`/`originCol`/`rowFrom`/`rowTo`/`colFrom`/`colTo`/`cellText` 等字段。

### 第三步：解析数据（必须正确处理合并单元格）

**⚠️ 重要：金山文档 xlsx 存在大量合并单元格**
- 日期、开始时间、类型、番茄数列经常跨行合并
- 合并区域中只有左上角单元格有值，其余子单元格为空
- **必须** 通过 `rowFrom`/`rowTo`/`colFrom`/`colTo` 字段识别合并范围
- 并将左上角的值填充到所有被合并的行

解析模板（Python）：
```python
import json
from collections import defaultdict

def parse_kdocs_with_merge(json_path, target_dates):
    with open(json_path, 'r') as f:
        data = json.load(f)
    
    # sheet get-range-data 结构：data['detail']['rangeData']
    # drive read-file 结构：data['content']['range_data'][0]['detail']['rangeData']（勿用，有缓存问题）
    cells = data['detail']['rangeData'] if 'detail' in data and isinstance(data.get('detail'), dict) else data['content']['range_data'][0]['detail']['rangeData']
    if not cells:
        return []
    
    raw = {}    # (row, col) -> {'value': ..., 'rf': ..., 'rt': ..., 'cf': ..., 'ct': ...}
    merge = {}  # (row, col) -> (rf, rt, cf, ct)  被合并到的源区域
    
    for c in cells:
        r, col = c['originRow'], c['originCol']
        v = c.get('understandableType', {}).get('value', '')
        txt = c.get('cellText', '')
        val = str(v) if v else str(txt) if txt else ''
        
        rf, rt = c.get('rowFrom', r), c.get('rowTo', r)
        cf, ct = c.get('colFrom', col), c.get('colTo', col)
        
        raw[(r, col)] = {'value': val, 'rf': rf, 'rt': rt, 'cf': cf, 'ct': ct}
        
        # 记录合并区域：将每个子单元格映射到左上角
        if rf != rt or cf != ct:
            for mr in range(rf, rt + 1):
                for mc in range(cf, ct + 1):
                    merge[(mr, mc)] = (rf, rt, cf, ct)
    
    all_rows = sorted(set(k[0] for k in raw.keys()) | set(k[0] for k in merge.keys()))
    
    records = []
    for r in all_rows:
        row = {}
        for col in range(6):
            # 优先从合并区域取左上角的值
            if (r, col) in merge:
                rf, rt, cf, ct = merge[(r, col)]
                key = (rf, cf)
                if key in raw:
                    row[col] = raw[key]['value'].strip()
                    continue
            # 再取原始值
            key = (r, col)
            if key in raw:
                row[col] = raw[key]['value'].strip()
            else:
                row[col] = ''
        
        d = row.get(0, '')
        if d not in target_dates:
            continue
        
        activity = row.get(3, '').strip()
        if not activity:
            continue
        
        ts = row.get(4, '').strip() or '0'
        try:
            tomato = float(ts) if ts else 0.0
        except:
            tomato = 0.0
        
        records.append({
            'date': d,
            'time': row.get(1, '').strip(),
            'type': row.get(2, '').strip(),
            'activity': activity,
            'tomato': tomato,
            'description': row.get(5, '').strip()
        })
    
    return records
```

### 第四步-A：单天模式 — 按 PMS 分类汇总输出

读取 工作分类 sheet：col 0 = 分类名称，col 1 = 分类描述。

输出格式：
```
分类名称：分类描述 XX%
工作内容1 工作内容2 工作内容3
```

按番茄数占比降序排列各分类。总番茄数 = 当日所有记录番茄数之和。
注意：如果活动记录中的类型与 PMS 分类名称一致，"项目管理类" 等模糊分类需根据活动内容判断实际归属。

**⚠️ 分类内工作项去重合并（2026-09-22 用户明确要求）**：
- 同一分类下**重复出现或相似度超过 80% 的工作项必须合并为一个**（如"信创改造"出现多次只写一次、"AI辅助研发调研"多段合并、"慢sql治理"多次合并），番茄数仍按全部记录累加计算占比
- 相似度不够（内容实质不同的工作项）保持独立，不合并

### 第四步-B：多天模式 — 生成站会格式周报（新增）

从记录中按天整理，并按以下四个模块分类每组内容，生成 Markdown 文件：

**模块分类规则**：
- **线上问题排查&修复**：协助其他域排查线上问题、bug 排查及修复、线上问题修复方案讨论、系统运行异常原因排查及恢复等。特点：单次定位+修复，修完即结束，不涉及系统层面的机制性改进
- **需求开发**：研发类的编码、功能开发、发布
- **技术方案输出**：技术方案撰写、技术调研、方案讨论（产出型）、概要设计、技改类方案、代码 review、方案 review 等 review 类工作
- **需求跟进**：需求评审、系分评审、设计评审、方案评审、测试点评审等评审类工作、需求讨论、排期沟通、接口对齐
- **项目管理**：项目管理类、站会、周会、周报、进度跟进、PMS 审核、新人入职、复盘、需求排期、PMS 相关工作
- **线上运维**：生产环境配置调整、监控配置、任务调度配置等日常运维操作
- **信创改造**：凡涉及信创（信创改造、信创实施方案、信创沟通会、信创改造工作沟通等）的内容，单独列为一节，不混入其他模块
- **小角色大权限治理**：凡涉及"小角色大权限"治理的内容（方案宣导/培训、角色&权限&数据库矩阵自查、敏感数据过度公开自查、互联网资产安全风险挖掘、自建 nginx 排查等），单独列为一节，不混入其他模块
- **其他**：无法识别明确分类的内容，统一放到此分类下

**文件命名**：`周报/站会/{start_date}{end_date}.md`，如 `20260625-0629.md`（不含年份的月日）。

**文件格式**（参考已有站会周报）：
```markdown
# 工作总结 YYYY-MM-DD ~ YYYY-MM-DD

---

### 需求开发

1. **活动名称**
   - 详细描述点1；
   - 详细描述点2；

---

### 技术方案输出

1. **方案名称**
   - 要点1
   - 要点2

---

### 需求跟进

1. **事项名称** — 简要结论

---

### 项目管理

1. **事项名称**
```

**⚠️ 同一工作项的合并与日期处理（2026-09-09 用户明确要求）**：
- **不保留每条工作项后面的日期标注**（不要再写 `（9/2）` 这类后缀）
- **跨天重复或高度相似的工作项必须合并成一条**，把多天的要点作为该条目的多个描述点：
  - 同一治理/调研类工作（如慢 SQL 治理、数据安全自查）分散在多天 → 合并为一条，按天罗列要点
  - 同一需求的多轮讨论（如渠道大批量合并支付需求讨论）→ 合并为一条，按推进顺序罗列结论
  - 同类事务性工作（如各站会、PMS 提交技改申请/需求审核/立项）→ 合并为一条，用「、」连接名称
  - 关联紧密的同一主题（如 AI 辅助研发流程宣导 + spec 初始化流程答疑）→ 合并为一条
- 不同需求的评审/系分评审仍然各自独立成条，不做合并

**归类注意事项**：
- **每个大类内的条目按首次发生日期从早到晚排序**，不要按主题或重要程度排列
- **站会周报中不包含番茄数**，只写活动名称和详细描述
- 代码 review、方案 review 等 review 类工作 → 归入技术方案输出
- 系分评审、设计评审、方案评审、需求评审、测试点评审等评审类工作 → 归入需求跟进
- 线上问题修复（bug 修复、故障修复方案、问题治理讨论）→ 归入线上问题排查&修复
- "方案沟通"中产出了明确的方案结论或技术方向 → 归入技术方案输出
- "方案沟通"中仅涉及排期、接口对齐、上下游协调 → 归入需求跟进
- "财务中台需求进度沟通会" → 项目管理类（不是沟通类）
- 协助其他域排查问题、bug排查修复、系统运行异常恢复 → 线上问题排查&修复
- 生产环境配置/监控/任务调度调整 → 线上运维
- 涉及需求排期、PMS、站会的工作项 → 归入项目管理
- 无法识别明确分类的内容 → 统一归入其他
- 涉及信创改造/信创方案/信创沟通会/信创工作沟通 → 单独归入信创改造节
- 涉及"小角色大权限"治理（方案宣导/培训、权限矩阵自查、敏感数据过度公开自查、资产安全风险挖掘、自建 nginx 排查等）→ 单独归入「小角色大权限治理」节（与信创改造同为独立大事项）
- 没有详细描述的活动，仅列出名称不展开
- （已废弃）不再在每条后面标注日期

### 第五步：保存文件（仅多天模式）

文件保存到 `周报/站会/` 目录后，使用 `present_files` 工具展示给用户。

## 注意事项
- **sheet 名称**：根据目标日期动态计算，格式 `YYYYMM`（月份补零），例如 `202606`、`202607`、`202701`
- **跨月处理**：如果日期范围跨越不同月份，分别读取多个 sheet 再合并 records
- **日期格式**：Kdocs 存储格式为 `YYYY-M-D`（月和日不补零），例如 `2026-6-29`，不是 `2026-06-29`
- **⚠️ 日期以表格记录为准（2026-10-14 用户明确要求："按照我表格里面的日志来，不要自己猜"）**：不要根据周末/工作日、块的出现时间等信息推断或"校正"日期；调休工作日（周末上班）同样是工作日，按表格里的日期原样统计与展示
- 如果数据为空，先扩大行号范围重试，再询问用户是否已在文档中填写
- 始终使用合并单元格解析逻辑，不要跳过
- 多天模式输出文件后，同时展示简要的日汇总表格（每天番茄数 + 核心产出1句话）
