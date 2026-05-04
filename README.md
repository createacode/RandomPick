# RandomPick
一个完成度很高的随机抽选工具，从一堆人里面抽取一个或几个，提取微秒（小数点后六位）作为随机数种子。

![界面截图-抽选](https://github.com/createacode/RandomPick/blob/main/tool%20(2).png)

![界面截图-数据统计](https://github.com/createacode/RandomPick/blob/main/tool%20(4).png)


# 随机抽选工具（RandomPick）开发文档

## 1. 项目概述

### 1.1 项目名称
随机抽选工具（RandomPick）

### 1.2 项目简介
本软件是一款基于 PyQt5 开发的桌面应用程序，用于从预定义的人员列表中公平随机抽取一人或多人，适用于抽奖、随机点名、活动选拔等场景。系统采用**双随机机制**（映射随机 + 微秒级时间种子），保证抽选的不可预测性和均匀性。

### 1.3 主要特性
- **人员管理**：支持增删改查、头像上传裁剪、Excel 批量导入（含单元格内嵌图片）。
- **随机抽取**：基于微秒映射 + `secrets` 模块的安全随机算法，支持不放回多人抽取。
- **动画交互**：抽取过程卡片高亮轮播，速度分为前、中、后三档可调，最终定位中奖者。
- **结果展示**：置顶弹窗显示中奖者头像和姓名，非模态且自动关窗可配置，弹窗始终居中于主窗口。
- **数据统计**：记录每个人员的中选次数、概率，提供表格排序和柱状图。
- **系统集成**：单实例运行、系统托盘、窗口布局持久化、日志记录。
- **外观配置**：窗口大小、分割条位置、字体时钟、结果显示方式、动画速度、自动关闭时间等均可按需调整。

### 1.4 开发环境
- **语言**：Python 3.13
- **GUI框架**：PyQt5 5.15.x
- **第三方库**：
  - Pillow（图像处理）
  - openpyxl（Excel 导入导出）
  - matplotlib（数据可视化，支持中文）
  - pyinstaller（打包）
- **运行平台**：Windows 10/11（理论上支持 Linux/macOS，但未深度测试）

---

## 2. 系统架构

### 2.1 整体结构
采用 **MVC 模式**，将数据模型（`PersonManager`）、视图（PyQt5 界面）和控制逻辑（`MainWindow`、`RandomPicker`）分离。

```text
+----------------+        +------------------+
|   PersonManager |------> |  JSON 文件存储    |
|   (人员数据)     |        | persons.json     |
+----------------+        +------------------+
         |
         v
+----------------+        +------------------+
|  RandomPicker  |        |  Logger          |
|  (抽取算法)     |------> |  (日志记录)       |
+----------------+        +------------------+
         |
         v
+------------------------------------------------+
|               MainWindow (主界面)               |
|  - 卡片展示区域                                 |
|  - 抽取控制栏（人数、次数、开始按钮）             |
|  - 右侧面板（实时时钟、本次结果、多组结果、过程日志）|
|  - 工具栏（人员维护、数据统计、显示设置、抽取记录、关于）|
+------------------------------------------------+
         |                  |
         v                  v
+----------------+  +---------------------+
| PersonManage   |  | StatsDialog         |
| Dialog         |  | (统计对话框)         |
+----------------+  +---------------------+
```

### 2.2 核心模块说明

| 模块 | 文件位置 | 职责 |
|------|----------|------|
| 日志系统 | `Logger` 类 | 生成日志文件，记录抽取关键步骤和用户操作 |
| 人员模型 | `Person`, `PersonManager` | 人员数据的增删改查、JSON 持久化、Excel 导入 |
| 随机抽取 | `RandomPicker` | 实现映射随机和微秒分块算法，返回中选人员 ID 列表 |
| 头像裁剪 | `CropDialog` | 提供鼠标拖拽裁剪头像功能 |
| 统计 | `StatsDialog` | 展示中选次数、概率，绘制 matplotlib 柱状图 |
| 结果显示 | `ResultDialog` | 半透明浮层显示头像和姓名，非模态，支持自动关闭，绝对居中 |
| 主窗口 | `MainWindow` | 卡片布局、动画轮播、右侧面板交互、配置持久化 |
| 显示设置 | `SettingsDialog` | 控制各项显示参数、自动关闭时间、动画速度、右侧组件显隐 |
| 系统托盘 | `QSystemTrayIcon` | 最小化到托盘，双击恢复，右键退出 |
| 单实例 | `QLocalServer` | 确保同时只有一个进程运行，第二个进程会激活已有窗口 |

---

## 3. 功能详解

### 3.1 人员管理
- **手动录入**：通过“新增人员”对话框填写昵称（必填）、真实姓名（可选）和头像。
- **头像处理**：
  - 选择本地图片后进入裁剪界面（缩放适应窗口，鼠标拖拽矩形区域，可确认/取消）。
  - 裁剪后自动保存为 256×256 的 PNG 到 `person/photo` 目录，文件名格式为 `{id}_{时间戳}.png`。
- **编辑删除**：表格中每行提供“编辑”“删除”按钮，删除时同时移除关联的头像文件。
- **参与抽取标记**：表格中的复选框可单独控制某人是否参与抽取，状态保存在 `persons.json`。
- **搜索筛选**：根据昵称或真实姓名实时过滤表格行。
- **批量导入**：
  - 支持 Excel（.xlsx/.xls）文件，必须包含工作表“人员导入”。
  - 模板结构：列依次为“序号”、“绰号*”、“姓名”、“头像”（其中“*”表示必填）。
  - 两种模式：
    - **直接导入（带照片）**：如果“头像”列填写了文件名且文件存在，或单元格内嵌图片，则自动复制到 `person/photo`。
    - **先导入无照片名单（稍后匹配头像）**：先导入人员信息，随后打开临时文件夹，用户按照 `img (1).jpg` 的命名规则（序号对应）放入照片，点击“匹配头像”自动关联。
  - 临时文件夹在匹配完成后自动删除。
  - 导入模板可通过“下载模板”按钮生成。

### 3.2 随机抽取
- **算法原理**：
  1. 获取当前参与抽取的人员列表。
  2. 使用 `secrets` 模块随机打乱人员顺序，为每个人员分配一个 1~N 的序号（映射表）。
  3. 取当前时间的微秒值（1 ~ 999999），按人数 N 分成若干区间，舍弃无法均分的尾部余数。
  4. 计算微秒值所在区间，得到该次抽取的序号，再通过映射表找到对应人员。
  5. 如果不放回且需要抽取多人，则从剩余人员中重新生成映射表，重复步骤 3~4，直到抽完。
- **日志记录**：每次抽取的映射表、微秒值、分组计算过程、最终中选者均写入日志文件和右侧过程面板。
- **动画**：抽取时所有参与卡片按随机顺序轮流高亮，速度分为前段、中段、后段（可分别设置毫秒/步），最终停留在中奖者卡片上。

### 3.3 数据统计
- **详细数据**：表格展示每个人员的 ID、昵称、真实姓名、中选次数、中选概率。点击表头可排序，所有单元格居中对齐。
- **柱状图**：使用 matplotlib 绘制概率分布图，包含 X 轴标签（昵称）、Y 轴数值、标题，自动处理中文显示，柱上数字不重叠。
- **清空记录**：删除所有历史抽选记录（不影响人员数据）。

### 3.4 界面控制与配置
- **抽取控制**：设置“抽取人数”和“抽取次数”（支持 1~100 次），点击“开始抽取”启动流程。
- **右侧面板**：
  - 实时时钟（显示到毫秒，大字体 Consolas）。
  - 本次抽取结果（按配置的“结果显示时机”即时或延后显示）。
  - 多组结果：垂直排列每次抽取的中奖名单（每行最多 5 个，居中显示）。
  - 抽取过程：详细展示映射表、微秒计算、中奖者，每条记录带时间戳。延后模式下只有箭头以下部分延迟显示。
- **工具栏**：人员维护、数据统计、显示设置、抽取记录（右侧面板整体显隐）、关于。
- **系统托盘**：关闭主窗口时最小化到托盘，双击托盘图标恢复，右键菜单“退出”彻底关闭。

### 3.5 配置持久化
- 所有配置保存在 `config/settings.json`，包括：
  - 窗口大小和位置（`window_geometry`）
  - 主水平分割器和右侧垂直分割器的尺寸
  - 抽取人数、抽取次数
  - 显示设置（名片/上次结果/抽中结果是否显示真实姓名、标题后缀、右侧面板组件显隐、弹窗自动关闭时间、结果显示时机、动画速度等）
  - 上次打开照片的目录路径
- 程序启动时自动恢复这些配置，拖拽分割条、调整窗口大小后实时保存。

---

## 4. 目录结构

```
程序根目录/
├── App14345.exe                 # 打包后的可执行文件（发布时）
├── app.ico                      # 应用程序图标
├── main.py                      # 源代码入口
├── config/                      # 配置目录
│   ├── settings.json            # 用户设置
│   └── stats.json               # 统计数据（抽选历史）
├── person/                      # 人员数据目录
│   ├── persons.json             # 人员列表（JSON数组）
│   └── photo/                   # 头像文件存储
├── 日志/                        # 日志文件目录
│   └── 年月日_时分秒_4位串号.log
└── template/                    # 导入模板目录
    └── 人员导入模板.xlsx
```

---

## 5. 安装与运行

### 5.1 开发环境准备
1. 安装 Python 3.13（或更高版本）。
2. 安装依赖包：
   ```bash
   pip install PyQt5 Pillow openpyxl matplotlib
   ```
3. 确保有可用的中文字体（系统默认即可，matplotlib 已配置为 Microsoft YaHei / SimHei）。

### 5.2 运行源代码
```bash
python main.py
```

### 5.3 打包为独立可执行文件
使用 PyInstaller（推荐命令）：
```bash
pyinstaller --onefile --windowed --icon=app.ico --name=App14345 --hidden-import=matplotlib.backends.backend_qt5agg --hidden-import=openpyxl --add-data "app.ico;." main.py
```
生成的可执行文件位于 `dist/App14345.exe`，可复制到任意位置运行，无需 Python 环境。

注：打包前需确保 `app.ico` 存在于当前目录。

---

## 6. 核心算法详解

### 6.1 随机抽取算法（伪代码）

```python
def pick(participant_ids, k):
    N = len(participant_ids)
    # 步骤1：映射随机
    shuffled = secrets.shuffle(participant_ids)  # 随机打乱
    seq_to_id = {i+1: id for i, id in enumerate(shuffled)}
    # 步骤2：微秒区间划分
    total = 999999  # 有效微秒范围 1..999999
    block = total // N
    remainder = total % N
    max_valid = total - remainder   # 舍弃尾部余数
    # 步骤3：抽取 k 次（不放回）
    chosen = []
    for _ in range(k):
        while True:
            ms = get_microsecond()   # 当前微秒
            if ms <= max_valid:
                break
        seq = (ms - 1) // block + 1   # 计算序号
        chosen_id = seq_to_id[seq]
        chosen.append(chosen_id)
        # 不放回：从参与者中移除已中人，重新映射（N-1, 新区间）
        participant_ids.remove(chosen_id)
        N -= 1
        if N == 0: break
        # 重新计算 block, max_valid, 重新打乱剩余人员
        shuffled = secrets.shuffle(participant_ids)
        seq_to_id = {i+1: id for i, id in enumerate(shuffled)}
        block = total // N
        remainder = total % N
        max_valid = total - remainder
    return chosen
```

### 6.2 日志和动画
- 所有抽取步骤会通过 `ui_callback` 实时显示在右侧过程面板，并同时写入 `.log` 文件。
- 动画使用 `QTimer` 逐步更新卡片高亮，速度分为前、中、后三段可调（默认 40/80/120 毫秒/步），步数分配：前段 = 总步数 - 40，中段 20 步，后段 20 步；总步数 = 5 × 参与人数 + 10。

---

## 7. 数据存储格式

### 7.1 persons.json
```json
[
  {
    "id": "1",
    "nickname": "✨星_1",
    "realname": "张伟",
    "photo": "1_20260504123456.png",
    "participate": true
  }
]
```

### 7.2 stats.json
```json
[
  {
    "time": "2026-05-04T12:34:56.789123",
    "chosen_ids": ["3", "7"]
  }
]
```

### 7.3 settings.json
完整默认配置示例如下（实际会根据用户修改变化）：
```json
{
  "last_photo_dir": "C:\\Users\\xxx\\Pictures",
  "card_show_nickname": true,
  "card_show_realname": false,
  "last_show_nickname": true,
  "last_show_realname": false,
  "draw_show_nickname": true,
  "draw_show_realname": false,
  "log_panel_visible": true,
  "draw_k": 1,
  "draw_repeat": 1,
  "show_title_copyright": true,
  "auto_close_single_enabled": false,
  "auto_close_single_seconds": 2.0,
  "auto_close_multi_enabled": true,
  "auto_close_multi_seconds": 0.5,
  "anim_speed_front": 40,
  "anim_speed_mid": 80,
  "anim_speed_back": 120,
  "result_display_immediate": false,
  "show_history_list": true,
  "show_process_log": true,
  "main_splitter_sizes": [800, 400],
  "right_splitter_sizes": [300, 400],
  "window_geometry": "AAAA/..."
}
```

---

## 8. 常见问题与解决方案

| 问题 | 可能原因 | 解决方案 |
|------|----------|----------|
| 批量导入时提示缺少 openpyxl | 未安装依赖 | `pip install openpyxl` |
| 柱状图显示乱码 / 方格 | matplotlib 字体未配置 | 代码中已配置微软雅黑/黑体，确保系统有该字体 |
| 裁剪照片后图片变形 | 缩放比例计算有误 | 已使用 `Qt.KeepAspectRatio` 保持比例，确保原图非空 |
| 单实例无效 | QLocalServer 名称冲突或权限 | 重启系统或更改 server name（代码中为 `App14345_single_instance`） |
| 抽取动画卡顿 | 动画序列过长或定时器过密 | 已优化序列长度（5 轮 + 10 次目标），高亮切换采用局部刷新 |
| 托盘图标不显示 | 未找到 app.ico 或系统限制 | 确保 icon 文件存在，或使用内置标准图标 |
| 多次抽取弹窗不自动关闭 | 配置中未启用或定时器错误 | 检查 `auto_close_multi_enabled` 是否为 true，代码已使用 `QTimer.singleShot` 确保关闭 |
| 弹窗未居中 | 窗口移动后未更新位置 | 已重写 `showEvent` 和 `resizeEvent`，每次显示时重新计算主窗口中心 |

---

## 9. 扩展与维护指南

### 9.1 添加新的显示配置
1. 在 `SettingsDialog` 中添加新的控件（如 `QCheckBox`、`QSpinBox`）。
2. 在 `load_settings()` 和 `accept()` 中处理对应配置项的读和写。
3. 在主窗口中应用配置（如刷新卡片显示样式）。

### 9.2 修改抽取算法
- 替换 `RandomPicker.pick()` 中的随机映射或微秒分组逻辑。
- 注意保留 `ui_callback` 和日志记录功能，保持对用户透明。

### 9.3 增加新的导入格式（如 CSV）
- 在 `PersonManager` 中添加 `import_from_csv()` 方法。
- 在 `PersonManageDialog` 中增加对应的按钮和对话框。
- 遵循现有异常处理和反馈机制。

### 9.4 升级 PyQt6
- 将 import 中的 `PyQt5` 改为 `PyQt6`，并修改对应模块（如 `QDesktopServices` 仍在 `QtGui` 中）。注意信号槽语法的变化（`pyqtSignal` 替换为 `QtCore.pyqtSignal`）。

---

## 10. 版本历史

| 版本 | 日期 | 修改内容 |
|------|------|----------|
| 1.3.22 | 2026-05-04 | 最终稳定版：弹窗非模态绝对居中、自动关闭修复、动画速度可调、Excel 内嵌图片修复、表格居中、柱状图中文支持、重置配置功能、分割线状态持久化等 |
| 1.3.21 | 2026-05-03 | 完善日志延迟显示、多次抽取动画后更新结果 |
| 1.3.20 | 2026-05-02 | 初始完整版 |

---

## 11. 开发者联系

- **作者**：XAF
- **邮箱**：暂无
- **版权**：Copyright © XAF 2026.5

**文档结束**
