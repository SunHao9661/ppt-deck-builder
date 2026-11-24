# Development Conversation

> This is an archival development record, not runtime guidance for the skill.
> It records the main user/assistant conversation. Tool logs and hidden
> reasoning are excluded. Local absolute paths are normalized to `<WORKSPACE>`,
> `<OUTPUTS>`, or `<CODEX_HOME>` before publication.

## Turn 1

### User

> 去github上找制作ppt热度比较高的几个skill 学习并制作一个关于ppt的skill

### Assistant

已完成第一版 `ppt-deck-builder`，并安装到：

`<CODEX_HOME>/skills/ppt-deck-builder`

我参考了截至 2026-11-24 热度较高、且路线不同的项目：

| GitHub 项目 | Stars | 吸收的方法 |
|---|---:|---|
| `anthropics/skills` 的 PPTX Skill | 199,599（仓库） | 原生 PPTX、模板编辑、文件检查、逐页视觉 QA |
| `hugohe3/ppt-master` | 79,679 | 先定沟通目标和幻灯片规划，再开始画页面 |
| `zarazhangrui/frontend-slides` | 32,824 | 用具体视觉方案做选择，HTML 演示路线 |
| `op7418/guizang-ppt-skill` | 30,329 | 锁定一套视觉系统、版式角色和演讲者模式 |
| `chuspeeism/dashi-ppt-skill` | 11,053 | 模板约束、成果验收和返工闭环 |
| `ningzimu/codex-ppt-skill` | 8,136 | 图片型 PPT 的样页确认与“不可逐元素编辑”边界 |

第一版 Skill 的核心流程：

- 先判断做原生可编辑 PPTX、HTML，还是图片型 PPT
- 用 audience → outcome → story spine → ghost deck 先锁叙事
- 内置 editorial、business、tech、academic 四套视觉系统
- 支持封面、章节页、目录、结论页、卡片、对比、时间线、指标、图表、引语、结尾等页面
- 自动生成原生文本、形状、图表和演讲者备注
- 强制进行 PPTX 结构审计与逐页 PNG 渲染检查
- 修复中文 CJK 字体在无界面 LibreOffice 中显示方框的问题

第一版验证了 16:9、4:3、深色主题、图表、备注和 `contain` 图片模式。12 页测试稿的严格审计结果为：

```text
No automated issues found.
```

初始交付文件：

- `<OUTPUTS>/ppt-deck-builder-demo.pptx`
- `<OUTPUTS>/ppt-deck-builder/SKILL.md`
- `<OUTPUTS>/ppt-deck-builder/references/research-basis.md`

## Turn 2

### User

> 在结合这些github上的cyberppt、gordensupperpptskills、nanobananapptskills、gordenpptskill、htmlpptskill、pptmaster这些关于ppt的skill继续完善

### Assistant

已完成第二轮升级，并同步到：

`<CODEX_HOME>/skills/ppt-deck-builder`

六个项目的整合方式：

| 项目 | 融入能力 |
|---|---|
| CyberPPT | SCR 论证、证据底表、caveat/so-what、咨询式高密度页面、可编辑信息层与视觉语义双门槛 |
| GordenSuperPPTSkills | 图片成稿 → 可编辑重构的完整路线，强制样页确认，区分背景、框架、视觉资产和文字层 |
| NanoBanana PPT Skills | 图片型演示、转场规划、交互播放器和自运行视频链路 |
| GordenPPTSkill | 企业模板保真、只替换文字与图表数据、不破坏母版和文本框格式 |
| HTML PPT Skill | 固定 16:9 舞台、主题/布局分离、演讲者模式、键盘导航和打印/PDF 检查 |
| PPT Master | 生成、套模板、编辑原生 PPTX 三类路线分离，避免把“改现有文件”误做成“重新生成” |

Skill 现在支持的路线：

1. 原生可编辑 PPTX
2. 企业模板保真编辑
3. 图片型 PPT
4. 原生文字 + 复杂视觉资产的混合重构
5. HTML 演示
6. 带转场或自运行播放的视频演示
7. 按已批准蓝图做 1:1 高保真还原

新增工具：

- `validate_deck_plan.py`：检查故事线、行动标题、证据、来源、密度和路线缺口
- `inspect_pptx.py`：展开模板的每页形状、文本框、段落、run、坐标和图表
- `assemble_image_pptx.py`：把逐页成品图合成图片型 PPTX，保留备注和替代文本
- `compare_renders.py`：生成参考图/成品的并排图、50% 叠图和差异热图

验证结果：

- Skill 结构校验通过
- 8 页结构化 Deck Plan 严格校验通过
- 原生 PPTX 构建与严格审计，0 issues
- 图片型 PPTX 构建与严格审计，0 issues
- 16:9、4:3、中文 CJK 字体、图表、WebP 图片通过
- 模板形状/表格结构检查通过
- 单页和多页目录级参考图对照通过

这轮只吸收工作流原理，没有复制这些仓库的模板、图片、提示词或代码，避免混入不同许可证限制。

## Turn 3

### User

> 上传到我的github上

### Execution record

- 检查 GitHub 凭据：已登录为 `SunHao9661`
- 检查同名仓库：不存在
- 创建私有仓库：`SunHao9661/ppt-deck-builder`
- 执行敏感信息扫描：通过
- 准备首次提交并推送

### User

> 设置为公开

### Assistant

已设置为公开并完成上传：

**GitHub 仓库：** https://github.com/SunHao9661/ppt-deck-builder

- 可见性：Public
- 默认分支：`main`
- 首次提交：`0fd4430`
- 已添加 topics：`codex-skill`、`agent-skill`、`ppt`、`pptx`、`presentation`、`slides`
- 已验证匿名访问仓库和 `SKILL.md` 均返回 HTTP 200
- 上传前已执行敏感信息扫描，未发现 Token、私钥或密钥文件

## Turn 4

### User

> 将此对话放在制作skill文件下

### Resolution

This file was added under:

```text
ppt-deck-builder/
└── references/
    └── development-conversation.md
```

The same file is synchronized to the installed skill and the public GitHub
repository.

## Final Artifacts

- Skill entrypoint: `SKILL.md`
- Route matrix: `references/production-routes.md`
- Structured plan format: `references/deck-plan.md`
- HTML/motion guide: `references/html-motion.md`
- QA guide: `references/qa-and-delivery.md`
- Research basis: `references/research-basis.md`
- Public repository: https://github.com/SunHao9661/ppt-deck-builder
