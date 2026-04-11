# 求职自动化工作流（Seek + LinkedIn）

这个项目可以做到：

1. 读取职位列表（`jobs/jobs.csv`）
2. 基于两份简历模板按 JD 生成定制简历
3. 基于同一 JD 生成 Cover Letter
4. 评估 ATS 对齐度（关键词覆盖、章节完整度）并输出报表（CSV + Excel）
4. 连接浏览器会话（你手动授权登录 Seek / LinkedIn）
5. 半自动打开岗位页面并更新投递状态

## 1) 安装

```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium
Copy-Item .env.example .env
```

如果你希望使用 AI 改写简历，请在 `.env` 填入：

```env
OPENAI_API_KEY=你的key
OPENAI_MODEL=gpt-5-mini
```

## 2) 准备你的简历模板

支持 `docx` 和 `markdown` 模板。你当前主模板可直接用：

- `resumes/Douglas_Mo_Resume_Formatted.docx`

也可以保留备用模板：

- `resumes/resume_template_1.md`
- `resumes/resume_template_2.md`

## 3) 准备岗位列表

编辑 `jobs/jobs.csv`，字段如下：

- `job_id`: 自定义唯一 ID（可空，空则自动生成）
- `platform`: `seek` 或 `linkedin`
- `company`
- `title`
- `location`
- `job_url`: 岗位链接（推荐填）
- `jd_text`: 直接粘贴 JD 文本（若有则优先使用）
- `notes`

## 4) 首次授权登录（浏览器会话持久化）

Seek：

```powershell
py scripts/job_workflow.py auth --platform seek
```

LinkedIn：

```powershell
py scripts/job_workflow.py auth --platform linkedin
```

运行后会打开浏览器，你手动登录并完成验证，回到终端按 Enter 即保存会话。

## 5) 生成定制简历 + 评估结果

使用 AI（推荐）：

```powershell
py scripts/job_workflow.py run --use-openai
```

不使用 AI（仅规则增强）：

```powershell
py scripts/job_workflow.py run
```

输出位置：

- 定制简历：`outputs/custom_resumes/`
- Cover Letter：`outputs/cover_letters/`
- 报表：`outputs/reports/job_results_*.csv`
- Excel 表：`outputs/reports/job_results_*.xlsx`
- 最新追踪表：`outputs/reports/application_tracker.csv`

报表关键字段：

- `match_score`: 规则关键词匹配分
- `ats_score`: JD高频关键词命中分（ATS导向）
- `section_score`: 简历核心章节完整度
- `resume_file`
- `cover_letter_file`

默认输出为 Word 文档（`.docx`），可直接用于投递上传。

## 6) 投递助手（半自动）

该步骤会按平台逐个打开岗位链接，你在浏览器中手动点击投递，终端记录状态。

```powershell
py scripts/job_workflow.py apply --platform seek --limit 20
py scripts/job_workflow.py apply --platform linkedin --limit 20
```

自动尝试点击投递按钮（浏览器对话式）：

```powershell
py scripts/job_workflow.py apply --platform seek --limit 20 --auto-click
py scripts/job_workflow.py apply --platform linkedin --limit 20 --auto-click
```

状态支持：

- `submitted`
- `skipped`
- `failed`

## 7) 你的完整循环

1. 新增岗位到 `jobs/jobs.csv`
2. 执行 `run` 生成定制简历
3. 执行 `apply` 完成投递并更新状态
4. 查看 `application_tracker.csv` 或 Excel 做复盘

## 说明

- Seek/LinkedIn 页面结构经常变化，完全自动点击投递容易失效，因此项目采用“会话授权 + 半自动投递”的稳定方案。
- 如果你希望下一步升级为“自动识别页面字段并自动填写”，可以在此脚本基础上继续加 Playwright 选择器策略。
