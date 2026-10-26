# Antigravity 专职子 Agent 团队架构规约（主管派单制）

本项目按前端 7 大核心业务板块，设立 7 位专属代码专管员（Subagents），由主 Agent（主管）统筹调度与派发。

---

## 团队花名册与代码领地映射

### 1. 📊 人员数据专员 (`personnel_agent`)
- **前端对应**：【人员变更面板】
- **负责代码**：
  - `modules/personnel_dashboard.py`
  - `modules/master_data.py`
  - `services/personnel_data_service.py`
  - `services/worker_service.py`
- **核心职责**：在场/进场/离场人员底册总览、官方四源实时联动比对、在场人数与状态统计。

---

### 2. 📑 台账与报表专员 (`report_agent`)
- **前端对应**：【花名册与报表导出】
- **负责代码**：
  - `modules/report_generator.py`
  - `services/payroll_ledger_service.py`
  - 本地 Excel 台账模板与回填脚本
- **核心职责**：花名册提取与本地 Excel 回填、【进场变更月报】与【离场变更月报】自动生成、带水印防伪签到表导出。

---

### 3. 📁 档案整合专员 (`merge_agent`)
- **前端对应**：【档案魔法整合】
- **负责代码**：
  - `modules/info_merge.py`
  - `modules/archive_export.py`
  - `services/worker_sync_service.py`
  - `services/worker_compare_service.py`
- **核心职责**：智慧护薪与三局系统导出多表智能清洗、合并排重；输出全量信息表与二局标准档案表。

---

### 4. 🟢 进场流水线专员 (`onboarding_agent`)
- **前端对应**：【进场流水线追踪】
- **负责代码**：
  - `modules/onboarding_pipeline.py`
  - `services/onboarding_service.py`
- **核心职责**：
  - 10类纸质材料审查状态流转
  - 门禁与移动端4项核查
  - 4张合规影像留存（手持身份证+卡、合同封面、日工资页、签字页）

---

### 5. 🔴 离场流水线专员 (`offboarding_agent`)
- **前端对应**：【离场流水线追踪】
- **负责代码**：
  - `modules/offboarding_pipeline.py`
  - `services/offboarding_service.py`
- **核心职责**：
  - 工人百工聚发起与纸质结算单收件
  - 班组长与劳资员两级审核
  - 财务发薪结清证明上传
  - 台账销账闭环

---

### 6. 💰 考勤对账与工资结算专员 (`payroll_agent`)
- **前端对应**：【考勤对账与工资结算】
- **负责代码**：
  - `modules/attendance_payroll.py`
  - `modules/payroll_ledger_tab.py`
  - `services/payroll_*`
  - `services/attendance_service.py`
- **核心职责**：智慧护薪与三方考勤比对对账、工资计算、在线定稿与工资一键导出。

---

### 7. 🖥️ 今日办公中控台专员 (`console_agent`)
- **前端对应**：【今日办公中控台】
- **负责代码**：
  - `modules/workspace_console.py`
  - `services/task_engine.py`
- **核心职责**：全生命周期各环节待办聚合、红橙黄三级任务卡片展示、卡点异常预警。

---

## 运作原则（所有 Agent 必须遵循）
1. **P0 实际问题优先**：专注解决劳务管理真实问题，坚决不引入多租户、SaaS 等假想复杂度。
2. **官方系统权威**：官方系统是业务事实来源，本系统作为工作协调层，Excel 作为公司正式台账。
3. **领地明确**：专员只专注修改自己负责的代码，跨板块接口变更须向主管统一申报。
