# 第五章：基本面与估值评分引擎
**WQRS V4.1 — Ch05 Fundamental & Valuation Engine**

> **修订说明**：
> 1. 新增 Ch04 行业退出联动规则：订阅 `dropped_sectors` 字段，触发白名单自动降级为 `watchlist_suspended`
> 2. 白名单输出新增 `whitelist_status` 字段，供 Ch06 信号引擎判断是否允许新开仓信号

---

## 5.1 设计目标

### 5.1.1 本章定位

- Ch02 识别市场所处宏观状态
- Ch03 确定风险预算与资产配置比例
- Ch04 确定最具配置价值的行业方向

**本章任务**：在优势行业中筛选真正值得持有的公司。

**本章只回答一个问题：买什么？**

第六章信号引擎负责回答：什么时候买？什么时候卖？**两章职责边界严格分离，Ch05 不做择时，Ch06 不做选股。**

### 5.1.2 核心原则

WQRS 寻找的是：质量优秀、估值合理、风险可控、具备持续经营能力的公司。

优先级：**风险控制 > 质量评估 > 估值评估 > 排序**

---

## 5.2 处理流程

```
Sector Rotation Engine (Ch04)
          ↓
    Candidate Universe
          ↓
    [订阅 dropped_sectors → 触发白名单降级]
          ↓
      Hard Filter
          ↓
  Accounting Red Flag
          ↓
    Quality Score
          ↓
   Valuation Score
          ↓
 Fundamental Score
          ↓
    Final Rank
          ↓
  Candidate Pool (含 whitelist_status)
          ↓
  Signal Engine (Ch06)
```

---

## 5.3 Hard Filter（财务排雷层）

Hard Filter 目的是**先排除明显危险公司**，避免后续评分失真。

**排除条件：**

| 条件 | 标准 |
|---|---|
| 连续亏损 | 近3年累计净利润 ≤ 0 |
| ST 风险 | ST / *ST / 退市整理 |
| 财务杠杆过高 | 资产负债率 > 80%（金融行业除外）|
| 流动性不足 | 近60日平均成交额 < 5000万元 |

命中任意条件 → 直接淘汰（`Reject`）。

---

## 5.4 Accounting Red Flag Engine

### 设计目标

识别潜在财务异常。注意：**异常 ≠ 造假**，因此不直接淘汰，而是进入 `Fraud Review Queue` 等待人工复核。

### 触发规则

| 规则 | 触发条件 |
|---|---|
| 存贷双高 | 货币资金 > 有息负债 且 货币资金/总资产 > 30% |
| 应收账款异常 | 应收账款增长率 > 营收增长率 × 2，持续两个报告期 |
| 现金流背离 | 净利润增长 > 0 且 经营现金流增长 < 0，持续两个报告期 |
| 商誉风险 | 商誉 > 净资产 30% |

命中任意规则 → 进入 `Fraud Review Queue` 并附加 `Risk Flag`，**复核前禁止进入白名单**。

---

## 5.5 Quality Score（企业质量评分）

```
Quality Score = 40% × Profitability
              + 30% × Financial Health
              + 30% × Business Durability
```

总分：0～100

### Profitability（盈利能力）
- ROE Percentile：40%
- ROIC Percentile：40%
- Net Margin Percentile：20%

### Financial Health（财务健康度）
- 经营现金流：40%
- 资产负债率：30%
- 利息覆盖倍数：30%

### Business Durability（经营持续性）
防止周期高点企业获得虚高评分，观察近5年稳定性：
- ROE 波动率：40%
- 毛利率波动率：30%
- 营业收入行业排名变化：30%

---

## 5.6 Valuation Score（估值评分）

估值用于**避免高估**，而非寻找最低估。

```
Valuation Score = 50% × Relative Valuation
               + 30% × Historical Valuation
               + 20% × Cashflow Yield
```

- **Relative Valuation**：行业内比较 PE/PB/EV·EBITDA 分位数
- **Historical Valuation**：公司历史近5年 PE/PB 分位数
- **Cashflow Yield**：FCF Yield 现金流收益率

---

## 5.7 Fundamental Score

```
Fundamental Score = 70% × Quality Score + 30% × Valuation Score
```

企业质量优先于估值。

---

## 5.8 Container Ranking Logic

| 容器 | 排序公式 | 说明 |
|---|---|---|
| A容器（权益增长池）| `30% × Sector Score + 70% × Fundamental Score` | 顺势成长，结合行业动量 |
| B容器（高股息压舱石）| `100% × Fundamental Score` | 不受行业轮动约束 |
| C容器（商品与通胀对冲）| ETF评分体系 | 不使用企业质量评分 |

---

## 5.9 Candidate Tier System

| Tier | Final Rank | 说明 |
|---|---|---|
| Tier 1 | ≥ 90 | 核心持仓候选 |
| Tier 2 | 80–89 | 观察名单 |
| Tier 3 | 70–79 | 备选名单 |
| 剔除 | < 70 | 不进入候选池 |

---

## 5.10 白名单联动降级机制（V4.1 新增）

### 触发条件

Ch05 订阅 Ch04 输出的 `dropped_sectors` 字段。当某行业出现在 `dropped_sectors` 中时，该行业内**所有白名单股票**状态自动更新为 `watchlist_suspended`。

### 状态定义

| whitelist_status | 含义 | Ch06 行为 |
|---|---|---|
| `active` | 正常白名单 | 允许新开仓信号 |
| `watchlist_suspended` | 行业退出触发降级 | **拒绝新开仓信号**，存量仓位正常管理 |
| `fraud_review` | 财务红旗待复核 | 禁止进入白名单 |
| `expired` | 超过 TTL 365天 | 自动重评 |

### 关键设计原则

- **不清除存量仓位**：`watchlist_suspended` 仅阻止新信号，存量持仓由 Ch06 的退出矩阵（跌破20MA、基本面退出等）正常管理
- **降级可逆**：若该行业重新进入 Ch04 Top 3，对应股票状态自动恢复为 `active`

---

## 5.11 White List Lifecycle

白名单有效期管理，避免白名单永久有效：

| 触发事件 | 动作 |
|---|---|
| 季报更新 | Quarterly Review |
| 年报更新 | Annual Review |
| 监管事件（立案/财务重述/重大处罚）| Regulatory Review |
| 生命周期到期（TTL = 365天）| 自动重评 |
| **Ch04 dropped_sectors 触发（新增）** | **状态降级为 watchlist_suspended** |

---

## 5.12 输出接口

```json
{
  "candidate_pool": [
    {
      "ticker": "600519",
      "company_name": "贵州茅台",
      "sector": "食品饮料",
      "quality_score": 92,
      "valuation_score": 68,
      "fundamental_score": 84,
      "final_rank": 87,
      "risk_flags": [],
      "candidate_tier": 2,
      "whitelist_status": "active"
    }
  ],
  "suspended_count": 3,
  "fraud_review_count": 1
}
```

**新增字段（V4.1）：**
- `whitelist_status`：`active` / `watchlist_suspended` / `fraud_review` / `expired`
- `suspended_count`：当前因行业退出而降级的股票数量
- `fraud_review_count`：待人工复核的股票数量

---

## 5.13 与第六章接口定义

| 章节 | 职责 |
|---|---|
| Ch05 | What To Buy（选股） |
| Ch06 | When To Buy / When To Sell（择时）|

**严格边界**：Ch06 的 Signal Engine 只能从 Ch05 的 Candidate Pool 中选取标的，**禁止绕过 Fundamental Engine 直接选股**。Ch06 在生成信号前必须检查 `whitelist_status`，`watchlist_suspended` 状态的股票不得触发新开仓信号。

---

## 5.14 本章总结

Fundamental & Valuation Engine 是 WQRS 的第二道风险防线，通过 Hard Filter → Accounting Red Flag → Quality Score → Valuation Score → Tier Ranking 逐层过滤噪音与风险，最终形成**高质量、可解释、可复核、可量化**的 Candidate Pool。

至此，WQRS 已完成：宏观判断 → 风险预算 → 行业选择 → 公司选择，四层决策框架的完整闭环。
