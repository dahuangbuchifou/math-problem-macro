# 第八章：工业级自动执行与全局状态机控制矩阵
**WQRS V4.1 — Ch08 Execution Engine V4.1 全维闭环版**

> **修订说明（V4.0 → V4.1）**：
> 1. 明确 GTSM 对 Ch07 输出的紧急接管优先级：极端盘面触发时注入 `EMERGENCY_OVERRIDE` 信号
> 2. 明确跌停抢跑与盘中止损的时序关系
> 3. Flash Crash 触发阈值从 -3% 修正至 -2.5%（与总纲保持一致）
> 4. 完善审计日志结构

---

## 8.1 本章本质

本章是 WQRS 的**物理世界硬核接口与分布式状态控制层**，誓死捍卫三件事：

- **执行原子性**：信号正确 ≠ 成交正确。确保任意一笔总单从生成到终结，中途绝不丢失状态，绝不重复下注。
- **微观结构对抗**：通过算法拆单，对抗 A 股盘口滑点、撮合延迟、机构算法钓鱼与流动性黑洞。
- **黑天鹅生存权**：强行接管涨停锁死、跌停闷杀、闪崩穿透等极端盘面的拦截与断臂求生。

---

## 8.2 模块控制流拓扑总线

```
[第七章 仓位 JSON]
        │
        ▼
┌────────────────────────────────┐
│    8.3 全局交易状态机 (GTSM)    │ ◀── EMERGENCY_OVERRIDE 通道（来自极端盘面拦截）
│    (全局 Master_State 守护者)   │
└───────────────┬────────────────┘
                │
                ▼
    [8.4 订单清洗与规范层]
                │
                ▼
    [8.5 智能拆单与路由引擎]
                │
                ▼
    [8.6 极端盘面风控拦截网]  ──→ EMERGENCY_OVERRIDE 注入 GTSM
                │
                ▼
    [8.7 物理接口与安全重试]
                │
                ▼
        [券商柜台 / 交易所]
```

---

## 8.3 核心大脑：全局交易状态机（GTSM）

每一笔进入第八章的总订单，必须分配全域唯一 `Master_Order_ID`，并在系统内存中注册对应的确定性有限状态机（DFA）。

### 8.3.1 状态全生命周期定义

| 状态 | 含义 |
|---|---|
| `PENDING` | 第七章信号刚送达，等待进入规范清洗 |
| `VALIDATED` | 通过第一层规范检查，冻结相应预算 |
| `SPLITTING` | 正在被拆单算法切割成子订单流 |
| `SUBMITTED` | 首笔或当前批次子订单已发送至券商柜台 |
| `PARTIAL_FILLED` | 子订单部分成交，剩余挂单驻留盘口 |
| `RETRYING` | 子订单超时未成，触发撤单，重试引擎重新计算补报 |
| `FILLED` | 所有子订单全部 100% 成交，状态归档 |
| `DEAD_ORDER` | 触及硬拒单、最大撤单红线或人工熔断，释放冻结资金，写入死单骨架 |

### 8.3.2 合法状态转移路径

- **路径一（完美顺延）**：`PENDING → VALIDATED → SPLITTING → SUBMITTED → FILLED`
- **路径二（盘口震荡容错）**：`SUBMITTED → PARTIAL_FILLED → RETRYING → SUBMITTED → FILLED`
- **路径三（熔断断臂）**：`SUBMITTED / PARTIAL_FILLED → RETRYING(超限) → DEAD_ORDER`
- **路径四（紧急接管，V4.1 新增）**：`任意状态 → EMERGENCY_OVERRIDE → DEAD_ORDER（买入）/ 立即强制卖出`

**任何非法状态跳跃（如 PENDING 直接跳到 FILLED）将触发操作系统级死机保护。**

### 8.3.3 EMERGENCY_OVERRIDE 接管机制（V4.1 新增）

**触发来源**：8.6 极端盘面风控拦截网检测到系统性闪崩或多个持仓同步跌停。

**接管逻辑**：

```
1. GTSM 接收 EMERGENCY_OVERRIDE 信号
2. 立即冻结所有 BUY / 加仓操作（不论当前状态）
3. 中止 Ch07 的所有挂起输出（不再接收新的仓位JSON）
4. 将所有算力与 API 通道让渡给止损与清仓状态机
5. 写入审计日志：OVERRIDE_TRIGGERED，附带触发原因与时间戳
```

**与 Ch07 止损的时序关系**：
- **正常盘面**：Ch07 计算止损价 → Ch06 触发退出信号 → Ch08 执行卖出
- **极端盘面（跌停/闪崩）**：Ch08 的 GTSM 直接注入 EMERGENCY_OVERRIDE，**绕过 Ch07 的正常止损计算链路**，以最高优先级执行卖出，事后将该事件写入审计日志，反哺修正 Ch07 的滑点模型

---

## 8.4 第一层：订单清洗与规范层

当 GTSM 接收到 `PENDING` 信号后，立刻启动物理环境对齐清洗。

### 标准输入契约 JSON

```json
{
  "master_order_id": "WQRS-20260603-001A",
  "asset": {"code": "002371", "market": "SZ"},
  "order": {
    "side": "BUY",
    "total_shares": 2400,
    "target_price": 312.50
  }
}
```

### 拦截过滤规则

| 检查项 | 条件 | 处理 |
|---|---|---|
| 标的资格 | ST / *ST / 退 / 停牌 | 直接 `DEAD_ORDER` |
| 物理整倍数 | `total_shares % 100 ≠ 0` 或 `≤ 0` | 直接 `DEAD_ORDER` |
| 资金合规 | 可用资金不足 | 直接 `DEAD_ORDER`（硬失败，禁止重试）|

---

## 8.5 第二层：智能拆单与分时路由引擎

### A股专用冰山对齐算法

**核心原则**：严禁粗暴除法切分，所有子单必须为100股整倍数。

```
设目标拆分笔数 K（默认 K=3）
Base = floor(Total / (K × 100)) × 100
前 K-1 笔 = Base
最后一笔 = Total - Base × (K-1)  # 自动消化余数
```

示例：2400股拆3笔 → 800 + 800 + 800，完全对齐。

### 时段分发调度矩阵

| 时段 | 策略 | 说明 |
|---|---|---|
| 09:30–09:45（极速攻击期）| 一波流合并，LATEST_PRICE 全额扫货 | VCP 突破建仓与硬止损的生死时速 |
| 10:00–14:30（冰山潜行期）| 随机 3000ms–7000ms 冰山脉冲路由 | 均价咬住日内 VWAP，防止被对手算法探测 |
| 14:50–15:00（尾盘清扫期）| 强制撤单，14:56.500 集合竞价跨价单出清 | 清理所有 SUBMITTED / PARTIAL_FILLED 残单 |

---

## 8.6 第三层：极端盘面风控拦截网

### 涨停锁死诱多陷阱（Limit Up Shield）

- **触发条件**：`Ask_Price[0] ≥ Limit_Up_Price`（卖一价封死涨停）
- **拦截动作**：终止后续一切买入子单，对排队子单下达撤单，GTSM 强制转为 `DEAD_ORDER`
- **逻辑**：排队大概率遭遇尾盘炸板闷杀，低成交概率 + 高被收割风险，WQRS 拒绝参与

### 跌停生死一毫秒（Limit Down Survival）

- **触发条件**：持仓股 `Bid_Price[0] ≤ Limit_Down_Price`
- **拦截动作**：全面作废冰山规则，在 **09:29:59.999** 将所有持仓汇聚成单笔跌停价卖出单倾泻而出
- **时序说明（V4.1 明确）**：此为 EMERGENCY_OVERRIDE 触发场景，绕过 Ch07 止损计算链路，GTSM 直接接管执行，事后写入审计日志

### 市场全域流动性枯竭（Flash Crash Melt-down）

- **触发条件**：沪深300/创业板指 30分钟内暴跌 **> 2.5%**，或持仓列表中多于 3 只个股同步触发跌停
- **拦截动作**：
  1. GTSM 注入 `EMERGENCY_OVERRIDE` 信号（见 8.3.3）
  2. 全局冻结所有 BUY 与加仓功能
  3. 全部算力与 API 通道让渡给清仓止损状态机

---

## 8.7 第四层：物理接口、死锁解锁与安全重试机制

### 软失败与进攻性滑点补偿（Soft Failure Retry）

- **定义**：子单发出后价格上行过快，超时未成（5秒），状态切入 `RETRYING`
- **解法**：
  1. 发出撤单指令，挂起等待券商返回 `ORDER_CANCELED` 回调
  2. 接收成功撤单后，新子单价格 = `Current_Latest_Price + 0.02`
  3. 重新下发，状态回 `SUBMITTED`
- **硬性计数器红线**：单一 Master 订单累计重试 **MAX_RETRIES = 3 次**，超限强制 `DEAD_ORDER`

### 硬失败原子化拒绝（Hard Failure Termination）

收到 `ERROR_INSUFFICIENT_FUNDS`、`ERROR_INVALID_STOCK_CODE` 或权限受限：

**绝对禁止重试！** 直接 `DEAD_ORDER`，触发系统级微信/钉钉风控警报。

---

## 8.8 QMT/XtQuant 实盘状态机核心骨架

```python
from xtquant import xttrader
import math, random, time

class WQRSTradeStateMachine:
    def __init__(self):
        self.master_registry = {}
        self.max_retries = 3
        self.emergency_override_active = False  # V4.1 新增

    def inject_emergency_override(self, reason: str):
        """V4.1 新增：极端盘面接管入口"""
        self.emergency_override_active = True
        # 冻结所有 BUY 操作
        for m_id, meta in self.master_registry.items():
            if meta["side"] == "BUY" and meta["master_state"] not in ("FILLED", "DEAD_ORDER"):
                meta["master_state"] = "DEAD_ORDER"
                print(f"[EMERGENCY] {m_id} 被紧急接管，原因: {reason}")
        # 写入审计日志
        self._write_audit_log("OVERRIDE_TRIGGERED", {"reason": reason})

    def register_new_order(self, ch7_json):
        """接收第七章挂起信号，初始化 GTSM"""
        # 紧急接管激活时，拒绝所有新 BUY 订单
        if self.emergency_override_active and ch7_json.get("side") == "BUY":
            print("[EMERGENCY] 紧急接管激活，拒绝新买入订单")
            return
        m_id = ch7_json["master_order_id"]
        self.master_registry[m_id] = {
            "code": ch7_json["asset"]["code"],
            "side": ch7_json["order"]["side"],
            "total_shares": ch7_json["order"]["total_shares"],
            "target_price": ch7_json["order"]["target_price"],
            "master_state": "PENDING",
            "child_orders": {},
            "retry_counts": 0,
            "filled_shares": 0
        }
        self.process_state_engine(m_id)

    def process_state_engine(self, m_id):
        """状态机核心中央转移驱动逻辑"""
        order_meta = self.master_registry[m_id]
        state = order_meta["master_state"]

        if state == "PENDING":
            if order_meta["total_shares"] % 100 != 0 or order_meta["total_shares"] <= 0:
                order_meta["master_state"] = "DEAD_ORDER"
                print(f"[CRITICAL] {m_id} 股数非百股倍数，GTSM 击杀订单。")
                return
            order_meta["master_state"] = "VALIDATED"
            self.process_state_engine(m_id)

        elif state == "VALIDATED":
            order_meta["master_state"] = "SPLITTING"
            total = order_meta["total_shares"]
            # 工业级 A 股100股对齐拆单算子
            K = 3
            base_child = math.floor(total / (K * 100)) * 100
            splits = [base_child] * (K - 1)
            splits.append(total - base_child * (K - 1))  # 最后一笔消化余数
            order_meta["splits"] = splits
            order_meta["master_state"] = "SUBMITTED"
            self.execute_next_child(m_id)

    def execute_next_child(self, m_id):
        """物理执行路由层"""
        order_meta = self.master_registry[m_id]
        if not order_meta["splits"]:
            order_meta["master_state"] = "FILLED"
            print(f"[SUCCESS] {m_id} 全额成交，生命周期完美闭合。")
            return
        next_volume = order_meta["splits"].pop(0)
        # 第三层极端盘面拦截（此处为伪代码钩子）
        # if current_tick.ask >= limit_up: self.inject_emergency_override("涨停封板"); return
        print(f"[PHYSICAL] 发射子订单: {next_volume} 股, 标的: {order_meta['code']}")
        xt_id = "XT-" + str(random.randint(100000, 999999))
        order_meta["child_orders"][xt_id] = {"volume": next_volume, "status": "SENT"}

    def on_broker_callback(self, xt_id, broker_status, m_id):
        """券商异步回调，喂入 GTSM 中央大脑"""
        order_meta = self.master_registry[m_id]
        child = order_meta["child_orders"].get(xt_id)
        if not child: return

        if broker_status == "FILLED":
            child["status"] = "FILLED"
            order_meta["filled_shares"] += child["volume"]
            time.sleep(random.uniform(0.3, 0.7))
            self.execute_next_child(m_id)

        elif broker_status == "TIMEOUT_OR_FAILED":
            order_meta["master_state"] = "RETRYING"
            if order_meta["retry_counts"] >= self.max_retries:
                order_meta["master_state"] = "DEAD_ORDER"
                print(f"[风控熔断] {m_id} 重试超限，转为死单。")
                return
            order_meta["retry_counts"] += 1
            order_meta["splits"].insert(0, child["volume"])
            order_meta["master_state"] = "SUBMITTED"
            self.execute_next_child(m_id)

        elif broker_status in ("ERROR_INSUFFICIENT_FUNDS", "ERROR_INVALID_STOCK_CODE"):
            # 硬失败，绝对禁止重试
            order_meta["master_state"] = "DEAD_ORDER"
            print(f"[HARD FAILURE] {m_id} 硬失败，禁止重试，直接死单。")
            self._send_alert(m_id, broker_status)

    def _write_audit_log(self, event_type: str, payload: dict):
        """审计日志写入（生产环境接入持久化存储）"""
        import datetime
        log = {"event": event_type, "timestamp": datetime.datetime.utcnow().isoformat(), **payload}
        print(f"[AUDIT] {log}")

    def _send_alert(self, m_id: str, reason: str):
        """触发微信/钉钉风控警报（生产环境接入消息通道）"""
        print(f"[ALERT] 订单 {m_id} 触发警报: {reason}")
```

---

## 8.9 数字化资产执行总账与日志

实盘中每一次状态转移、撤单重试及滑点冲击成本，必须实时沉淀为**不可篡改的审计日志**，每日 15:00 收盘后反哺修正 Ch06 的择时模型与 Ch07 的滑点模型。

```json
{
  "telemetry_version": "WQRS-v4.1-Production",
  "timestamp": "2026-06-03T14:57:02.105Z",
  "master_order_id": "WQRS-20260603-001A",
  "audit_trail": {
    "gtsm_final_state": "FILLED",
    "total_requested_shares": 2400,
    "total_executed_shares": 2400,
    "target_benchmark_price": 312.50,
    "real_blended_execution_price": 312.56,
    "net_slippage_cost_ratio": "0.019%",
    "emergency_override_triggered": false
  },
  "hardware_diagnostics": {
    "total_network_retries": 1,
    "average_broker_latency_ms": 7.2,
    "firewall_limit_up_interceptions": 0,
    "firewall_limit_down_interceptions": 0,
    "flash_crash_overrides": 0
  }
}
```

**V4.1 新增审计字段：**
- `emergency_override_triggered`：本单是否触发过紧急接管
- `flash_crash_overrides`：本交易日触发 Flash Crash 熔断次数

---

## 8.10 本章终极总结

- **Ch07 仓位下注矩阵**：解决系统"该买多少"的算力分配
- **Ch08 全局交易状态机**：解决系统在物理世界中"如何绝对确定、合规、毫无歧义地拿到筹码并活下来"

通过引入 GTSM 与 V4.1 的 EMERGENCY_OVERRIDE 接管机制，WQRS 实盘执行层实现了完整的**六层闭环**：

```
宏观许可 → 行业选择 → 标的筛选 → 信号触发 → 仓位计算 → 物理执行
```

不论盘口如何闪崩、券商通道如何延迟，在这套刚性状态机的统治下，每一颗子弹都将以数学般的精确度，冷酷地砸向交易所的撮合核心。
