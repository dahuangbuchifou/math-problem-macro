# 数学题生成宏命令

## 项目简介

本项目提供一套数学题生成的宏命令，用于快速、批量地生成各类数学题目。

## 功能特性

- 📐 支持多种数学题型（加减乘除、分数、方程等）
- ⚡ 批量生成，提高效率
- 🎯 可配置难度和题目数量
- 📝 输出格式标准化，便于后续处理

## 使用方法

### 基本用法

```bash
# 生成指定数量的数学题
./math-problem-macro --count 10

# 生成指定类型的题目
./math-problem-macro --type addition --count 20

# 生成带答案的题目
./math-problem-macro --with-answer
```

### 配置选项

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--count` | 生成题目数量 | 10 |
| `--type` | 题目类型（addition/subtraction/multiplication/division） | addition |
| `--difficulty` | 难度等级（1-5） | 3 |
| `--with-answer` | 是否包含答案 | false |
| `--output` | 输出文件格式（txt/json/latex） | txt |

## 项目结构

```
math-problem-macro/
├── README.md          # 项目说明
├── src/               # 源代码
├── config/            # 配置文件
└── tests/             # 测试文件
```

## 开发计划

- [ ] 支持更多题型（分数、百分比、几何）
- [ ] 添加 LaTeX 输出支持
- [ ] 支持自定义题目模板
- [ ] 添加单元测试和集成测试

## 许可证

MIT License

---

_婉儿派单 · 大黄要往里面放需求_ 🎋
