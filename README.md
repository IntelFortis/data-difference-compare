# Data Diff Compare Studio

一个可直接部署到 **GitHub Pages** 的静态网页工具，用于对比结构相同（字段相同或高度相似）的两份或多份表单数据差异。

## 功能

- 支持上传 `.csv` / `.xlsx` / `.xls`
- Excel 文件自动读取全部 sheet（每个 sheet 作为一个可选数据源）
- 选择参与对比的数据源（至少 2 个）
- 为每个数据源单独配置：
  - 维度字段（X 轴）
  - 聚合字段（Y 值）
  - 聚合方式（sum / avg / count / min / max）
  - 可选筛选条件（`字段=值`）
- 图表由 Plotly 生成，可选：
  - 聚合趋势折线图
  - 聚合值柱状图
  - 与基准 sheet 的差值图
  - 与基准 sheet 的百分比差异图
  - 聚合结果对比表

## 本地运行

直接双击 `index.html` 即可（或使用任意静态服务器）。

## 部署到 GitHub Pages

1. 新建 GitHub 仓库并把当前目录代码推送上去。
2. 在仓库设置中打开 `Settings > Pages`。
3. `Build and deployment` 选择：
   - `Source`: `Deploy from a branch`
   - `Branch`: `main`（或你实际分支）/ `/ (root)`
4. 保存后等待 1-2 分钟，GitHub 会给出访问地址：
   - `https://<你的用户名>.github.io/<仓库名>/`

## 使用建议

- 尽量保证不同 sheet 的字段命名一致，这样配置更直观。
- 如果某些维度没有值，工具会把空值归类为 `(空值)`。
- 如果希望对齐口径，先配置第一个数据源，再点击“将第一个配置应用到全部已选”。
