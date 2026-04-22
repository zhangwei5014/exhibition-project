# EPMS - 展览项目管理系统

> 江苏移动全业务展厅项目进度管理与协作平台

## 功能特性

- ✅ 用户登录认证（多用户支持）
- ✅ 任务看板（按阶段分组、状态筛选）
- ✅ 风险预警（到期提醒、逾期警告）
- ✅ 施工日报管理（填写、查阅）
- ✅ Excel 导入/导出（兼容现有模板）
- ✅ 甘特图进度展示

## 技术栈

- **前端**: Streamlit + Plotly
- **后端**: Python + SQLite
- **部署**: Railway

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py --server.port 8501
```

## 部署到 Railway

1. Fork 此仓库
2. 登录 [Railway](https://railway.app)
3. New Project → Deploy from GitHub → 选择此仓库
4. Railway 自动安装依赖并部署

## 默认账号

- 用户名: `admin`
- 密码: `admin123`

## 项目结构

```
.
├── app.py              # 主程序
├── requirements.txt    # Python 依赖
├── railway.json        # Railway 部署配置
├── SPEC.md            # 需求规格说明书
└── epms.db            # SQLite 数据库（自动生成）
```

## 待部署功能

- [ ] 日报生成（导出 Word/PDF）
- [ ] 多项目管理
- [ ] 团队成员管理
- [ ] 邮件/消息通知
