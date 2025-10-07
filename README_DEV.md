# Excel订单管理插件

基于Office JS API开发的Excel加载项，用于订单管理系统的数据获取和修改。

## 功能特性

- ✅ 用户登录认证
- ✅ 获取订单列表（支持筛选和分页）
- ✅ 订单数据展示在Excel表格中
- ✅ 直接在Excel中修改订单备注
- ✅ 直接在Excel中修改派送人
- ✅ 自动同步修改到服务器

## 技术栈

- React 18
- Office JS API
- Axios
- Webpack 5
- Babel

## 开发环境要求

- Node.js 14+
- npm 6+
- Excel 2016+ 或 Excel网页版

## 安装步骤

1. 安装依赖：
```bash
npm install
```

2. 启动开发服务器：
```bash
npm start
```

3. 在Excel中加载插件：
   - 打开Excel
   - 文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 受信任的加载项目录
   - 添加项目根目录路径
   - 在Excel中：插入 -> 我的加载项 -> 共享文件夹 -> 选择"订单管理插件"

## 开发命令

```bash
# 启动开发服务器（带热重载）
npm start

# 构建生产版本
npm run build

# 运行代码检查
npm run lint

# 运行测试
npm test
```

## 项目结构

```
excel-order-addin/
├── src/
│   ├── api/              # API调用模块
│   ├── components/       # React组件
│   ├── services/         # 业务服务
│   ├── utils/           # 工具函数
│   ├── App.jsx          # 主应用组件
│   ├── taskpane.html    # 任务窗格HTML
│   └── taskpane.js      # 任务窗格入口
├── dist/                # 构建输出
├── manifest.xml         # 加载项清单
├── package.json         # 项目配置
└── webpack.config.js    # Webpack配置
```

## 使用说明

1. 点击Excel功能区的"订单管理"按钮打开任务窗格
2. 输入用户名和密码登录
3. 选择订单状态和分页参数
4. 点击"获取订单"按钮加载数据到Excel
5. 直接在Excel中修改备注或派送人
6. 修改会自动同步到服务器

## 注意事项

- 需要HTTPS环境运行（开发环境使用自签名证书）
- 首次运行需要信任自签名证书
- 修改数据时需要保持网络连接
- 关闭Excel前建议先退出登录

## 故障排除

1. **加载项无法加载**
   - 检查manifest.xml中的URL是否正确
   - 确认开发服务器是否运行
   - 检查是否信任了加载项目录

2. **无法登录**
   - 检查网络连接
   - 确认用户名密码是否正确
   - 查看浏览器控制台错误信息

3. **数据不更新**
   - 检查是否已登录
   - 确认网络连接正常
   - 查看任务窗格中的错误提示

## 开发者

开发于 Cursor 编辑器

