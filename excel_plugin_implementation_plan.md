# Excel加载项实现方案

## 1. 项目概述

本项目开发一个基于Office JS API的Excel加载项，实现与远程订单系统的交互功能，包括登录、获取订单数据和修改订单信息（派送人、备注）。该加载项可在Cursor编辑器中开发，无需Visual Studio环境。

## 2. 技术选型

### 2.1 Excel加载项开发技术

使用Office JS API开发Excel加载项，主要技术特点：

| 特性 | 描述 |
|------|------|
| 开发语言 | JavaScript/TypeScript |
| UI框架 | HTML/CSS/React |
| 开发环境 | Node.js + 任意代码编辑器(Cursor) |
| 构建工具 | Webpack/Babel |
| 运行环境 | Excel网页版、Windows版、Mac版 |
| 部署方式 | Office加载项(清单文件) |
| 跨平台 | 支持多种平台Excel |

**技术选择优势**：
- 无需Visual Studio，可在Cursor中完成所有开发
- 支持跨平台，包括Windows/Mac/Web版Excel
- 部署更简便，无需安装包
- 更新更便捷
- 使用现代Web技术栈，社区支持丰富

### 2.2 开发环境需求

- Node.js 14.x或更高版本
- npm 6.x或更高版本
- Cursor代码编辑器
- Office 2016或更高版本（或Excel网页版）
- Yeoman和Office加载项生成器(generator-office)

## 3. 加载项架构设计

### 3.1 项目结构

```
excel-order-addin/
├── .vscode/               # 编辑器配置
├── node_modules/          # npm依赖包
├── dist/                  # 构建输出目录
├── src/                   # 源代码
│   ├── api/               # API调用模块
│   │   ├── auth.js        # 认证API
│   │   ├── orders.js      # 订单API
│   │   └── httpClient.js  # HTTP客户端
│   ├── components/        # React组件
│   │   ├── Login/         # 登录组件
│   │   ├── OrderFilter/   # 订单筛选组件
│   │   └── Settings/      # 设置组件
│   ├── services/          # 业务逻辑
│   │   ├── authService.js # 认证服务
│   │   ├── orderService.js # 订单服务
│   │   └── syncService.js # 数据同步服务
│   ├── models/            # 数据模型
│   ├── utils/             # 工具函数
│   ├── taskpane.html      # 任务窗格HTML
│   ├── taskpane.css       # 任务窗格样式
│   └── taskpane.js        # 任务窗格逻辑
├── manifest.xml           # 加载项清单
├── package.json           # 项目依赖配置
├── webpack.config.js      # Webpack配置
└── README.md              # 项目说明
```

### 3.2 关键技术组件

1. **Office JS API**：与Excel交互的核心API
2. **React**：构建任务窗格UI
3. **Axios**：处理HTTP请求
4. **TypeScript**：类型支持和代码组织
5. **Webpack**：模块打包和构建
6. **localStorage/sessionStorage**：存储认证信息和配置

## 4. 用户界面设计

### 4.1 任务窗格设计

Excel加载项主要通过任务窗格提供用户界面：

1. **登录页面**
   - 用户名/密码输入框
   - 登录按钮
   - 登录状态显示
   - 记住登录信息选项

2. **订单操作页面**
   - 订单筛选选项（状态、页码）
   - 获取订单按钮
   - 刷新数据按钮
   - 订单统计信息

3. **设置页面**
   - API地址配置
   - 分页大小设置
   - 自动刷新间隔
   - 日志级别设置

### 4.2 Excel表格设计

订单数据在Excel表格中的展示格式：

```
| A列     | B列     | C列     | D列      | E列     | F列     | G列   | H列     |
|---------|---------|---------|----------|---------|---------|-------|---------|
| 订单ID  | 联系人  | 电话    | 配送地址  | 商品名称 | 派送人  | 备注  | 操作状态 |
| (只读)  | (只读)  | (只读)  | (只读)    | (只读)  | (下拉列表) | (可编辑) | (只读) |
```

- 使用Excel表格格式设置区分可编辑和只读单元格
- 派送人列设置为数据验证下拉列表
- 备注列允许自由编辑
- 操作状态列显示更新结果

### 4.3 Ribbon功能区按钮

在Excel功能区添加自定义按钮：
- "获取订单"按钮：打开任务窗格并加载订单数据
- "刷新数据"按钮：重新加载当前页订单数据
- "设置"按钮：打开设置界面

## 5. 功能实现

### 5.1 加载项初始化

```javascript
// taskpane.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // 初始化React应用
    ReactDOM.render(<App />, document.getElementById('container'));
    
    // 注册Excel事件监听器
    Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.onChanged.add(handleWorksheetChange);
      await context.sync();
    }).catch(handleError);
  }
});
```

### 5.2 认证服务实现

```javascript
// services/authService.js
export class AuthService {
  constructor() {
    this.token = sessionStorage.getItem('jkdAuthToken');
    this.baseUrl = 'https://bapi.jkdsaas.com/b/pc/login';
  }
  
  async login(username, password) {
    try {
      // 构造登录表单数据
      const formData = new URLSearchParams();
      formData.append('un', this.encodeCredential(username));
      formData.append('pw', this.encodeCredential(password));
      
      const response = await axios.post(this.baseUrl, formData, {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
        }
      });
      
      if (response.data.code === 200) {
        this.token = response.data.data.token;
        sessionStorage.setItem('jkdAuthToken', this.token);
        return true;
      }
      return false;
    } catch (error) {
      console.error('登录失败:', error);
      return false;
    }
  }
  
  isAuthenticated() {
    return !!this.token;
  }
  
  getAuthToken() {
    return this.token;
  }
  
  logout() {
    this.token = null;
    sessionStorage.removeItem('jkdAuthToken');
  }
  
  encodeCredential(input) {
    // 实现与原应用相同的凭据编码
    return btoa(input);
  }
}

export default new AuthService();
```

### 5.3 订单服务实现

```javascript
// services/orderService.js
import authService from './authService';
import httpClient from '../api/httpClient';

export class OrderService {
  constructor() {
    this.baseUrl = 'https://bapi.jkdsaas.com/b/order/list';
  }
  
  async getOrders(status = 'send', page = 1, limit = 20) {
    if (!authService.isAuthenticated()) {
      throw new Error('未登录，请先登录');
    }
    
    try {
      const response = await httpClient.get(
        `${this.baseUrl}?status=${status}&p=${page}&l=${limit}`
      );
      
      if (response.data.code === 200) {
        return response.data.data.data;
      }
      return [];
    } catch (error) {
      console.error('获取订单失败:', error);
      throw error;
    }
  }
  
  async displayOrdersToExcel(orders) {
    if (!orders || orders.length === 0) return;
    
    return Excel.run(async (context) => {
      // 获取当前工作表
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // 清除现有内容
      const usedRange = sheet.getUsedRange();
      usedRange.clear();
      
      // 设置表头
      const headers = [
        "订单ID", "联系人", "联系电话", "配送地址", 
        "商品名称", "派送人", "备注", "操作状态"
      ];
      
      const headerRange = sheet.getRange("A1:H1");
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#D3D3D3";
      
      // 填充订单数据
      const data = orders.map(order => {
        const goodsName = order.infos && order.infos.length > 0 
          ? order.infos[0].goods_name 
          : "";
          
        return [
          order.uuid,
          order.contact_info ? order.contact_info.contact : "",
          order.contact_info ? order.contact_info.phone : "",
          order.contact_info ? order.contact_info.address : "",
          goodsName,
          order.dm_name || "",
          order.memo || "",
          ""
        ];
      });
      
      // 写入数据
      if (data.length > 0) {
        const dataRange = sheet.getRange(`A2:H${data.length + 1}`);
        dataRange.values = data;
        
        // 设置派送人列的数据验证(下拉列表)
        // 这里需要获取派送人列表，暂时使用静态数据
        this.setDispatcherDropdown(context, sheet, data.length);
      }
      
      // 设置列宽自适应
      sheet.getUsedRange().format.autofitColumns();
      
      // 创建表格对象便于筛选
      const tableRange = sheet.getRange(`A1:H${data.length + 1}`);
      const table = sheet.tables.add(tableRange, true);
      table.name = "OrdersTable";
      
      // 保护除了可编辑单元格外的所有内容
      this.protectSheet(context, sheet, data.length);
      
      await context.sync();
    });
  }
  
  async setDispatcherDropdown(context, sheet, rowCount) {
    // 获取派送人列(F列)
    const dispatcherRange = sheet.getRange(`F2:F${rowCount + 1}`);
    
    // 设置数据验证为下拉列表
    const dispatcherNames = ["浦西配送中心", "浦东配送中心", "马师傅", "莫师傅"];
    const validation = dispatcherRange.dataValidation;
    validation.clear();
    validation.rule.list = { 
      inCellDropDown: true,
      source: dispatcherNames.join(",")
    };
  }
  
  async protectSheet(context, sheet, rowCount) {
    // 设置G列(备注)可编辑
    const memoRange = sheet.getRange(`G2:G${rowCount + 1}`);
    memoRange.format.fill.color = "#FFFFCC";  // 浅黄色背景标识可编辑单元格
    
    // 设置F列(派送人)可编辑
    const dispatcherRange = sheet.getRange(`F2:F${rowCount + 1}`);
    dispatcherRange.format.fill.color = "#FFFFCC";
  }
  
  async updateOrderMemo(orderUuid, memo) {
    if (!authService.isAuthenticated()) {
      throw new Error('未登录，请先登录');
    }
    
    try {
      const formData = new URLSearchParams();
      formData.append('order_uuid', orderUuid);
      formData.append('memo', memo);
      
      const response = await httpClient.post(
        'https://bapi.jkdsaas.com/b/order/set/buyer/memo',
        formData,
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
          }
        }
      );
      
      return response.data.code === 200;
    } catch (error) {
      console.error('更新订单备注失败:', error);
      return false;
    }
  }
  
  async updateOrderDispatcher(orderUuid, dmName) {
    // 获取派送人ID映射
    const dispatcherMap = {
      "浦西配送中心": "BSUDCMXM70TVZXAOHP8HVGMWIYPKXINK",
      "浦东配送中心": "BSUBUF32MMHXZLXY8U1RIYEUPJYD1NSM",
      "马师傅": "BSUV89NLHLH8ECSQOOQWQFLURXRGCH6O",
      "莫师傅": "BSUSVWNVWIFMVCZW5KW3RU2ZC6AEEC07"
    };
    
    const dmUuid = dispatcherMap[dmName];
    if (!dmUuid) {
      throw new Error('无效的派送人名称');
    }
    
    if (!authService.isAuthenticated()) {
      throw new Error('未登录，请先登录');
    }
    
    try {
      const formData = new URLSearchParams();
      formData.append('order_uuid', orderUuid);
      formData.append('dm_uuid', dmUuid);
      formData.append('dm_salary', '0');
      formData.append('re_flag', '1');
      
      const response = await httpClient.post(
        'https://bapi.jkdsaas.com/b/order/dispatch',
        formData,
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
          }
        }
      );
      
      return response.data.code === 200;
    } catch (error) {
      console.error('更新派送人失败:', error);
      return false;
    }
  }
}

export default new OrderService();
```

### 5.4 工作表变更监听实现

```javascript
// 工作表变更事件处理函数
async function handleWorksheetChange(event) {
  return Excel.run(async (context) => {
    // 获取变更单元格信息
    const changedRange = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(event.address);
    changedRange.load("values, rowIndex, columnIndex");
    
    await context.sync();
    
    // 获取表格中存储的订单信息
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // 如果是备注列(G列, columnIndex为6)发生变化
    if (changedRange.columnIndex === 6) {
      const orderIdCell = sheet.getCell(changedRange.rowIndex, 0); // A列订单ID
      orderIdCell.load("values");
      const statusCell = sheet.getCell(changedRange.rowIndex, 7); // H列状态单元格
      
      await context.sync();
      
      const orderUuid = orderIdCell.values[0][0];
      const newMemo = changedRange.values[0][0];
      
      if (orderUuid) {
        // 更新状态单元格为"更新中..."
        statusCell.values = [["更新中..."]];
        statusCell.format.font.color = "blue";
        await context.sync();
        
        // 调用API更新备注
        try {
          const success = await orderService.updateOrderMemo(orderUuid, newMemo);
          if (success) {
            statusCell.values = [["更新成功"]];
            statusCell.format.font.color = "green";
          } else {
            statusCell.values = [["更新失败"]];
            statusCell.format.font.color = "red";
          }
        } catch (error) {
          statusCell.values = [["更新失败"]];
          statusCell.format.font.color = "red";
          console.error('更新备注出错:', error);
        }
        
        await context.sync();
      }
    }
    // 类似地处理派送人列(F列, columnIndex为5)变化
    else if (changedRange.columnIndex === 5) {
      // 实现与备注列相似的逻辑，调用updateOrderDispatcher
    }
  }).catch(handleError);
}
```

### 5.5 HTTP客户端实现

```javascript
// api/httpClient.js
import axios from 'axios';
import authService from '../services/authService';

// 创建axios实例
const httpClient = axios.create({
  timeout: 10000
});

// 请求拦截器
httpClient.interceptors.request.use(
  config => {
    // 添加认证头
    const token = authService.getAuthToken();
    if (token) {
      config.headers['Jkdauth'] = token;
    }
    
    // 添加其他通用头
    config.headers['User-Agent'] = 'Excel-Addin/1.0';
    
    return config;
  },
  error => {
    return Promise.reject(error);
  }
);

// 响应拦截器
httpClient.interceptors.response.use(
  response => {
    // 处理成功响应
    return response;
  },
  error => {
    // 处理错误响应
    if (error.response && error.response.status === 401) {
      // 认证失败，清除token
      authService.logout();
      // 可以在这里触发重新登录
    }
    return Promise.reject(error);
  }
);

export default httpClient;
```

## 6. 数据模型定义

使用TypeScript接口定义数据模型：

```typescript
// models/types.ts

// 登录响应模型
export interface LoginResponse {
  code: number;
  data: LoginData;
  msg: string;
}

export interface LoginData {
  token: string;
  user_uuid: string;
  buser_uuid: string;
  // 其他字段...
}

// 订单列表响应模型
export interface OrderListResponse {
  code: number;
  data: OrderListData;
}

export interface OrderListData {
  data: Order[];
  total: number;
}

export interface Order {
  uuid: string;
  contact_info?: ContactInfo;
  bo_status: string;
  memo?: string;
  dm_name?: string;
  dm_buser_uuid?: string;
  infos?: OrderItem[];
  // 其他字段...
}

export interface ContactInfo {
  contact?: string;
  phone?: string;
  address?: string;
  // 其他字段...
}

export interface OrderItem {
  goods_name: string;
  name?: string;
  num: number;
  price: number;
  // 其他字段...
}

// 派送人模型
export interface Dispatcher {
  name: string;
  uuid: string;
}

// 基础响应模型
export interface BasicResponse {
  code: number;
  msg: string;
}
```

## 7. 加载项部署

### 7.1 创建加载项清单文件

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xsi:type="TaskPaneApp">

  <Id>12345678-1234-1234-1234-123456789012</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Your Company</ProviderName>
  <DefaultLocale>zh-CN</DefaultLocale>
  <DisplayName DefaultValue="订单管理插件" />
  <Description DefaultValue="Excel订单管理插件，用于获取和修改订单数据"/>

  <IconUrl DefaultValue="https://your-cdn.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://your-cdn.com/assets/icon-80.png"/>

  <SupportUrl DefaultValue="https://your-support-url.com" />

  <Hosts>
    <Host Name="Workbook" />
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html" />
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16" />
                  <bt:Image size="32" resid="Icon.32x32" />
                  <bt:Image size="80" resid="Icon.80x80" />
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://your-cdn.com/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://your-cdn.com/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://your-cdn.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://your-learn-more-url.com"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="开始使用订单管理插件"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="订单管理"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="打开订单管理"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="订单管理插件已成功加载。点击"订单管理"选项卡上的"打开订单管理"按钮开始。"/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="点击打开订单管理插件"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

### 7.2 部署选项

1. **开发测试**：
   - 使用`npm start`启动本地服务器
   - 在Excel中手动添加加载项（开发者加载项）

2. **共享文件夹部署**：
   - 构建生产版本：`npm run build`
   - 将构建后的文件放在网络共享文件夹或Web服务器上
   - 更新manifest.xml中的URLs指向实际部署位置
   - 将manifest.xml添加到Excel的受信任加载项目录

3. **集中部署** (Office 365管理员)：
   - 使用Office 365管理中心部署给组织内用户
   - 需要上传manifest.xml和托管好的Web资源

4. **AppSource商店发布** (可选)：
   - 需要准备符合Microsoft商店要求的加载项
   - 提交审核和发布

## 8. 异常处理与日志记录

```javascript
// utils/logger.js
export const logLevels = {
  ERROR: 0,
  WARN: 1,
  INFO: 2,
  DEBUG: 3
};

class Logger {
  constructor() {
    // 从localStorage获取日志级别，默认为INFO
    this.level = parseInt(localStorage.getItem('logLevel') || logLevels.INFO);
  }
  
  setLevel(level) {
    this.level = level;
    localStorage.setItem('logLevel', level);
  }
  
  debug(message, ...args) {
    if (this.level >= logLevels.DEBUG) {
      console.debug(`[DEBUG] ${message}`, ...args);
      this._saveLog('DEBUG', message, args);
    }
  }
  
  info(message, ...args) {
    if (this.level >= logLevels.INFO) {
      console.info(`[INFO] ${message}`, ...args);
      this._saveLog('INFO', message, args);
    }
  }
  
  warn(message, ...args) {
    if (this.level >= logLevels.WARN) {
      console.warn(`[WARN] ${message}`, ...args);
      this._saveLog('WARN', message, args);
    }
  }
  
  error(message, error, ...args) {
    if (this.level >= logLevels.ERROR) {
      console.error(`[ERROR] ${message}`, error, ...args);
      this._saveLog('ERROR', message, [error, ...args]);
    }
  }
  
  _saveLog(level, message, args) {
    // 保存日志到localStorage，限制大小
    try {
      const logs = JSON.parse(localStorage.getItem('appLogs') || '[]');
      logs.push({
        time: new Date().toISOString(),
        level,
        message,
        details: args.map(arg => {
          if (arg instanceof Error) {
            return {message: arg.message, stack: arg.stack};
          }
          return String(arg);
        })
      });
      
      // 保留最近100条日志
      while (logs.length > 100) {
        logs.shift();
      }
      
      localStorage.setItem('appLogs', JSON.stringify(logs));
    } catch (e) {
      console.error('保存日志失败:', e);
    }
  }
  
  getLogs() {
    return JSON.parse(localStorage.getItem('appLogs') || '[]');
  }
  
  clearLogs() {
    localStorage.removeItem('appLogs');
  }
}

export default new Logger();
```

## 9. 错误处理

```javascript
// utils/errorHandler.js
import logger from './logger';

// 全局错误处理函数
export function handleError(error) {
  logger.error('发生错误:', error);
  
  // 根据错误类型显示不同的错误提示
  let message = '发生错误，请重试';
  
  if (error.message === '未登录，请先登录') {
    message = '登录已过期，请重新登录';
    // 可以在这里触发重新登录
    return;
  }
  
  if (error.response) {
    // HTTP错误响应
    switch (error.response.status) {
      case 400:
        message = '请求参数错误';
        break;
      case 401:
        message = '未授权，请重新登录';
        break;
      case 403:
        message = '您没有权限执行此操作';
        break;
      case 404:
        message = '请求的资源不存在';
        break;
      case 500:
        message = '服务器内部错误';
        break;
      default:
        message = `服务器响应错误(${error.response.status})`;
    }
  } else if (error.request) {
    // 发送请求但没有收到响应
    message = '无法连接到服务器，请检查网络';
  }
  
  // 在Excel中显示错误消息
  showNotification(message);
}

// 在Excel任务窗格中显示通知
export function showNotification(message, type = 'error') {
  // 如果使用React，可以通过状态更新显示通知
  // 这里简单实现，实际中应该使用更好的UI组件
  const notificationDiv = document.getElementById('notification');
  if (notificationDiv) {
    notificationDiv.textContent = message;
    notificationDiv.className = `notification ${type}`;
    notificationDiv.style.display = 'block';
    
    // 3秒后自动隐藏
    setTimeout(() => {
      notificationDiv.style.display = 'none';
    }, 3000);
  } else {
    console.log('通知:', message);
  }
}
```

## 10. 总结

这个基于Office JS API的Excel加载项方案完全可以在Cursor编辑器中开发，无需使用Visual Studio。它提供了与VSTO方案相同的核心功能，但具有更好的跨平台兼容性和更简便的部署方式。使用现代Web技术栈，开发人员可以更轻松地构建和维护这个加载项。