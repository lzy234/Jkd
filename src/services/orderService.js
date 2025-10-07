import httpClient from '../api/httpClient';
import authService from './authService';
import logger from '../utils/logger';

class OrderService {
  constructor() {
    this.baseUrl = 'https://bapi.jkdsaas.com/b/order/list';
    this.updateMemoUrl = 'https://bapi.jkdsaas.com/b/order/set/buyer/memo';
    this.updateDispatcherUrl = 'https://bapi.jkdsaas.com/b/order/dispatch';
  }
  
  async getOrders(status = 'send', page = 1, limit = 20) {
    if (!authService.isAuthenticated()) {
      throw new Error('未登录，请先登录');
    }
    
    try {
      const params = {
        status,
        p: page,
        l: limit,
        start_time: '',
        end_time: '',
        s_type: 'tel',
        bs_uuid: '',
        sort_type: 'p_desc',
        k: '',
        ch_uuid: '',
        intf_token: '',
        fileuuid: '',
        bo_type: '',
        is_printed: '',
        order_type: 'wap'
      };
      
      const response = await httpClient.get(this.baseUrl, { params });
      
      if (response.data.code === 200) {
        logger.info(`成功获取${response.data.data.data.length}条订单`);
        return {
          success: true,
          data: response.data.data.data,
          total: response.data.data.total
        };
      }
      
      return {
        success: false,
        message: response.data.msg || '获取订单失败',
        data: []
      };
    } catch (error) {
      logger.error('获取订单失败:', error);
      throw error;
    }
  }
  
  async displayOrdersToExcel(orders) {
    if (!orders || orders.length === 0) {
      logger.warn('没有订单数据可展示');
      return;
    }
    
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // 清除现有内容
      const usedRange = sheet.getUsedRange(true);
      if (usedRange) {
        usedRange.clear();
      }
      
      // 设置表头
      const headers = [
        "订单ID", "联系人", "联系电话", "配送地址", 
        "商品名称", "派送人", "备注", "操作状态"
      ];
      
      const headerRange = sheet.getRange("A1:H1");
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#4472C4";
      headerRange.format.font.color = "#FFFFFF";
      headerRange.format.horizontalAlignment = "Center";
      
      // 填充订单数据
      const data = orders.map(order => {
        const goodsName = order.infos && order.infos.length > 0 
          ? order.infos[0].goods_name 
          : "";
          
        return [
          order.uuid || "",
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
        
        // 设置数据格式
        dataRange.format.horizontalAlignment = "Left";
        dataRange.format.verticalAlignment = "Center";
        
        // 设置派送人列的数据验证(下拉列表)
        await this.setDispatcherDropdown(context, sheet, data.length);
        
        // 设置可编辑单元格的样式
        const memoRange = sheet.getRange(`G2:G${data.length + 1}`);
        memoRange.format.fill.color = "#FFF2CC";  // 浅黄色背景
        
        const dispatcherRange = sheet.getRange(`F2:F${data.length + 1}`);
        dispatcherRange.format.fill.color = "#FFF2CC";  // 浅黄色背景
      }
      
      // 自动调整列宽
      sheet.getUsedRange().format.autofitColumns();
      
      // 创建表格对象便于筛选
      try {
        const tableRange = sheet.getRange(`A1:H${data.length + 1}`);
        const table = sheet.tables.add(tableRange, true);
        table.name = "OrdersTable";
        table.style = "TableStyleMedium2";
      } catch (e) {
        // 如果表格已存在，忽略错误
        logger.warn('表格创建失败，可能已存在:', e.message);
      }
      
      await context.sync();
      logger.info(`成功展示${orders.length}条订单到Excel`);
    });
  }
  
  async setDispatcherDropdown(context, sheet, rowCount) {
    // 派送人列表
    const dispatcherNames = ["浦西配送中心", "浦东配送中心", "马师傅", "莫师傅"];
    
    // 设置数据验证为下拉列表
    const dispatcherRange = sheet.getRange(`F2:F${rowCount + 1}`);
    const validation = dispatcherRange.dataValidation;
    validation.clear();
    validation.rule = {
      list: {
        inCellDropDown: true,
        source: dispatcherNames.join(",")
      }
    };
  }
  
  async updateOrderMemo(orderUuid, memo) {
    if (!authService.isAuthenticated()) {
      throw new Error('未登录，请先登录');
    }
    
    try {
      const formData = new URLSearchParams();
      formData.append('order_uuid', orderUuid);
      formData.append('memo', memo);
      
      const response = await httpClient.post(this.updateMemoUrl, formData, {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
        }
      });
      
      if (response.data.code === 200) {
        logger.info(`订单${orderUuid}备注更新成功`);
        return true;
      }
      
      logger.warn(`订单${orderUuid}备注更新失败:`, response.data.msg);
      return false;
    } catch (error) {
      logger.error('更新订单备注失败:', error);
      return false;
    }
  }
  
  async updateOrderDispatcher(orderUuid, dmName) {
    // 派送人ID映射
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
      
      const response = await httpClient.post(this.updateDispatcherUrl, formData, {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
        }
      });
      
      if (response.data.code === 200) {
        logger.info(`订单${orderUuid}派送人更新成功`);
        return true;
      }
      
      logger.warn(`订单${orderUuid}派送人更新失败:`, response.data.msg);
      return false;
    } catch (error) {
      logger.error('更新派送人失败:', error);
      return false;
    }
  }
}

export default new OrderService();

