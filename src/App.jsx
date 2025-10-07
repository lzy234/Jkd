import React, { useState, useEffect } from 'react';
import Login from './components/Login/Login';
import OrderFilter from './components/OrderFilter/OrderFilter';
import authService from './services/authService';
import orderService from './services/orderService';
import { handleExcelError } from './utils/errorHandler';
import logger from './utils/logger';

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [userInfo, setUserInfo] = useState(null);

  useEffect(() => {
    // 检查是否已登录
    if (authService.isAuthenticated()) {
      setIsAuthenticated(true);
      setUserInfo(authService.getUserInfo());
    }

    // 监听认证过期事件
    window.addEventListener('authExpired', handleAuthExpired);

    // 注册Excel工作表变更事件
    registerExcelEvents();

    return () => {
      window.removeEventListener('authExpired', handleAuthExpired);
    };
  }, []);

  const handleAuthExpired = () => {
    setIsAuthenticated(false);
    setUserInfo(null);
  };

  const handleLoginSuccess = (data) => {
    setIsAuthenticated(true);
    setUserInfo(data);
  };

  const handleLogout = () => {
    authService.logout();
    setIsAuthenticated(false);
    setUserInfo(null);
  };

  const registerExcelEvents = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.onChanged.add(handleWorksheetChange);
        await context.sync();
        logger.info('Excel事件监听器注册成功');
      });
    } catch (error) {
      logger.error('注册Excel事件失败:', error);
    }
  };

  const handleWorksheetChange = async (event) => {
    return Excel.run(async (context) => {
      try {
        const changedRange = context.workbook.worksheets
          .getActiveWorksheet()
          .getRange(event.address);
        changedRange.load("values, rowIndex, columnIndex");
        
        await context.sync();
        
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // 备注列(G列, columnIndex为6)发生变化
        if (changedRange.columnIndex === 6 && changedRange.rowIndex > 0) {
          const orderIdCell = sheet.getCell(changedRange.rowIndex, 0);
          orderIdCell.load("values");
          const statusCell = sheet.getCell(changedRange.rowIndex, 7);
          
          await context.sync();
          
          const orderUuid = orderIdCell.values[0][0];
          const newMemo = changedRange.values[0][0] || '';
          
          if (orderUuid) {
            statusCell.values = [["更新中..."]];
            statusCell.format.font.color = "blue";
            await context.sync();
            
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
              logger.error('更新备注出错:', error);
            }
            
            await context.sync();
          }
        }
        // 派送人列(F列, columnIndex为5)发生变化
        else if (changedRange.columnIndex === 5 && changedRange.rowIndex > 0) {
          const orderIdCell = sheet.getCell(changedRange.rowIndex, 0);
          orderIdCell.load("values");
          const statusCell = sheet.getCell(changedRange.rowIndex, 7);
          
          await context.sync();
          
          const orderUuid = orderIdCell.values[0][0];
          const newDispatcher = changedRange.values[0][0] || '';
          
          if (orderUuid && newDispatcher) {
            statusCell.values = [["更新中..."]];
            statusCell.format.font.color = "blue";
            await context.sync();
            
            try {
              const success = await orderService.updateOrderDispatcher(orderUuid, newDispatcher);
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
              logger.error('更新派送人出错:', error);
            }
            
            await context.sync();
          }
        }
      } catch (error) {
        handleExcelError(error);
      }
    });
  };

  return (
    <div className="app-container">
      {!isAuthenticated ? (
        <Login onLoginSuccess={handleLoginSuccess} />
      ) : (
        <>
          <div className="app-header">
            <div className="user-info">
              <span>欢迎，{userInfo?.real_name || '用户'}</span>
              <button className="btn-logout" onClick={handleLogout}>
                退出登录
              </button>
            </div>
          </div>
          <OrderFilter />
        </>
      )}
    </div>
  );
}

