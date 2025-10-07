import logger from './logger';

// 全局错误处理函数
export function handleError(error) {
  logger.error('发生错误:', error);
  
  let message = '发生错误，请重试';
  
  if (error.message === '未登录，请先登录') {
    message = '登录已过期，请重新登录';
    window.dispatchEvent(new CustomEvent('authExpired'));
    return message;
  }
  
  if (error.response) {
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
    message = '无法连接到服务器，请检查网络';
  }
  
  return message;
}

// 在Excel任务窗格中显示通知
export function showNotification(message, type = 'error') {
  const notificationDiv = document.getElementById('notification');
  if (notificationDiv) {
    notificationDiv.textContent = message;
    notificationDiv.className = `notification ${type}`;
    notificationDiv.style.display = 'block';
    
    setTimeout(() => {
      notificationDiv.style.display = 'none';
    }, 3000);
  } else {
    console.log('通知:', message);
  }
}

// Excel中的错误处理
export function handleExcelError(error) {
  logger.error('Excel操作错误:', error);
  
  if (error instanceof OfficeExtension.Error) {
    logger.error('错误代码: ' + error.code);
    logger.error('错误消息: ' + error.message);
    logger.error('调试信息: ' + JSON.stringify(error.debugInfo));
  }
  
  return handleError(error);
}

