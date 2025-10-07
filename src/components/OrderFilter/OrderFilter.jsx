import React, { useState } from 'react';
import orderService from '../../services/orderService';
import { showNotification } from '../../utils/errorHandler';
import { handleError } from '../../utils/errorHandler';
import './OrderFilter.css';

export default function OrderFilter() {
  const [status, setStatus] = useState('send');
  const [page, setPage] = useState(1);
  const [limit, setLimit] = useState(20);
  const [loading, setLoading] = useState(false);
  const [totalOrders, setTotalOrders] = useState(0);

  const handleGetOrders = async () => {
    setLoading(true);
    
    try {
      const result = await orderService.getOrders(status, page, limit);
      
      if (result.success) {
        setTotalOrders(result.total);
        await orderService.displayOrdersToExcel(result.data);
        showNotification(`成功获取${result.data.length}条订单`, 'success');
      } else {
        showNotification(result.message || '获取订单失败', 'error');
      }
    } catch (error) {
      const message = handleError(error);
      showNotification(message, 'error');
    } finally {
      setLoading(false);
    }
  };

  const handleRefresh = async () => {
    await handleGetOrders();
  };

  const handlePrevPage = () => {
    if (page > 1) {
      setPage(page - 1);
    }
  };

  const handleNextPage = () => {
    const maxPage = Math.ceil(totalOrders / limit);
    if (page < maxPage) {
      setPage(page + 1);
    }
  };

  return (
    <div className="order-filter-container">
      <div className="filter-header">
        <h3>订单筛选</h3>
      </div>
      
      <div className="filter-group">
        <label>订单状态</label>
        <select 
          value={status} 
          onChange={(e) => setStatus(e.target.value)}
          disabled={loading}
        >
          <option value="send">派送中</option>
          <option value="wait">待派送</option>
          <option value="arrive">已完成</option>
          <option value="cancel">已取消</option>
        </select>
      </div>
      
      <div className="filter-group">
        <label>每页显示</label>
        <select 
          value={limit} 
          onChange={(e) => setLimit(Number(e.target.value))}
          disabled={loading}
        >
          <option value="10">10条</option>
          <option value="20">20条</option>
          <option value="50">50条</option>
          <option value="100">100条</option>
        </select>
      </div>
      
      <div className="filter-actions">
        <button 
          className="btn-primary"
          onClick={handleGetOrders}
          disabled={loading}
        >
          {loading ? '加载中...' : '获取订单'}
        </button>
        
        <button 
          className="btn-secondary"
          onClick={handleRefresh}
          disabled={loading}
        >
          刷新数据
        </button>
      </div>
      
      <div className="pagination">
        <button 
          onClick={handlePrevPage} 
          disabled={page === 1 || loading}
        >
          上一页
        </button>
        
        <span className="page-info">
          第 {page} 页 {totalOrders > 0 && `/ 共 ${Math.ceil(totalOrders / limit)} 页`}
        </span>
        
        <button 
          onClick={handleNextPage} 
          disabled={page >= Math.ceil(totalOrders / limit) || loading}
        >
          下一页
        </button>
      </div>
      
      {totalOrders > 0 && (
        <div className="order-stats">
          共 {totalOrders} 条订单
        </div>
      )}
    </div>
  );
}

