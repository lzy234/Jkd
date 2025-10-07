import React, { useState } from 'react';
import authService from '../../services/authService';
import { showNotification } from '../../utils/errorHandler';
import './Login.css';

export default function Login({ onLoginSuccess }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [rememberMe, setRememberMe] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    
    if (!username || !password) {
      showNotification('请输入用户名和密码', 'warning');
      return;
    }
    
    setLoading(true);
    
    try {
      const result = await authService.login(username, password);
      
      if (result.success) {
        showNotification('登录成功！', 'success');
        if (rememberMe) {
          localStorage.setItem('savedUsername', username);
        }
        onLoginSuccess && onLoginSuccess(result.data);
      } else {
        showNotification(result.message || '登录失败，请检查用户名和密码', 'error');
      }
    } catch (error) {
      showNotification('登录失败，请检查网络连接', 'error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="login-container">
      <div className="login-header">
        <h2>订单管理系统</h2>
        <p>请登录您的账户</p>
      </div>
      
      <form className="login-form" onSubmit={handleLogin}>
        <div className="form-group">
          <label htmlFor="username">用户名</label>
          <input
            type="text"
            id="username"
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            placeholder="请输入用户名"
            disabled={loading}
          />
        </div>
        
        <div className="form-group">
          <label htmlFor="password">密码</label>
          <input
            type="password"
            id="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="请输入密码"
            disabled={loading}
          />
        </div>
        
        <div className="form-group checkbox-group">
          <label>
            <input
              type="checkbox"
              checked={rememberMe}
              onChange={(e) => setRememberMe(e.target.checked)}
              disabled={loading}
            />
            记住用户名
          </label>
        </div>
        
        <button 
          type="submit" 
          className="btn-login"
          disabled={loading}
        >
          {loading ? '登录中...' : '登录'}
        </button>
      </form>
    </div>
  );
}

