import axios from 'axios';

// 创建axios实例
const httpClient = axios.create({
  timeout: 10000
});

// 请求拦截器
httpClient.interceptors.request.use(
  config => {
    // 添加认证头
    const token = sessionStorage.getItem('jkdAuthToken');
    if (token) {
      config.headers['Jkdauth'] = token;
    }
    
    // 添加其他通用头
    config.headers['User-Agent'] = 'Excel-Addin/1.0';
    config.headers['Accept'] = 'application/json, text/plain, */*';
    config.headers['Accept-Encoding'] = 'gzip, deflate, br';
    config.headers['Accept-Language'] = 'zh-CN';
    
    return config;
  },
  error => {
    return Promise.reject(error);
  }
);

// 响应拦截器
httpClient.interceptors.response.use(
  response => {
    return response;
  },
  error => {
    if (error.response && error.response.status === 401) {
      // 认证失败，清除token
      sessionStorage.removeItem('jkdAuthToken');
      // 触发重新登录事件
      window.dispatchEvent(new CustomEvent('authExpired'));
    }
    return Promise.reject(error);
  }
);

export default httpClient;

