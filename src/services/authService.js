import httpClient from '../api/httpClient';

class AuthService {
  constructor() {
    this.baseUrl = 'https://bapi.jkdsaas.com/b/pc/login';
    this.token = sessionStorage.getItem('jkdAuthToken');
  }
  
  async login(username, password) {
    try {
      // 构造登录表单数据
      const formData = new URLSearchParams();
      formData.append('un', this.encodeCredential(username));
      formData.append('pw', this.encodeCredential(password));
      
      const response = await httpClient.post(this.baseUrl, formData, {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
          'Jkdauth': '' // 登录时Jkdauth为空
        }
      });
      
      if (response.data.code === 200) {
        this.token = response.data.data.token;
        sessionStorage.setItem('jkdAuthToken', this.token);
        sessionStorage.setItem('userInfo', JSON.stringify(response.data.data));
        return {
          success: true,
          data: response.data.data
        };
      }
      
      return {
        success: false,
        message: response.data.msg || '登录失败'
      };
    } catch (error) {
      console.error('登录失败:', error);
      return {
        success: false,
        message: error.message || '登录失败，请检查网络连接'
      };
    }
  }
  
  isAuthenticated() {
    return !!this.token;
  }
  
  getAuthToken() {
    return this.token || sessionStorage.getItem('jkdAuthToken');
  }
  
  getUserInfo() {
    const userInfo = sessionStorage.getItem('userInfo');
    return userInfo ? JSON.parse(userInfo) : null;
  }
  
  logout() {
    this.token = null;
    sessionStorage.removeItem('jkdAuthToken');
    sessionStorage.removeItem('userInfo');
  }
  
  encodeCredential(input) {
    // 实现与原应用相同的凭据编码
    // 这里需要根据api.md中的实际加密方式实现
    // 暂时使用base64编码，实际可能需要RSA加密等
    return btoa(unescape(encodeURIComponent(input)));
  }
}

export default new AuthService();

