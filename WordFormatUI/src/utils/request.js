import axios from 'axios';
import { backendBaseUrl } from './settings';

const service = axios.create({
  timeout: 120000, // 文档处理可能较慢，超时设为 2 分钟
  headers: {
    'Content-Type': 'application/json;charset=UTF-8',
  },
});

// 请求拦截器：动态注入 baseURL
service.interceptors.request.use(
  (config) => {
    config.baseURL = backendBaseUrl.value;

    // GET 请求添加时间戳防止缓存
    if (config.method === 'get') {
      config.params = {
        ...config.params,
        _t: Date.now(),
      };
    }
    return config;
  },
  (error) => Promise.reject(error),
);

// 响应拦截器
service.interceptors.response.use(
  (response) => {
    const res = response.data;

    if (res.code !== undefined && (res.code < 200 || res.code >= 300)) {
      return Promise.reject(new Error(res.msg || 'Error'));
    }

    if (res.success === false) {
      return Promise.reject(new Error(res.msg || '操作失败'));
    }

    return res;
  },
  (error) => {
    let message = '请求失败';

    if (error.response) {
      switch (error.response.status) {
        case 400: message = '请求错误'; break;
        case 404: message = '请求地址不存在'; break;
        case 500: message = '服务器内部错误'; break;
        default: message = `连接错误 ${error.response.status}`;
      }
    } else if (error.request) {
      message = '网络连接异常，请检查后端服务是否启动';
    } else {
      message = error.message;
    }

    return Promise.reject(new Error(message));
  },
);

class ApiClient {
  get(url, params = {}, config = {}) {
    return service({ method: 'get', url, params, ...config });
  }

  post(url, data = {}, config = {}) {
    return service({ method: 'post', url, data, ...config });
  }

  put(url, data = {}, config = {}) {
    return service({ method: 'put', url, data, ...config });
  }

  patch(url, data = {}, config = {}) {
    return service({ method: 'patch', url, data, ...config });
  }

  delete(url, params = {}, config = {}) {
    return service({ method: 'delete', url, params, ...config });
  }

  upload(url, formData, onProgress = null) {
    return service.post(url, formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
      onUploadProgress: (progressEvent) => {
        if (onProgress && progressEvent.total) {
          onProgress(Math.round((progressEvent.loaded * 100) / progressEvent.total));
        }
      },
    });
  }
}

const request = new ApiClient();
export default request;
