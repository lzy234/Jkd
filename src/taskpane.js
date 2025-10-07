import React from 'react';
import ReactDOM from 'react-dom';
import App from './App';
import './taskpane.css';

// Office.js初始化
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // 渲染React应用
    ReactDOM.render(<App />, document.getElementById('container'));
  }
});

