import React from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';
import reportWebVitals from './reportWebVitals';
import AppLayout from './AppLayouts/AppLayout';
import { HashRouter, MemoryRouter } from 'react-router-dom';

const root = ReactDOM.createRoot(document.getElementById('root'));

// const Router = process.env.NODE_ENV === "development" ? BrowserRouter : HashRouter;
ReactDOM.createRoot(document.getElementById("root")).render(
  <MemoryRouter>
      <AppLayout />
  </MemoryRouter>
);

// If you want to start measuring performance in your app, pass a function
// to log results (for example: reportWebVitals(console.log))
// or send to an analytics endpoint. Learn more: https://bit.ly/CRA-vitals
reportWebVitals();
