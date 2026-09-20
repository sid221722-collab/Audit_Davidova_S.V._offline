// Конфигурация автономной PWA.
// Поддерживаются оба имени переменной для совместимости с текущим index.html.
window.AUDIT_PWA_CONFIG = {
  CHECKER_NAME: '',
    APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbxGrRdeuETE-UnEgDboAVgqpUV0LfE0KOMcD38C6ZFCuRR84YU7uoEuMlfF_o9Hwx9CmA/exec'
};
window.AUDIT_APP_CONFIG = {
  GOOGLE_SCRIPT_URL: window.AUDIT_PWA_CONFIG.APPS_SCRIPT_URL
};
